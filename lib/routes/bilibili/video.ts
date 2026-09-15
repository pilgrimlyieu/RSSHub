import type { Context } from 'hono';

import { config } from '@/config';
import type { Route } from '@/types';
import { ViewType } from '@/types';
import got from '@/utils/got';
import { parseDuration } from '@/utils/helpers';
import logger from '@/utils/logger';
import { getPlaywrightPage, type Page } from '@/utils/playwright';

import cache from './cache';
import utils, { getVideoUrl } from './utils';

export const route: Route = {
    path: '/user/video/:uid/:embed?',
    categories: ['social-media'],
    view: ViewType.Videos,
    example: '/bilibili/user/video/2267573',
    parameters: { uid: '用户 id, 可在 UP 主主页中找到', embed: '默认为开启内嵌视频, 任意值为关闭' },
    features: {
        requireConfig: false,
        requirePuppeteer: false,
        antiCrawler: false,
        supportBT: false,
        supportPodcast: false,
        supportScihub: false,
    },
    radar: [
        {
            source: ['space.bilibili.com/:uid'],
            target: '/user/video/:uid',
        },
    ],
    name: 'UP 主投稿',
    maintainers: ['DIYgod', 'Konano', 'pseudoyu'],
    handler,
};

interface VideoItem {
    aid: number;
    author?: string;
    bvid?: string;
    comment?: number;
    created: number;
    description: string;
    length: string;
    pic: string;
    title: string;
}

interface VideoListData {
    list?: {
        vlist?: VideoItem[];
    };
    v_voucher?: string;
}

interface VideoListResponse {
    code?: number;
    data?: VideoListData;
    message?: string;
}

type BrowserResponse = Awaited<ReturnType<Page['waitForResponse']>>;

const videoListApiPath = '/x/space/wbi/arc/search';
const allowedBrowserRequestTypes = new Set(['document', 'script', 'xhr', 'fetch', 'image', 'font', 'stylesheet', 'other']);
const browserResponseTimeout = 45000;
const browserCloseTimeout = 90000;

class BilibiliRiskControlError extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'BilibiliRiskControlError';
    }
}

const getErrorMessage = (cause: unknown) => (cause instanceof Error ? cause.message : String(cause));

const getApiErrorMessage = (code: number | undefined, message: string | undefined, mode = 'fetching', data?: VideoListData) =>
    `Got error code ${code} while ${mode}: ${message}${data?.v_voucher ? ' (v_voucher present; captcha/risk verification required)' : ''}`;

const hasRiskControlVoucher = (data: VideoListResponse) => data.code === -352 && !!data.data?.v_voucher;

const isVideoListApiResponse = (response: BrowserResponse) => {
    const request = response.request();
    return response.url().includes(videoListApiPath) && request.method() === 'GET' && (request.resourceType() === 'xhr' || request.resourceType() === 'fetch');
};

const waitForVideoListResponse = async (page: Page) => {
    try {
        return {
            response: await page.waitForResponse(isVideoListApiResponse, { timeout: browserResponseTimeout }),
        };
    } catch (error) {
        return { error };
    }
};

const navigateToVideoPage = async (page: Page, videoUrl: string) => {
    try {
        await page.goto(videoUrl, { timeout: browserResponseTimeout, waitUntil: 'domcontentloaded' });
    } catch (error) {
        logger.warn(`[bilibili/video] video page navigation did not finish before the response wait ended: ${getErrorMessage(error)}`);
    }
};

const getVideoListResponse = async (responsePromise: ReturnType<typeof waitForVideoListResponse>) => {
    const videoListResponseResult = await responsePromise;
    if ('error' in videoListResponseResult) {
        throw new Error(`Bilibili browser mode did not receive a video list response within ${browserResponseTimeout}ms: ${getErrorMessage(videoListResponseResult.error)}`);
    }

    return videoListResponseResult.response;
};

const waitForVideoListResponseFromVideoPage = async (page: Page, videoUrl: string): Promise<BrowserResponse> => {
    const videoListResponsePromise = waitForVideoListResponse(page);
    await navigateToVideoPage(page, videoUrl);

    const response = await getVideoListResponse(videoListResponsePromise);

    if (response.status() !== 200) {
        throw new Error(`Bilibili browser mode returned unexpected video list API status ${response.status()}`);
    }

    const contentType = response.headers()['content-type'];
    if (!contentType?.includes('application/json')) {
        throw new Error(`Bilibili browser mode returned non-JSON response with status ${response.status()}; BILIBILI_COOKIE_* may be required`);
    }

    return response;
};

async function applyCookie(page: Page, cookie: string) {
    const cookies = cookie
        .split(';')
        .map((item) => item.trim())
        .filter(Boolean)
        .map((item) => {
            const equalIndex = item.indexOf('=');
            if (equalIndex <= 0) {
                return;
            }

            return {
                name: item.slice(0, equalIndex).trim(),
                value: item.slice(equalIndex + 1).trim(),
                domain: '.bilibili.com',
                path: '/',
            };
        })
        .filter((item) => item !== undefined);

    if (cookies.length > 0) {
        await page.context().addCookies(cookies);
    }
}

async function fetchVideoListFromApi(uid: string): Promise<VideoListData> {
    const cookie = await cache.getCookie();
    const wbiVerifyString = await cache.getWbiVerifyString();
    const dmImgList = utils.getDmImgList();
    const dmImgInter = utils.getDmImgInter();
    const renderData = await cache.getRenderData(uid);

    const params = utils.addWbiVerifyInfo(
        utils.addRenderData(utils.addDmVerifyInfoWithInter(`mid=${uid}&ps=30&tid=0&pn=1&keyword=&order=pubdate&platform=web&web_location=1550101&order_avoided=true`, dmImgList, dmImgInter), renderData),
        wbiVerifyString
    );
    const response = await got(`https://api.bilibili.com${videoListApiPath}?${params}`, {
        headers: {
            Referer: `https://space.bilibili.com/${uid}`,
            origin: 'https://space.bilibili.com',
            Cookie: cookie,
        },
    });
    const data = response.data;
    if (data.code) {
        logger.error(JSON.stringify(data.data));
        const message = getApiErrorMessage(data.code, data.message, 'fetching', data.data);
        if (hasRiskControlVoucher(data)) {
            throw new BilibiliRiskControlError(message);
        }

        throw new Error(message);
    }

    return data.data;
}

async function fetchVideoListFromBrowser(uid: string): Promise<VideoListData> {
    const cookie = cache.getConfiguredCookie();
    const videoUrl = `https://space.bilibili.com/${uid}/video`;
    logger.info(`[bilibili/video] fetching via playwright: ${videoUrl}`);

    const { destroy, page } = await getPlaywrightPage(videoUrl, {
        closeTimeout: browserCloseTimeout,
        noGoto: true,
        onBeforeLoad: async (page) => {
            if (cookie) {
                await applyCookie(page, cookie);
            }

            await page.route('**/*', (route) => {
                const request = route.request();
                allowedBrowserRequestTypes.has(request.resourceType()) ? route.continue() : route.abort();
            });
        },
        gotoConfig: { waitUntil: 'domcontentloaded' },
    });

    try {
        const response = await waitForVideoListResponseFromVideoPage(page, videoUrl);
        const data = (await response.json()) as VideoListResponse;
        if (data.code) {
            logger.error(JSON.stringify(data.data));
            throw new Error(getApiErrorMessage(data.code, data.message, 'fetching in browser mode', data.data));
        }

        if (!data.data) {
            throw new Error('Bilibili browser response does not contain video list data');
        }

        return data.data;
    } finally {
        await destroy();
    }
}

async function getVideoList(uid: string): Promise<VideoListData> {
    try {
        return await fetchVideoListFromApi(uid);
    } catch (error) {
        if (error instanceof BilibiliRiskControlError) {
            throw error;
        }

        logger.warn(`[bilibili/video] API request failed, falling back to browser mode: ${error}`);
        return fetchVideoListFromBrowser(uid);
    }
}

async function handler(ctx: Context) {
    const isJsonFeed = ctx.req.query('format') === 'json';

    const uid = ctx.req.param('uid');
    const embed = !ctx.req.param('embed');
    const data = await getVideoList(uid!);
    const videos = data.list?.vlist ?? [];

    let name = videos[0]?.author || uid;
    let face: string | undefined;

    try {
        const usernameAndFace = await cache.getUsernameAndFaceFromUID(uid);
        name = usernameAndFace[0] || name;
        face = usernameAndFace[1] ?? undefined;
    } catch (error) {
        logger.warn(`[bilibili/video] failed to fetch user profile: ${error}`);
    }

    return {
        title: `${name} 的 bilibili 空间`,
        link: `https://space.bilibili.com/${uid}`,
        description: `${name} 的 bilibili 空间`,
        image: face ?? undefined,
        logo: face ?? undefined,
        icon: face ?? undefined,
        item: await Promise.all(
            videos.map(async (item) => {
                const subtitles = isJsonFeed && !config.bilibili.excludeSubtitles && item.bvid ? await cache.getVideoSubtitleAttachment(item.bvid) : [];
                return {
                    title: item.title,
                    description: utils.renderUGCDescription(embed, item.pic, item.description, String(item.aid), undefined, item.bvid),
                    pubDate: new Date(item.created * 1000).toUTCString(),
                    link: item.created > utils.bvidTime && item.bvid ? `https://www.bilibili.com/video/${item.bvid}` : `https://www.bilibili.com/video/av${item.aid}`,
                    author: name,
                    comments: item.comment,
                    attachments: item.bvid
                        ? [
                              {
                                  url: getVideoUrl(item.bvid),
                                  mime_type: 'text/html',
                                  duration_in_seconds: parseDuration(item.length),
                              },
                              ...subtitles,
                          ]
                        : undefined,
                };
            })
        ),
    };
}
