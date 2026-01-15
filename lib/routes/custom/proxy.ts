import { promisify } from 'node:util';
import zlib from 'node:zlib';

import RSSParser from 'rss-parser';

import type { Data, Route } from '@/types';
import got from '@/utils/got';
import logger from '@/utils/logger';

const gunzip = promisify(zlib.gunzip);
const inflate = promisify(zlib.inflate);
const brotliDecompress = promisify(zlib.brotliDecompress);

export const route: Route = {
    path: '/proxy',
    categories: ['other'],
    example: '/custom/proxy?url=https://rsshub.app/sspai/index',
    parameters: {
        url: 'RSS feed URL to proxy',
    },
    features: {
        requireConfig: false,
        requirePuppeteer: false,
        antiCrawler: false,
        supportBT: false,
        supportPodcast: false,
        supportScihub: false,
    },
    name: 'RSS Proxy',
    maintainers: [],
    handler,
    description: 'RSS Proxy',
};

async function handler(ctx): Promise<Data> {
    const url = ctx.req.query('url');

    if (!url) {
        throw new Error('URL parameter is required');
    }

    // 1. 实例化一个新的 Parser，配置 customFields 以强制捕获 Atom 的 category 标签
    const parser = new RSSParser({
        customFields: {
            item: [
                ['category', 'categories', { keepArray: true }],
                ['content', 'content', { keepArray: false }],
                ['summary', 'summary', { keepArray: false }],
            ],
        },
    });

    // 2. 发起请求
    const response = await got({
        method: 'get',
        url,
        responseType: 'buffer',
        headers: {
            'Accept-Encoding': 'gzip, deflate, br',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 RSSHub/1.0',
        },
    });

    // 3. 数据获取与防御性处理
    let data: Buffer;
    // @ts-ignore
    const rawData = response.data || response.body || response;

    if (Buffer.isBuffer(rawData)) {
        data = rawData;
    } else if (typeof rawData === 'string') {
        data = Buffer.from(rawData);
    } else {
        data = Buffer.from('');
    }

    // 4. 安全读取 Headers
    // @ts-ignore
    const headers = response.headers || {};
    const contentEncoding = (headers['content-encoding'] || '').toLowerCase();

    // 5. 智能解压
    try {
        if (data.length > 2 && data[0] === 0x1f && data[1] === 0x8b) {
            data = await gunzip(data);
        } else if (data.length > 2 && data[0] === 0x78) {
            data = await inflate(data);
        } else if (contentEncoding === 'br') {
            data = await brotliDecompress(data);
        }
    } catch (error) {
        logger.error(`Decompression error for ${url}:`, error);
    }

    // 6. 解析 XML
    const contentString = data.toString('utf-8');
    if (!contentString.trim()) {
        throw new Error('Empty response body');
    }

    const feed = await parser.parseString(contentString);

    // 7. 映射数据
    const items = (feed.items || []).map((item: any) => {
        const enclosure = item.enclosure;

        let categories: string[] = [];
        if (Array.isArray(item.categories)) {
            categories = item.categories
                .map((c: any) => {
                    // 情况 1：RSS 标准 <category>Text</category> -> 解析为字符串 "Text"
                    if (typeof c === 'string') {
                        return c;
                    }
                    // 情况 2：Atom 标准 <category term="Text" /> -> 解析为 { $: { term: "Text" } }
                    if (c?.$?.term) {
                        return c.$.term;
                    }
                    // 兜底：尝试读取 name 或 label
                    if (c?.name) {
                        return c.name;
                    }
                    if (c?.label) {
                        return c.label;
                    }
                    if (c?._) {
                        return c._;
                    }
                    return null;
                })
                .filter((c) => c && typeof c === 'string');
        }

        let authorName = '';
        if (typeof item.creator === 'string') {
            authorName = item.creator;
        } else if (typeof item.author === 'string') {
            authorName = item.author;
        } else if (item.author && item.author.name) {
            // 处理 Atom 的 author 结构
            authorName = Array.isArray(item.author.name) ? item.author.name[0] : item.author.name;
        }

        return {
            title: item.title ?? item.link ?? 'Untitled',
            // 优先顺序：content -> summary (Atom) -> description (RSS)
            description: item.content || item.summary || item.description || '',
            link: item.link,
            pubDate: item.isoDate ?? item.pubDate,
            author: authorName,
            category: categories,
            guid: item.guid || item.id || item.link,
            ...(enclosure?.url && {
                enclosure_url: enclosure.url,
                enclosure_type: enclosure.type,
                enclosure_length: enclosure.length ? Number(enclosure.length) : undefined,
            }),
        };
    });

    return {
        title: feed.title ?? 'RSS Proxy Feed',
        description: feed.description ?? '',
        link: feed.link ?? url,
        image: feed.image?.url,
        language: feed.language,
        item: items,
        ...(feed.itunes?.author && { itunes_author: feed.itunes.author }),
        ...(feed.itunes?.categories?.[0] && { itunes_category: feed.itunes.categories[0] }),
    };
}
