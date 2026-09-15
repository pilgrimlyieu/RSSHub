import { Hono } from 'hono';
import Parser from 'rss-parser';
import undici from 'undici';
import { describe, expect, it, vi } from 'vitest';

import app from '@/app';
import { config } from '@/config';
import { route as hackernewsRoute } from '@/routes/hackernews/index';

describe('index', () => {
    it('exports app entrypoint', () => {
        expect(app.request).toBeInstanceOf(Function);
    });

    it('serve index', async () => {
        const res = await app.request('/');
        expect(res.status).toBe(200);
        expect(await res.text()).toContain('Welcome to RSSHub!');
    });
});

describe('request-rewriter', () => {
    it('should rewrite request', async () => {
        const fetchSpy = vi.spyOn(undici, 'fetch');
        await app.request('/test/httperror');

        // headers
        const request = fetchSpy.mock.lastCall?.[0] as Request | undefined;
        expect(request?.headers.get('user-agent')).toMatch(/Chrome/);
    });
});

const parser = new Parser();

describe('Hacker News best stories', () => {
    it('keeps unavailable and deleted stories out of the feed when details are cached', async () => {
        const fetchSpy = vi.spyOn(undici, 'fetch').mockImplementation((input) => {
            const url = String(input instanceof undici.Request ? input.url : input);
            if (url.endsWith('/beststories.json')) {
                return Promise.resolve(undici.Response.json([91101, 91102, 91103]));
            }
            if (url.endsWith('/91101.json')) {
                return Promise.resolve(undici.Response.json(null));
            }
            if (url.endsWith('/91102.json')) {
                return Promise.resolve(undici.Response.json({ id: 91102, deleted: true }));
            }
            return Promise.resolve(undici.Response.json({ id: 91103, title: 'A surviving story' }));
        });
        fetchSpy.mockClear();
        const testApp = new Hono();
        testApp.get('/hackernews/:section', async (ctx) => ctx.json(await hackernewsRoute.handler(ctx)));

        try {
            const first = await testApp.request('/hackernews/best?limit=3');
            const second = await testApp.request('/hackernews/best?limit=3');
            const firstData = await first.json();
            const secondData = await second.json();
            expect(firstData.item).toHaveLength(1);
            expect(secondData.item).toEqual(firstData.item);
            expect(secondData.item[0].pubDate).toBeUndefined();
            expect(fetchSpy).toHaveBeenCalledTimes(5);
        } finally {
            fetchSpy.mockRestore();
        }
    });

    it('uses the official API when the website rejects requests and preserves story order and links', async () => {
        const requests: string[] = [];
        const stories = {
            91001: { id: 91001, type: 'story', title: 'An external story', url: 'https://example.com/article', by: 'alice', time: 1_790_812_800, descendants: 12, score: 120 },
            91002: { id: 91002, type: 'story', title: 'Ask HN: A discussion', text: '<p>Discussion body</p>', by: 'bob', time: 1_790_812_900, descendants: 3, score: 30 },
        };
        const fetchSpy = vi.spyOn(undici, 'fetch').mockImplementation((input) => {
            const url = String(input instanceof undici.Request ? input.url : input);
            requests.push(url);
            if (url === 'https://hacker-news.firebaseio.com/v0/beststories.json') {
                return Promise.resolve(undici.Response.json([91001, 91002, 91003]));
            }
            const id = /\/item\/(\d+)\.json$/.exec(url)?.[1];
            if (id && Object.hasOwn(stories, id)) {
                return Promise.resolve(undici.Response.json(stories[id]));
            }
            return Promise.resolve(new undici.Response('Sorry', { status: 419 }));
        });
        const testApp = new Hono();
        testApp.get('/hackernews/:section/:type?', async (ctx) => ctx.json(await hackernewsRoute.handler(ctx)));

        try {
            const response = await testApp.request('/hackernews/best?limit=2');
            expect(response.status).toBe(200);
            const data = await response.json();
            expect(data.link).toBe('https://news.ycombinator.com/best');
            expect(data.item).toMatchObject([
                { guid: '91001', title: 'An external story', link: 'https://example.com/article', author: 'alice', comments: 12, upvotes: 120 },
                { guid: '91002', title: 'Ask HN: A discussion', link: 'https://news.ycombinator.com/item?id=91002', author: 'bob' },
            ]);
            expect(data.item[0].pubDate).toBe(new Date(stories[91001].time * 1000).toISOString());
            expect(data.item[1].description).toContain('<p>Discussion body</p>');
            expect(requests).toHaveLength(3);
            expect(requests.every((url) => url.startsWith('https://hacker-news.firebaseio.com/'))).toBe(true);
        } finally {
            fetchSpy.mockRestore();
        }
    });
});

process.env.ALLOW_USER_SUPPLY_UNSAFE_DOMAIN = 'true';

const routes = {
    '/test/:id': '/test/1',
};
if (process.env.FULL_ROUTES_TEST) {
    const { namespaces } = await import('@/registry');
    for (const namespace in namespaces) {
        for (const route in namespaces[namespace].routes) {
            const requireConfig = namespaces[namespace].routes[route].features?.requireConfig;
            let configs;
            if (Array.isArray(requireConfig)) {
                configs = requireConfig
                    .filter((config) => !config.optional)
                    .map((config) => config.name)
                    .filter((name) => name !== 'ALLOW_USER_SUPPLY_UNSAFE_DOMAIN');
            }
            if (namespaces[namespace].routes[route].example && !configs?.length) {
                routes[`/${namespace}${route}`] = namespaces[namespace].routes[route].example;
            }
        }
    }
}

async function checkRSS(response) {
    const checkDate = (date) => {
        expect(date).toEqual(expect.any(String));
        expect(Date.parse(date)).toEqual(expect.any(Number));
        expect(Date.now() - +new Date(date)).toBeGreaterThan(-1000 * 60 * 60 * 24 * 5);
        expect(Date.now() - +new Date(date)).toBeLessThan(1000 * 60 * 60 * 24 * 30 * 12 * 10);
    };

    const parsed = await parser.parseString(await response.text());

    expect(parsed).toEqual(expect.any(Object));
    expect(parsed.title).toEqual(expect.any(String));
    expect(parsed.title).not.toBe('RSSHub');
    expect(parsed.description).toEqual(expect.any(String));
    expect(parsed.link).toEqual(expect.any(String));
    expect(parsed.lastBuildDate).toEqual(expect.any(String));
    expect(parsed.ttl).toEqual(Math.trunc(config.cache.routeExpire / 60) + '');
    expect(parsed.items).toEqual(expect.any(Array));
    checkDate(parsed.lastBuildDate);

    // check items
    const guids: Array<string | undefined> = [];
    for (const item of parsed.items) {
        expect(item).toEqual(expect.any(Object));
        expect(item.title).toEqual(expect.any(String));
        expect(item.link).toEqual(expect.any(String));
        expect(item.content).toEqual(expect.any(String));
        expect(item.guid).toEqual(expect.any(String));
        if (item.pubDate) {
            expect(item.pubDate).toEqual(expect.any(String));
            checkDate(item.pubDate);
        }

        // guid must be unique
        expect(guids).not.toContain(item.guid);
        guids.push(item.guid);
    }
}

describe('routes', () => {
    for (const route in routes) {
        it.concurrent(
            route,
            {
                timeout: 60000,
            },
            async () => {
                const response = await app.request(routes[route]);
                expect(response.status).toBe(200);
                await checkRSS(response);
            }
        );
    }
});

describe('Bilibili video risk control', () => {
    it('reports a v_voucher challenge without falling back to browser mode', async () => {
        vi.resetModules();
        const getPlaywrightPage = vi.fn();
        vi.doMock('@/routes/bilibili/cache', () => ({
            default: {
                getCookie: vi.fn().mockResolvedValue('cookie'),
                getRenderData: vi.fn().mockResolvedValue('webid'),
                getWbiVerifyString: vi.fn().mockResolvedValue('wbi-key'),
            },
        }));
        vi.doMock('@/utils/got', () => ({
            default: vi.fn().mockResolvedValue({
                data: { code: -352, data: { v_voucher: 'test-voucher' }, message: '风控校验失败' },
            }),
        }));
        vi.doMock('@/utils/playwright', () => ({ getPlaywrightPage }));

        try {
            const { route } = await import('@/routes/bilibili/video');
            const testApp = new Hono();
            testApp.get('/bilibili/user/video/:uid', async (ctx) => ctx.json(await route.handler(ctx)));
            testApp.onError((error, ctx) => ctx.json({ message: error.message }, 500));

            const response = await testApp.request('/bilibili/user/video/646730844');
            expect(response.status).toBe(500);
            expect(await response.json()).toMatchObject({ message: expect.stringContaining('v_voucher present; captcha/risk verification required') });
            expect(getPlaywrightPage).not.toHaveBeenCalled();
        } finally {
            vi.doUnmock('@/routes/bilibili/cache');
            vi.doUnmock('@/utils/got');
            vi.doUnmock('@/utils/playwright');
            vi.resetModules();
        }
    });
});
