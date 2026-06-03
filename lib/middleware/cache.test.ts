import { Context } from 'hono';
import Parser from 'rss-parser';
import { afterEach, describe, expect, it, vi } from 'vitest';
import xxhash from 'xxhash-wasm';

import wait from '@/utils/wait';

process.env.CACHE_EXPIRE = '1';
process.env.CACHE_CONTENT_EXPIRE = '2';

const parser = new Parser();

afterEach(() => {
    vi.resetModules();
});

const noCacheTestFunc = async () => {
    const app = (await import('@/app')).default;

    const response1 = await app.request('/test/cache');
    const response2 = await app.request('/test/cache');

    const parsed1 = await parser.parseString(await response1.text());
    const parsed2 = await parser.parseString(await response2.text());

    expect(response2.status).toBe(200);
    expect(response2.headers).not.toHaveProperty('rsshub-cache-status');

    expect(parsed1.items[0].content).toBe('Cache1');
    expect(parsed2.items[0].content).toBe('Cache2');

    expect(parsed1.ttl).toEqual('1');
};

describe('cache', () => {
    it('memory', async () => {
        process.env.CACHE_TYPE = 'memory';
        const app = (await import('@/app')).default;

        const response1 = await app.request('/test/cache');
        const response2 = await app.request('/test/cache');

        const parsed1 = await parser.parseString(await response1.text());
        const parsed2 = await parser.parseString(await response2.text());

        delete parsed1.lastBuildDate;
        delete parsed2.lastBuildDate;
        delete parsed1.feedUrl;
        delete parsed2.feedUrl;
        delete parsed1.paginationLinks;
        delete parsed2.paginationLinks;
        expect(parsed2).toMatchObject(parsed1);

        expect(response2.status).toBe(200);
        expect(response2.headers.get('rsshub-cache-status')).toBe('HIT');

        expect(parsed1.ttl).toEqual('1');

        await wait(1 * 1000 + 100);
        const response3 = await app.request('/test/cache');
        expect(response3.headers).not.toHaveProperty('rsshub-cache-status');
        const parsed3 = await parser.parseString(await response3.text());

        await wait(2 * 1000 + 100);
        const response4 = await app.request('/test/cache');
        const parsed4 = await parser.parseString(await response4.text());

        expect(parsed1.items[0].content).toBe('Cache1');
        expect(parsed2.items[0].content).toBe('Cache1');
        expect(parsed3.items[0].content).toBe('Cache1');
        expect(parsed4.items[0].content).toBe('Cache2');

        await app.request('/test/refreshCache');
        await wait(1 * 1000 + 100);
        const response5 = await app.request('/test/refreshCache');
        const parsed5 = await parser.parseString(await response5.text());
        await wait(1 * 1000 + 100);
        const response6 = await app.request('/test/refreshCache');
        const parsed6 = await parser.parseString(await response6.text());

        expect(parsed5.items[0].content).toBe('1 1');
        expect(parsed6.items[0].content).toBe('1 0');
    }, 10000);

    it('redis', async () => {
        process.env.CACHE_TYPE = 'redis';
        const app = (await import('@/app')).default;

        await wait(500);
        const response1 = await app.request('/test/cache');
        const response2 = await app.request('/test/cache');

        const parsed1 = await parser.parseString(await response1.text());
        const parsed2 = await parser.parseString(await response2.text());

        delete parsed1.lastBuildDate;
        delete parsed2.lastBuildDate;
        delete parsed1.feedUrl;
        delete parsed2.feedUrl;
        delete parsed1.paginationLinks;
        delete parsed2.paginationLinks;
        expect(parsed2).toMatchObject(parsed1);

        expect(response2.status).toBe(200);
        expect(response2.headers.get('rsshub-cache-status')).toBe('HIT');

        expect(parsed1.ttl).toEqual('1');

        await wait(1 * 1000 + 100);
        const response3 = await app.request('/test/cache');
        expect(response3.headers).not.toHaveProperty('rsshub-cache-status');
        const parsed3 = await parser.parseString(await response3.text());

        await wait(2 * 1000 + 100);
        const response4 = await app.request('/test/cache');
        const parsed4 = await parser.parseString(await response4.text());

        expect(parsed1.items[0].content).toBe('Cache1');
        expect(parsed2.items[0].content).toBe('Cache1');
        expect(parsed3.items[0].content).toBe('Cache1');
        expect(parsed4.items[0].content).toBe('Cache2');

        await app.request('/test/refreshCache');
        await wait(1 * 1000 + 100);
        const response5 = await app.request('/test/refreshCache');
        const parsed5 = await parser.parseString(await response5.text());
        await wait(1 * 1000 + 100);
        const response6 = await app.request('/test/refreshCache');
        const parsed6 = await parser.parseString(await response6.text());

        expect(parsed5.items[0].content).toBe('1 1');
        expect(parsed6.items[0].content).toBe('1 0');

        const cache = (await import('@/utils/cache')).default;
        await cache.clients.redisClient!.quit();
    }, 10000);

    it('redis with quit', async () => {
        process.env.CACHE_TYPE = 'redis';
        const cache = (await import('@/utils/cache')).default;
        await cache.clients.redisClient!.quit();
        await noCacheTestFunc();
    });

    it('redis with error', async () => {
        process.env.CACHE_TYPE = 'redis';
        process.env.REDIS_URL = 'redis://wrongpath:6379';
        await noCacheTestFunc();
        const cache = (await import('@/utils/cache')).default;
        cache.clients.redisClient?.disconnect();
    }, 20000);

    it('no cache', async () => {
        process.env.CACHE_TYPE = 'NO';
        await noCacheTestFunc();
    });

    it('no cache (empty string)', async () => {
        process.env.CACHE_TYPE = '';
        await noCacheTestFunc();
    });

    it('smooth cache stale response and refresh bypass', async () => {
        process.env.CACHE_TYPE = 'memory';
        process.env.CACHE_EXPIRE = '1';
        process.env.CACHE_CONTENT_EXPIRE = '1';
        process.env.CACHE_SMOOTH = '1';
        process.env.CACHE_SMOOTH_PERIOD = '60';
        process.env.CACHE_SMOOTH_STALE_EXPIRE = '10';
        process.env.CACHE_SMOOTH_REFRESH_BASE_URL = 'http://127.0.0.1:9';
        process.env.CACHE_SMOOTH_REFRESH_TOKEN = 'test-token';
        process.env.REQUEST_TIMEOUT = '100';

        try {
            const app = (await import('@/app')).default;

            const response1 = await app.request('/test/cache');
            const parsed1 = await parser.parseString(await response1.text());

            const hitResponse = await app.request('/test/cache');
            const parsedHit = await parser.parseString(await hitResponse.text());
            expect(hitResponse.headers.get('rsshub-cache-status')).toBe('HIT');
            expect(hitResponse.headers.get('rsshub-cache-smooth-refresh-after')).not.toBeNull();

            const cacheModule = (await import('@/utils/cache/index')).default;
            const { h64ToString } = await xxhash();
            const cacheHash = h64ToString('/test/cache:rss');
            expect(await cacheModule.globalCache.get(`rsshub:koa-redis-cache:${cacheHash}`)).toBe('rsshub:smooth:fresh');
            expect(await cacheModule.globalCache.get(`rsshub:koa-redis-cache-stale:${cacheHash}`)).toContain('Cache1');

            await wait(1 * 1000 + 100);

            const response2 = await app.request('/test/cache');
            const parsed2 = await parser.parseString(await response2.text());
            expect(response2.headers.get('rsshub-cache-status')).toBe('STALE');
            expect(response2.headers.get('rsshub-cache-smooth-refresh-after')).not.toBeNull();

            const response3 = await app.request('/test/cache', {
                headers: {
                    'RSSHub-Cache-Smooth-Refresh': 'test-token',
                },
            });
            const parsed3 = await parser.parseString(await response3.text());
            expect(response3.headers.get('rsshub-cache-status')).toBe('REFRESH');

            await wait(1 * 1000 + 100);

            const response4 = await app.request('/test/cache');
            const parsed4 = await parser.parseString(await response4.text());

            expect(parsed1.items[0].content).toBe('Cache1');
            expect(parsedHit.items[0].content).toBe('Cache1');
            expect(parsed2.items[0].content).toBe('Cache1');
            expect(parsed3.items[0].content).toBe('Cache2');
            expect(parsed4.items[0].content).toBe('Cache2');
        } finally {
            delete process.env.CACHE_SMOOTH;
            delete process.env.CACHE_SMOOTH_PERIOD;
            delete process.env.CACHE_SMOOTH_STALE_EXPIRE;
            delete process.env.CACHE_SMOOTH_REFRESH_BASE_URL;
            delete process.env.CACHE_SMOOTH_REFRESH_TOKEN;
            delete process.env.REQUEST_TIMEOUT;
            process.env.CACHE_CONTENT_EXPIRE = '2';
        }
    }, 10000);

    it('smooth cache serves stale data while refresh is in progress', async () => {
        process.env.CACHE_TYPE = 'memory';
        process.env.CACHE_EXPIRE = '1';
        process.env.CACHE_CONTENT_EXPIRE = '1';
        process.env.CACHE_SMOOTH = '1';
        process.env.CACHE_SMOOTH_PERIOD = '60';
        process.env.CACHE_SMOOTH_STALE_EXPIRE = '10';
        process.env.CACHE_SMOOTH_REFRESH_TOKEN = 'test-token';

        try {
            const app = (await import('@/app')).default;

            await app.request('/test/slow');
            await wait(1 * 1000 + 100);

            const refresh = app.request('/test/slow', {
                headers: {
                    'RSSHub-Cache-Smooth-Refresh': 'test-token',
                },
            });
            await wait(100);

            const staleResponse = await app.request('/test/slow');
            expect(staleResponse.headers.get('rsshub-cache-status')).toBe('STALE');
            expect(staleResponse.headers.get('rsshub-cache-smooth-refresh-after')).toBe('0');

            await refresh;
        } finally {
            delete process.env.CACHE_SMOOTH;
            delete process.env.CACHE_SMOOTH_PERIOD;
            delete process.env.CACHE_SMOOTH_STALE_EXPIRE;
            delete process.env.CACHE_SMOOTH_REFRESH_TOKEN;
            process.env.CACHE_CONTENT_EXPIRE = '2';
        }
    }, 10000);

    it('smooth cache marker remains safe when smooth config changes', async () => {
        process.env.CACHE_TYPE = 'memory';
        process.env.CACHE_EXPIRE = '3';
        process.env.CACHE_CONTENT_EXPIRE = '3';
        process.env.CACHE_SMOOTH = '1';
        process.env.CACHE_SMOOTH_PERIOD = '60';
        process.env.CACHE_SMOOTH_STALE_EXPIRE = '1';
        process.env.CACHE_SMOOTH_REFRESH_BASE_URL = 'http://127.0.0.1:9';
        process.env.REQUEST_TIMEOUT = '100';

        try {
            const app = (await import('@/app')).default;
            const { config } = await import('@/config');

            const response1 = await app.request('/test/cache');
            const parsed1 = await parser.parseString(await response1.text());

            const cacheModule = (await import('@/utils/cache/index')).default;
            const { h64ToString } = await xxhash();
            const cacheHash = h64ToString('/test/cache:rss');
            expect(await cacheModule.globalCache.get(`rsshub:koa-redis-cache:${cacheHash}`)).toBe('rsshub:smooth:fresh');
            expect(await cacheModule.globalCache.get(`rsshub:koa-redis-cache-stale:${cacheHash}`)).toContain('Cache1');

            config.cache.smooth.enabled = false;

            const disabledHitResponse = await app.request('/test/cache');
            const parsedDisabledHit = await parser.parseString(await disabledHitResponse.text());
            expect(disabledHitResponse.headers.get('rsshub-cache-status')).toBe('HIT');
            expect(disabledHitResponse.headers.get('rsshub-cache-smooth-refresh-after')).toBeNull();

            config.cache.smooth.enabled = true;
            await wait(1 * 1000 + 500);

            const ttlHitResponse = await app.request('/test/cache');
            const parsedTtlHit = await parser.parseString(await ttlHitResponse.text());
            expect(ttlHitResponse.headers.get('rsshub-cache-status')).toBe('HIT');

            expect(parsed1.items[0].content).toBe('Cache1');
            expect(parsedDisabledHit.items[0].content).toBe('Cache1');
            expect(parsedTtlHit.items[0].content).toBe('Cache1');
        } finally {
            delete process.env.CACHE_SMOOTH;
            delete process.env.CACHE_SMOOTH_PERIOD;
            delete process.env.CACHE_SMOOTH_STALE_EXPIRE;
            delete process.env.CACHE_SMOOTH_REFRESH_BASE_URL;
            delete process.env.REQUEST_TIMEOUT;
            process.env.CACHE_EXPIRE = '1';
            process.env.CACHE_CONTENT_EXPIRE = '2';
        }
    }, 10000);

    it('throws URL key', async () => {
        process.env.CACHE_TYPE = 'memory';
        const app = (await import('@/app')).default;

        try {
            const response = await app.request('/test/cacheUrlKey');
            expect(response).toThrow(Error);
        } catch (error: any) {
            expect(error.message).toContain('Cache key must be a string');
        }
    });

    it('RSS TTL (no cache)', async () => {
        process.env.CACHE_TYPE = '';
        process.env.CACHE_EXPIRE = '600';
        const app = (await import('@/app')).default;
        const response = await app.request('/test/cache');
        const parsed = await parser.parseString(await response.text());
        expect(parsed.ttl).toEqual('1');
    });

    it('RSS TTL (w/ cache)', async () => {
        process.env.CACHE_TYPE = 'memory';
        process.env.CACHE_EXPIRE = '600';
        const app = (await import('@/app')).default;
        const response = await app.request('/test/cache');
        const parsed = await parser.parseString(await response.text());
        expect(parsed.ttl).toEqual('10');
    });
});

describe('cache middleware error handling', () => {
    it('clears control key when downstream throws', async () => {
        process.env.CACHE_TYPE = 'memory';
        const cache = (await import('@/utils/cache')).default;
        const setSpy = vi.spyOn(cache.globalCache, 'set');

        const { default: cacheMiddleware } = await import('@/middleware/cache');

        const ctx = new Context(new Request('http://localhost/test'), { env: {}, path: '/test' });

        await expect(
            cacheMiddleware(ctx, () => {
                throw new Error('boom');
            })
        ).rejects.toThrow('boom');

        expect(setSpy.mock.calls.some(([key, value]) => key.startsWith('rsshub:path-requested:') && value === '0')).toBe(true);
        setSpy.mockRestore();
    });
});

const setupSmoothCache = async () => {
    process.env.CACHE_TYPE = 'memory';
    const { config } = await import('@/config');
    config.cache.routeExpire = 30;
    config.cache.smooth.enabled = true;
    config.cache.smooth.includePathPrefixes = [];
    config.cache.smooth.excludePathPrefixes = [];
    const smooth = await import('@/utils/cache/smooth');
    vi.spyOn(smooth, 'scheduleSmoothRefresh').mockResolvedValue(false);
    const cache = (await import('@/utils/cache')).default;
    const cacheMiddleware = (await import('@/middleware/cache')).default;
    return { cache, cacheMiddleware, config };
};

describe('smooth cache request coordination', () => {
    afterEach(() => {
        vi.useRealTimers();
        vi.restoreAllMocks();
    });

    it.each([false, true])('fetches once for concurrent misses with an orphaned marker: %s', async (orphanedMarker) => {
        const { cache, cacheMiddleware, config } = await setupSmoothCache();
        const path = '/test/smooth-concurrent';
        const { h64ToString } = await xxhash();
        const cacheHash = h64ToString(`${path}:${config.format}`);
        if (orphanedMarker) {
            await cache.globalCache.set(`rsshub:koa-redis-cache:${cacheHash}`, 'rsshub:smooth:fresh', 30);
        }

        const contexts = [new Context(new Request(`http://localhost${path}`), { env: {}, path }), new Context(new Request(`http://localhost${path}`), { env: {}, path })];
        const release = Promise.withResolvers<void>();
        let fetches = 0;
        const next = async (ctx: Context) => {
            if (ctx.get('data')) {
                return;
            }
            fetches++;
            await release.promise;
            ctx.set('data', { title: 'cached feed', link: 'https://example.com', item: [{ title: 'item', link: 'https://example.com/1' }] });
        };

        vi.useFakeTimers();
        const requests = Promise.all(contexts.map((ctx) => cacheMiddleware(ctx, () => next(ctx))));
        try {
            await vi.waitFor(() => expect(fetches).toBeGreaterThan(0));
            release.resolve();
            await vi.advanceTimersByTimeAsync(3000);
            await requests;

            expect(fetches).toBe(1);
            expect(contexts[1].get('data')).toEqual(contexts[0].get('data'));
            expect(contexts[1].res.headers.get('RSSHub-Cache-Status')).toBe('HIT');
            expect(await cache.globalCache.get(`rsshub:path-requested:${cacheHash}`)).toBe('0');
        } finally {
            release.resolve();
        }
    });

    it('allows a new fetch after storing the stale payload fails', async () => {
        const { cache, cacheMiddleware } = await setupSmoothCache();
        const path = '/test/smooth-write-error';
        const ctx = new Context(new Request(`http://localhost${path}`), { env: {}, path });
        const data = { title: 'cached feed', link: 'https://example.com', item: [{ title: 'item', link: 'https://example.com/1' }] };
        const originalSet = cache.globalCache.set;
        const setSpy = vi.spyOn(cache.globalCache, 'set').mockImplementation((key, value, maxAge) => {
            if (key.startsWith('rsshub:koa-redis-cache-stale:')) {
                throw new Error('stale cache write failed');
            }
            return originalSet(key, value, maxAge);
        });

        await expect(cacheMiddleware(ctx, () => Promise.resolve(ctx.set('data', data)))).rejects.toThrow('stale cache write failed');

        setSpy.mockRestore();
        const retryCtx = new Context(new Request(`http://localhost${path}`), { env: {}, path });
        await cacheMiddleware(retryCtx, () => Promise.resolve(retryCtx.set('data', data)));
        expect(retryCtx.get('data')).toEqual(data);
    });
});
