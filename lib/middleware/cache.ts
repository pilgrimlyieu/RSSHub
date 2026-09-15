import type { MiddlewareHandler } from 'hono';
import xxhash from 'xxhash-wasm';

import { config } from '@/config';
import RequestInProgressError from '@/errors/types/request-in-progress';
import type { Data } from '@/types';
import cacheModule from '@/utils/cache/index';
import { getSmoothDelaySeconds, isSmoothRefreshRequest, scheduleSmoothRefresh, shouldSmoothPath, smoothRefreshHeader } from '@/utils/cache/smooth';

const bypassList = new Set(['/', '/robots.txt', '/logo.png', '/favicon.ico']);
const smoothFreshCacheMarker = 'rsshub:smooth:fresh';

const { h64ToString } = await xxhash();

const getCachedValue = async (key: string, staleKey: string) => {
    const value = await cacheModule.globalCache.get(key);
    return value === smoothFreshCacheMarker ? await cacheModule.globalCache.get(staleKey) : value;
};

// only give cache string, as the `!` condition tricky
// XXH64 is used to shrink key size
// plz, write these tips in comments!
const middleware: MiddlewareHandler = async (ctx, next) => {
    if (!cacheModule.status.available || bypassList.has(ctx.req.path)) {
        await next();
        return;
    }

    const requestPath = ctx.req.path;
    const format = `:${ctx.req.query('format') || config.format}`;
    const limit = ctx.req.query('limit') ? `:${ctx.req.query('limit')}` : '';
    const cacheIdentity = requestPath + format + limit;
    const cacheHash = h64ToString(cacheIdentity);
    const key = 'rsshub:koa-redis-cache:' + cacheHash;
    const staleKey = 'rsshub:koa-redis-cache-stale:' + cacheHash;
    const controlKey = 'rsshub:path-requested:' + cacheHash;
    const smoothEnabled = shouldSmoothPath(requestPath);
    const forceSmoothRefresh = smoothEnabled && isSmoothRefreshRequest(ctx.req.header(smoothRefreshHeader));

    let value = forceSmoothRefresh ? undefined : await getCachedValue(key, staleKey);

    if (smoothEnabled && !forceSmoothRefresh) {
        const isRefreshing = cacheModule.globalCache.supportsAtomicClaims && (await cacheModule.globalCache.get(controlKey)) === '1';
        if (!value || isRefreshing) {
            const staleValue = await cacheModule.globalCache.get(staleKey);
            if (staleValue) {
                const delaySeconds = isRefreshing ? 0 : getSmoothDelaySeconds(cacheIdentity);
                if (!isRefreshing) {
                    await scheduleSmoothRefresh(cacheHash, ctx.req.url, delaySeconds);
                }

                ctx.status(200);
                ctx.header('RSSHub-Cache-Status', 'STALE');
                ctx.header('RSSHub-Cache-Smooth-Refresh-After', delaySeconds.toString());
                ctx.set('data', JSON.parse(staleValue));
                await next();
                return;
            }
        }
    }

    // Only atomic backends can coordinate fetchers. HTTP/KV may return stale
    // control keys after a completed request, while their feed cache stays useful.
    let isRequesting = false;
    let ownsClaim = false;
    if (!value && cacheModule.globalCache.supportsAtomicClaims) {
        ownsClaim = await cacheModule.globalCache.claim(controlKey, config.cache.requestTimeout);
        isRequesting = !ownsClaim;
    }

    if (isRequesting) {
        let retryTimes = process.env.NODE_ENV === 'test' ? 1 : 10;
        let bypass = false;
        while (retryTimes > 0) {
            // eslint-disable-next-line no-await-in-loop
            await new Promise((resolve) => setTimeout(resolve, process.env.NODE_ENV === 'test' ? 3000 : 6000));
            // eslint-disable-next-line no-await-in-loop
            if ((await cacheModule.globalCache.get(controlKey)) !== '1') {
                bypass = true;
                break;
            }
            retryTimes--;
        }
        if (!bypass) {
            throw new RequestInProgressError('This path is currently fetching, please come back later!');
        }
        value = forceSmoothRefresh ? undefined : await getCachedValue(key, staleKey);
    }

    if (value) {
        if (smoothEnabled && !forceSmoothRefresh) {
            const delaySeconds = getSmoothDelaySeconds(cacheIdentity);
            await scheduleSmoothRefresh(cacheHash, ctx.req.url, delaySeconds);
            ctx.header('RSSHub-Cache-Smooth-Refresh-After', delaySeconds.toString());
        }

        ctx.status(200);
        ctx.header('RSSHub-Cache-Status', 'HIT');
        ctx.set('data', JSON.parse(value));
        await next();
        return;
    }

    if (isRequesting) {
        // waited out a stale claim without finding a cache entry, take over the fetch
        ownsClaim = await cacheModule.globalCache.claim(controlKey, config.cache.requestTimeout);
        if (!ownsClaim) {
            throw new RequestInProgressError('This path is currently fetching, please come back later!');
        }
    }

    // let routers control cache
    ctx.set('cacheKey', key);
    if (ownsClaim) {
        ctx.set('cacheControlKey', controlKey);
    }

    if (forceSmoothRefresh) {
        ctx.header('RSSHub-Cache-Status', 'REFRESH');
    }

    try {
        await next();

        const data: Data = ctx.get('data');
        if (ctx.res.headers.get('Cache-Control') !== 'no-cache' && data) {
            data.lastBuildDate = new Date().toUTCString();
            ctx.set('data', data);
            const body = JSON.stringify(data);
            if (smoothEnabled) {
                await cacheModule.globalCache.set(staleKey, body, Math.max(config.cache.smooth.staleExpire, config.cache.routeExpire));
                await cacheModule.globalCache.set(key, smoothFreshCacheMarker, config.cache.routeExpire);
            } else {
                await cacheModule.globalCache.set(key, body, config.cache.routeExpire);
            }
        }
    } finally {
        // Release after writing the feed, including failures after the route ran.
        if (ownsClaim) {
            await cacheModule.globalCache.set(controlKey, '0', config.cache.requestTimeout);
        }
    }
};

export default middleware;
