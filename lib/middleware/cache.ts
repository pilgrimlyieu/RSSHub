import type { MiddlewareHandler } from 'hono';
import xxhash from 'xxhash-wasm';

import { config } from '@/config';
import RequestInProgressError from '@/errors/types/request-in-progress';
import type { Data } from '@/types';
import cacheModule from '@/utils/cache/index';
import { getSmoothDelaySeconds, isSmoothRefreshRequest, scheduleSmoothRefresh, shouldSmoothPath, smoothRefreshHeader } from '@/utils/cache/smooth';

const bypassList = new Set(['/', '/robots.txt', '/logo.png', '/favicon.ico']);
const smoothFreshCacheMarker = 'rsshub:smooth:fresh';
// only give cache string, as the `!` condition tricky
// XXH64 is used to shrink key size
// plz, write these tips in comments!
const middleware: MiddlewareHandler = async (ctx, next) => {
    if (!cacheModule.status.available || bypassList.has(ctx.req.path)) {
        await next();
        return;
    }

    const requestPath = ctx.req.path;
    const format = `:${ctx.req.query('format') || 'rss'}`;
    const limit = ctx.req.query('limit') ? `:${ctx.req.query('limit')}` : '';
    const cacheIdentity = requestPath + format + limit;
    const { h64ToString } = await xxhash();
    const cacheHash = h64ToString(cacheIdentity);
    const key = 'rsshub:koa-redis-cache:' + cacheHash;
    const staleKey = 'rsshub:koa-redis-cache-stale:' + cacheHash;
    const controlKey = 'rsshub:path-requested:' + cacheHash;
    const smoothEnabled = shouldSmoothPath(requestPath);
    const forceSmoothRefresh = smoothEnabled && isSmoothRefreshRequest(ctx.req.header(smoothRefreshHeader));

    const isRequesting = await cacheModule.globalCache.get(controlKey);

    if (isRequesting === '1') {
        if (smoothEnabled && !forceSmoothRefresh) {
            const staleValue = await cacheModule.globalCache.get(staleKey);

            if (staleValue) {
                ctx.status(200);
                ctx.header('RSSHub-Cache-Status', 'STALE');
                ctx.header('RSSHub-Cache-Smooth-Refresh-After', '0');
                ctx.set('data', JSON.parse(staleValue));
                await next();
                return;
            }
        }

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
    }

    const value = forceSmoothRefresh ? undefined : await cacheModule.globalCache.get(key);
    const cachedValue = value === smoothFreshCacheMarker ? await cacheModule.globalCache.get(staleKey) : value;

    if (cachedValue) {
        if (smoothEnabled && !forceSmoothRefresh) {
            const delaySeconds = getSmoothDelaySeconds(cacheIdentity);
            await scheduleSmoothRefresh(cacheHash, ctx.req.url, delaySeconds);
            ctx.header('RSSHub-Cache-Smooth-Refresh-After', delaySeconds.toString());
        }

        ctx.status(200);
        ctx.header('RSSHub-Cache-Status', 'HIT');
        ctx.set('data', JSON.parse(cachedValue));
        await next();
        return;
    }

    if (smoothEnabled && !forceSmoothRefresh) {
        const staleValue = await cacheModule.globalCache.get(staleKey);

        if (staleValue) {
            const delaySeconds = getSmoothDelaySeconds(cacheIdentity);
            await scheduleSmoothRefresh(cacheHash, ctx.req.url, delaySeconds);

            ctx.status(200);
            ctx.header('RSSHub-Cache-Status', 'STALE');
            ctx.header('RSSHub-Cache-Smooth-Refresh-After', delaySeconds.toString());
            ctx.set('data', JSON.parse(staleValue));
            await next();
            return;
        }
    }

    // Doesn't hit the cache? We need to let others know!
    await cacheModule.globalCache.set(controlKey, '1', config.cache.requestTimeout);

    // let routers control cache
    ctx.set('cacheKey', key);
    ctx.set('cacheControlKey', controlKey);

    if (forceSmoothRefresh) {
        ctx.header('RSSHub-Cache-Status', 'REFRESH');
    }

    try {
        await next();
    } catch (error) {
        await cacheModule.globalCache.set(controlKey, '0', config.cache.requestTimeout);
        throw error;
    }

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

    // We need to let it go, even no cache set.
    // Wait to set cache so the next request could be handled correctly
    await cacheModule.globalCache.set(controlKey, '0', config.cache.requestTimeout);
};

export default middleware;
