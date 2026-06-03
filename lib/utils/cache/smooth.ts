import { createHash, randomBytes } from 'node:crypto';

import { config } from '@/config';
import cacheModule from '@/utils/cache/index';
import logger from '@/utils/logger';

export const smoothRefreshHeader = 'RSSHub-Cache-Smooth-Refresh';

const smoothRefreshToken = process.env.CACHE_SMOOTH_REFRESH_TOKEN || randomBytes(16).toString('hex');
const scheduledRefreshes = new Map<string, ReturnType<typeof setTimeout>>();

const getRefreshUrl = (requestUrl: string) => {
    const url = new URL(requestUrl);
    const baseUrl = config.cache.smooth.refreshBaseUrl || `http://127.0.0.1:${config.connect.port}`;
    return `${baseUrl.replace(/\/$/, '')}${url.pathname}${url.search}`;
};

const getSafeLogPath = (requestUrl: string) => {
    try {
        return new URL(requestUrl).pathname;
    } catch {
        return requestUrl.split('?')[0];
    }
};

export const shouldSmoothPath = (requestPath: string) => {
    if (!config.cache.smooth.enabled) {
        return false;
    }

    const { excludePathPrefixes, includePathPrefixes } = config.cache.smooth;
    if (excludePathPrefixes.some((prefix) => requestPath.startsWith(prefix))) {
        return false;
    }

    return includePathPrefixes.length === 0 || includePathPrefixes.some((prefix) => requestPath.startsWith(prefix));
};

export const isSmoothRefreshRequest = (token: string | undefined) => config.cache.smooth.enabled && token === smoothRefreshToken;

export const getSmoothDelaySeconds = (cacheIdentity: string, now = Date.now()) => {
    const digest = createHash('sha256').update(cacheIdentity).digest('hex');
    const slot = Number(BigInt(`0x${digest.slice(0, 12)}`) % BigInt(config.cache.smooth.period));
    const elapsed = Math.floor(now / 1000) % config.cache.smooth.period;

    return (slot - elapsed + config.cache.smooth.period) % config.cache.smooth.period;
};

const refreshSmoothCache = async (requestUrl: string, refreshKey: string) => {
    const refreshUrl = getRefreshUrl(requestUrl);
    const safeLogPath = getSafeLogPath(refreshUrl);
    const start = Date.now();

    logger.info(`Smooth cache refresh started for ${safeLogPath}`);

    try {
        const response = await fetch(refreshUrl, {
            headers: {
                [smoothRefreshHeader]: smoothRefreshToken,
            },
            signal: AbortSignal.timeout(config.requestTimeout),
        });

        await response.arrayBuffer();

        if (response.ok) {
            logger.info(`Smooth cache refresh finished for ${safeLogPath}: HTTP ${response.status} ${Date.now() - start}ms`);
        } else {
            logger.warn(`Smooth cache refresh failed for ${safeLogPath}: HTTP ${response.status}`);
        }
    } catch (error) {
        logger.warn(`Smooth cache refresh failed for ${safeLogPath}: ${error}`);
    } finally {
        await cacheModule.globalCache.set(refreshKey, '', 1);
    }
};

export const scheduleSmoothRefresh = async (cacheHash: string, requestUrl: string, delaySeconds: number) => {
    const refreshKey = `rsshub:smooth-refresh:${cacheHash}`;

    if (scheduledRefreshes.has(cacheHash) || (await cacheModule.globalCache.get(refreshKey)) === '1') {
        return false;
    }

    await cacheModule.globalCache.set(refreshKey, '1', delaySeconds + config.cache.requestTimeout + 60);

    const timer = setTimeout(() => {
        scheduledRefreshes.delete(cacheHash);
        void refreshSmoothCache(requestUrl, refreshKey);
    }, delaySeconds * 1000);
    timer.unref?.();
    scheduledRefreshes.set(cacheHash, timer);
    logger.info(`Smooth cache refresh scheduled for ${getSafeLogPath(requestUrl)} in ${delaySeconds}s`);

    return true;
};
