import type { MiddlewareHandler } from 'hono';

import { getPath, time } from '@/utils/helpers';
import logger from '@/utils/logger';
import { requestMetric } from '@/utils/otel';

enum LogPrefix {
    Outgoing = '-->',
    Incoming = '<--',
    Error = 'xxx',
}

const colorStatus = (status: number) => {
    const out: { [key: string]: string } = {
        7: `\u001B[35m${status}\u001B[0m`,
        5: `\u001B[31m${status}\u001B[0m`,
        4: `\u001B[33m${status}\u001B[0m`,
        3: `\u001B[36m${status}\u001B[0m`,
        2: `\u001B[32m${status}\u001B[0m`,
        1: `\u001B[32m${status}\u001B[0m`,
        0: `\u001B[33m${status}\u001B[0m`,
    };

    const calculateStatus = Math.trunc(status / 100);

    return out[calculateStatus];
};

const getCacheLogSuffix = (headers: Headers) => {
    const cacheStatus = headers.get('RSSHub-Cache-Status');
    if (!cacheStatus) {
        return '';
    }

    const smoothRefreshAfter = headers.get('RSSHub-Cache-Smooth-Refresh-After');
    return smoothRefreshAfter ? ` cache=${cacheStatus} smooth=${smoothRefreshAfter}s` : ` cache=${cacheStatus}`;
};

const middleware: MiddlewareHandler = async (ctx, next) => {
    const { method, raw, routePath } = ctx.req;
    const path = getPath(raw);

    logger.info(`${LogPrefix.Incoming} ${method} ${path}`);

    const start = Date.now();

    await next();

    const status = ctx.res.status;
    const cacheLogSuffix = getCacheLogSuffix(ctx.res.headers);

    logger.info(`${LogPrefix.Outgoing} ${method} ${path} ${colorStatus(status)} ${time(start)}${cacheLogSuffix}`);
    requestMetric.success(Date.now() - start, { path: routePath, method, status });
};

export default middleware;
