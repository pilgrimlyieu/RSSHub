import { execFile, spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdir, mkdtemp, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createInterface } from 'node:readline';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import type { OutlookMailAccount } from '../lib/utils/mail-auth/account';
import { createState, readState, statePath, withStateLock, writeState } from '../lib/utils/mail-auth/store';

const exec = promisify(execFile);
const projectRoot = fileURLToPath(new URL('..', import.meta.url));
const storeUrl = new URL('../lib/utils/mail-auth/store.ts', import.meta.url).href;
const account: OutlookMailAccount = {
    email: 'alice@outlook.com',
    username: 'alice@outlook.com',
    host: 'outlook.office365.com',
    port: 993,
    auth: 'oauth2',
    provider: 'outlook',
    clientId: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa',
    tenant: 'consumers',
};
let root: string;
let directory: string;

beforeEach(async () => {
    root = await mkdtemp(path.join(os.tmpdir(), 'rsshub-mail-auth-store-'));
    directory = path.join(root, 'state');
});

afterEach(async () => {
    await rm(root, { recursive: true, force: true });
});

async function saveState(accessToken = 'original-access') {
    const state = createState(account, { accessToken, refreshToken: 'private-refresh', expiresAt: Date.now() + 3_600_000 });
    await withStateLock(account, directory, (lease) => writeState(account, directory, state, lease));
    return state;
}

function startLockWorker(source: string) {
    const child = spawn(process.execPath, ['--import', 'tsx', '--input-type=module', '-e', source], { cwd: projectRoot, stdio: 'pipe', timeout: 12000 });
    const ready = Promise.withResolvers<boolean>();
    let stderr = '';
    child.stderr.on('data', (chunk: Buffer) => {
        stderr += chunk.toString();
    });
    const completion = new Promise<{ code: number | null; stderr: string }>((resolve) => {
        child.once('error', (error) => {
            ready.resolve(false);
            resolve({ code: -1, stderr: error.message });
        });
        child.once('close', (code) => {
            ready.resolve(false);
            resolve({ code, stderr });
        });
    });
    const output = createInterface({ input: child.stdout });
    output.once('line', (line: string) => ready.resolve(line === 'ready'));
    return { child, completion, ready: ready.promise };
}

describe('mail OAuth state files', () => {
    it('creates a private directory and file before any prior token file exists', async () => {
        const state = await saveState();
        const file = statePath(account, directory);

        expect(await readState(account, directory)).toEqual(state);
        expect(await readdir(directory)).toEqual([path.basename(file)]);
        if (process.platform !== 'win32') {
            expect((await stat(directory)).mode & 0o777).toBe(0o700);
            expect((await stat(file)).mode & 0o777).toBe(0o600);
        }
    });

    it('uses a hash filename and normalizes identity without using the route alias', () => {
        const file = statePath(account, directory);

        expect(path.basename(file)).toMatch(/^[\da-f]{64}\.json$/);
        expect(file).not.toContain(account.username);
        expect(statePath({ ...account, email: 'route-alias@example.com' }, directory)).toBe(file);
        expect(statePath({ ...account, username: account.username.toUpperCase(), clientId: account.clientId.toUpperCase(), tenant: account.tenant.toUpperCase() }, directory)).toBe(file);
        expect(statePath({ ...account, username: 'another@outlook.com' }, directory)).not.toBe(file);
        expect(statePath({ ...account, clientId: '22222222-2222-4222-8222-222222222222' }, directory)).not.toBe(file);
        expect(statePath({ ...account, tenant: 'common' }, directory)).not.toBe(file);
    });

    it('returns no state for a missing file', async () => {
        expect(await readState(account, directory)).toBeUndefined();
    });

    it('rejects damaged JSON without exposing its content or parser cause', async () => {
        await mkdir(directory);
        await writeFile(statePath(account, directory), '{"refreshToken":"private-corrupted-refresh", broken JSON');
        let failure: unknown;
        try {
            await readState(account, directory);
        } catch (error) {
            failure = error;
        }

        expect(failure).toMatchObject({ name: 'MailAuthError', message: expect.stringContaining('state is damaged') });
        expect(String(failure)).not.toContain('private-corrupted-refresh');
        expect(String(failure)).not.toContain('broken JSON');
        expect(failure).not.toHaveProperty('cause');
    });

    it.each([
        ['unsupported version', { version: 2 }],
        ['empty access token', { accessToken: '' }],
        ['missing refresh token', { refreshToken: undefined }],
        ['string expiry', { expiresAt: '3600' }],
        ['negative expiry', { expiresAt: -1 }],
        ['fractional expiry', { expiresAt: 1.5 }],
        ['different identity', { identity: { provider: 'outlook', username: 'another@outlook.com', clientId: account.clientId, tenant: 'consumers' } }],
    ])('rejects %s in a state file', async (_name, change) => {
        const state = await saveState();
        await writeFile(statePath(account, directory), JSON.stringify({ ...state, ...change }));

        await expect(readState(account, directory)).rejects.toThrow('state is damaged or does not match this account');
    });

    it('reads the persisted credentials in a fresh Node process', async () => {
        const state = await saveState();
        const source = `
            import assert from 'node:assert/strict';
            import { readState } from ${JSON.stringify(storeUrl)};
            const state = await readState(${JSON.stringify(account)}, ${JSON.stringify(directory)});
            assert.deepEqual(state, ${JSON.stringify(state)});
            process.stdout.write('state-reloaded');
        `;

        const result = await exec(process.execPath, ['--import', 'tsx', '--input-type=module', '-e', source], { cwd: projectRoot, timeout: 10000 });

        expect(result.stdout).toBe('state-reloaded');
        expect(result.stderr).toBe('');
    });
});

describe('mail OAuth state locking', () => {
    it('releases the lock after a failed operation so a later write can succeed', async () => {
        const failure = new Error('Refresh request failed.');

        await expect(withStateLock(account, directory, () => Promise.reject(failure))).rejects.toBe(failure);
        const state = await saveState('recovered-access');

        expect(await readState(account, directory)).toEqual(state);
        expect(await readdir(directory)).toEqual([path.basename(statePath(account, directory))]);
    });

    it('prevents lost updates when two Node processes contend for the same state lock', async () => {
        await saveState('0');
        const source = `
            import { once } from 'node:events';
            import { setTimeout as delay } from 'node:timers/promises';
            import { createState, readState, withStateLock, writeState } from ${JSON.stringify(storeUrl)};
            const account = ${JSON.stringify(account)};
            const directory = ${JSON.stringify(directory)};
            const start = once(process.stdin, 'data');
            process.stdout.write('ready\\n');
            await start;
            await withStateLock(account, directory, async (lease) => {
                const previous = await readState(account, directory);
                await delay(250);
                const next = createState(account, {
                    accessToken: String(Number(previous.accessToken) + 1),
                    refreshToken: previous.refreshToken,
                    expiresAt: previous.expiresAt,
                });
                await writeState(account, directory, next, lease);
            });
        `;
        const workers = [startLockWorker(source), startLockWorker(source)];
        try {
            expect(await Promise.all(workers.map((worker) => worker.ready))).toEqual([true, true]);
            for (const worker of workers) {
                worker.child.stdin.end('start\n');
            }
            expect(await Promise.all(workers.map((worker) => worker.completion))).toEqual([
                { code: 0, stderr: '' },
                { code: 0, stderr: '' },
            ]);
            expect(await readState(account, directory)).toMatchObject({ accessToken: '2', refreshToken: 'private-refresh' });
            expect(await readdir(directory)).toEqual([path.basename(statePath(account, directory))]);
        } finally {
            for (const worker of workers) {
                worker.child.kill();
            }
            await Promise.all(workers.map((worker) => worker.completion));
        }
    }, 15000);

    it('aborts a compromised lease and preserves the old state instead of writing through a lost lock', async () => {
        await saveState();
        const file = statePath(account, directory);
        const previous = await readFile(file, 'utf8');

        await expect(
            withStateLock(account, directory, async (lease) => {
                const aborted = once(lease.signal, 'abort');
                await rm(`${file}.lock`, { recursive: true });
                await aborted;

                expect(lease.signal.aborted).toBe(true);
                expect(lease.assertActive).toThrow('lock was lost');
                const replacement = createState(account, { accessToken: 'must-not-be-written', refreshToken: 'must-not-replace-refresh', expiresAt: Date.now() + 3_600_000 });
                await expect(writeState(account, directory, replacement, lease)).rejects.toThrow('Unable to save mail OAuth state');
            })
        ).rejects.toThrow('lock was lost');

        expect(await readFile(file, 'utf8')).toBe(previous);
        expect(await readdir(directory)).toEqual([path.basename(file)]);
    }, 15000);
});
