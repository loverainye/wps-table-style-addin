import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, writeFile, mkdtemp, rm } from 'node:fs/promises';
import path from 'node:path';
import { tmpdir } from 'node:os';
import { wpsLibcPaths } from '../scripts/linux-libc.mjs';

test('macro build embeds WPS libc mappings without using other applications or deleted mappings', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'wps-proc-test-'));
  try {
    const store = '/nix/store/actual-wps-glibc/lib/libc.so.6';
    for (const [pid, comm, maps] of [
      ['1', 'node', '100-200 r-xp 0 0:0 1 /nix/store/node-glibc/lib/libc.so.6\n'],
      ['2', 'wpp', `100-200 r-xp 0 0:0 1 ${store}\n200-300 r--p 0 0:0 1 ${store}\n`],
      ['3', 'wpp', '100-200 r-xp 0 0:0 1 /nix/store/deleted/lib/libc.so.6 (deleted)\n'],
      ['4', 'wpp', '100-200 r-xp 0 0:0 1 /usr/lib/libc-2.31.so\n'],
    ]) {
      await mkdir(path.join(dir, pid));
      await writeFile(path.join(dir, pid, 'comm'), comm + '\n');
      await writeFile(path.join(dir, pid, 'maps'), maps);
    }
    await mkdir(path.join(dir, '5')); // Exiting or inaccessible processes are skipped.
    assert.deepEqual(await wpsLibcPaths(dir), [store, '/usr/lib/libc-2.31.so'].sort());
    assert.deepEqual(await wpsLibcPaths(path.join(dir, 'absent')), []);
  } finally { await rm(dir, { recursive: true, force: true }); }
});
