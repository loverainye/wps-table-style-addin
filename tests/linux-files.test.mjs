import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { createLinuxFileSystem, macroApplication } from '../macro/linux-files.js';
import { buildMacro } from '../scripts/build.mjs';
import { fixture, plan } from './fixtures.mjs';
import { binaryString } from '../src/host.js';
import { openDeck, readPart } from '../src/styles.js';

function fakeLibc(initial = [], chunk = 1111) {
  const files = new Map(initial.map(([p, data]) => [p, Buffer.from(data)]));
  const fds = new Map(), calls = [], fail = {};
  let next = 3;
  const api = {
    open(p, flags, mode) {
      calls.push(['open', p, flags, mode]);
      if (fail.open || (flags === 0 && !files.has(p))) return -1;
      if (flags & 512) files.set(p, Buffer.alloc(0));
      const fd = next++; fds.set(fd, { p, at: 0 }); return fd;
    },
    read(fd, pointer, size) {
      assert.equal(typeof size, 'bigint');
      if (fail.read) return -1n;
      const f = fds.get(fd), data = files.get(f.p);
      const used = Math.min(chunk, Number(size), data.length - f.at);
      new Uint8Array(pointer).set(data.subarray(f.at, f.at + used));
      f.at += used; return BigInt(used);
    },
    write(fd, pointer, size) {
      assert.equal(typeof size, 'bigint');
      if (fail.write) return -1n;
      if (fail.zeroWrite) return 0n;
      const f = fds.get(fd), used = Math.min(chunk, Number(size));
      const bytes = Buffer.from(new Uint8Array(pointer)).subarray(0, used);
      files.set(f.p, Buffer.concat([files.get(f.p), bytes]));
      f.at += used; return BigInt(used);
    },
    fsync(fd) { calls.push(['fsync', fd]); return fail.fsync ? -1 : 0; },
    close(fd) { calls.push(['close', fd]); fds.delete(fd); return fail.close ? -1 : 0; },
    access(p) { return files.has(p) ? 0 : -1; },
    unlink(p) { return files.delete(p) ? 0 : -1; },
  };
  const ffi = { LoadLibrary(name, signature) {
    calls.push(['LoadLibrary', name]);
    assert.deepEqual(signature.open.parameters, ['string', 'int32', 'uint32']);
    assert.deepEqual(signature.read.parameters, ['int32', 'pointer', 'uint64']);
    assert.equal(signature.write.returnType, 'int64');
    return api;
  } };
  return { ffi, files, fds, calls, fail };
}

test('macro loads the WPS absolute libc hint when its FFI rejects sonames and standard paths', () => {
  const m = fakeLibc(), library = '/nix/store/test-glibc/lib/libc.so.6', attempted = [];
  const ffi = { LoadLibrary(name, signature) {
    attempted.push(name);
    if (name !== library) throw new Error('Library not found');
    return m.ffi.LoadLibrary(name, signature);
  } };
  const fs = createLinuxFileSystem(ffi, [library]);
  assert.deepEqual(attempted, [library]);
  assert.equal(fs.library, library);
  fs.writeAsBinaryString('/test', 'working');
  assert.equal(fs.readAsBinaryString('/test'), 'working');
});

test('macro falls back from an unavailable build hint to an absolute distro path', () => {
  const m = fakeLibc(), attempted = [], library = '/lib64/libc.so.6';
  const fs = createLinuxFileSystem({ LoadLibrary(name, signature) {
    attempted.push(name);
    if (name !== library) throw new Error('Library not found');
    return m.ffi.LoadLibrary(name, signature);
  } }, ['/nix/store/old-glibc/lib/libc.so.6']);
  assert.equal(attempted[0], '/nix/store/old-glibc/lib/libc.so.6');
  assert.equal(fs.library, library);
  assert.equal(new Set(attempted).size, attempted.length);
});

test('macro reports attempted library paths and loader errors when no libc can load', () => {
  assert.throws(() => createLinuxFileSystem({ LoadLibrary() { throw new Error('Library not found'); } },
    ['/nix/store/missing/lib/libc.so.6']), error => {
    assert.match(error.message, /无法加载 Linux libc/);
    assert.match(error.message, /\/nix\/store\/missing\/lib\/libc.so.6/);
    assert.match(error.message, /Library not found/);
    return true;
  });
  assert.throws(() => createLinuxFileSystem({ LoadLibrary() { return { open() {} }; } }), /文件接口不完整/);
});

test('macro libc adapter preserves all binary bytes with short reads and writes across chunks', () => {
  const data = Buffer.from(Array.from({ length: 140000 }, (_, i) => i % 256));
  const m = fakeLibc([['/文稿/测试.pptx', data]]), fs = createLinuxFileSystem(m.ffi);
  assert.equal(fs.readAsBinaryString('/文稿/测试.pptx'), binaryString(data));
  fs.writeAsBinaryString('/文稿/保存.pptx', binaryString(data));
  assert.deepEqual(m.files.get('/文稿/保存.pptx'), data);
  assert.equal(m.fds.size, 0);
  assert.equal(m.calls.filter(c => c[0] === 'fsync').length, 1);
  assert.equal(fs.Exists('/文稿/保存.pptx'), true);
  fs.Remove('/文稿/保存.pptx'); assert.equal(fs.Exists('/文稿/保存.pptx'), false);
});

test('macro libc adapter reads UTF-8 and empty files and rejects invalid paths before opening', () => {
  const m = fakeLibc([['/文本', Buffer.from('中文🙂')], ['/空', Buffer.alloc(0)]]);
  const fs = createLinuxFileSystem(m.ffi);
  assert.equal(fs.readFileString('/文本'), '中文🙂');
  assert.equal(fs.readAsBinaryString('/空'), '');
  fs.writeAsBinaryString('/新空', ''); assert.equal(m.files.get('/新空').length, 0);
  const count = m.calls.length;
  for (const invalid of ['relative.pptx', '/a\0b', null]) {
    assert.throws(() => fs.readAsBinaryString(invalid), /绝对路径/);
    assert.throws(() => fs.writeAsBinaryString(invalid, 'x'), /绝对路径/);
  }
  assert.equal(m.calls.length, count, 'invalid paths must never open or truncate a file');
});

test('macro libc adapter reports I/O failures and closes each descriptor exactly once', () => {
  for (const failure of ['read', 'write', 'zeroWrite', 'fsync', 'close', 'open']) {
    const m = fakeLibc([['/source', Buffer.from('test')]]), fs = createLinuxFileSystem(m.ffi);
    m.fail[failure] = true;
    assert.throws(() => failure === 'read' ? fs.readAsBinaryString('/source') : fs.writeAsBinaryString('/target', 'test'), /失败/);
    assert.equal(m.fds.size, 0, failure);
    assert.equal(m.calls.filter(c => c[0] === 'close').length, failure === 'open' ? 0 : 1, failure);
  }
});

test('macro Application adapter leaves native proxies untouched and delegates live documents', () => {
  const m = fakeLibc(); let active = { FullName: '/first.pptx' };
  const app = Object.freeze({ get ActivePresentation() { return active; },
    ActiveWindow: {}, Presentations: {}, FileDialog(kind) { assert.equal(this, app); return kind; } });
  const host = macroApplication(app, { ffi: m.ffi });
  assert.equal(host.FileDialog(3), 3);
  active = { FullName: '/second.pptx' };
  assert.equal(host.ActivePresentation, active);
  assert.equal(host.ActiveWindow, app.ActiveWindow);
  assert.equal(host.Presentations, app.Presentations);
  assert.equal(app.FileSystem, undefined);
  const existing = { FileSystem: { readAsBinaryString() {}, writeAsBinaryString() {} } };
  assert.equal(macroApplication(existing), existing);
  assert.throws(() => macroApplication(app, {}), /ffi.LoadLibrary/);
});

test('bundled macro with FFI and no Application.FileSystem edits a closed PPTX and keeps a backup', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'wps-macro-libc-'));
  try {
    const library = '/nix/store/mock-glibc/lib/libc.so.6';
    const bundlePath = path.join(dir, 'runtime.js'); await buildMacro(bundlePath, [library]);
    const source = await readFile(bundlePath, 'utf8'), file = '/文稿/其它汇报.pptx', original = await fixture();
    const m = fakeLibc([[file, original], ['/proc/sys/kernel/random/uuid',
      Buffer.from('eb0f4276-5626-4b15-a499-17bfbb804bc5\n')]]);
    const context = vm.createContext({ Application: {
      ActivePresentation: { FullName: '/工具/编辑器.pptm' },
      Presentations: { Count: 0, Open() { throw new Error('must not open a document'); } },
      FileDialog: () => ({ Filters: { Clear() {}, Add() {} }, Show: () => -1, SelectedItems: { Item: () => file } }),
    }, ffi: m.ffi });
    // Compare signatures across VM realms by value.
    context.ffi = { LoadLibrary(name, signature) {
      if (name !== library) throw new Error('Library not found');
      return m.ffi.LoadLibrary(name, JSON.parse(JSON.stringify(signature)));
    } };
    vm.runInContext('globalThis.setImmediate = callback => Promise.resolve().then(callback);' + source, context);
    const hostApp = context.NativeTableStyles.installRuntime(context.Application);
    assert.equal(hostApp.FileSystem.library, library);
    const randomBytes = new Uint8Array(20);
    context.crypto.getRandomValues(randomBytes);
    assert.equal(Buffer.from(randomBytes).toString('hex'), 'eb0f427656264b15a49917bfbb804bc5eb0f4276');
    const { FileHost } = await import('../src/file-host.js');
    const host = new FileHost(hostApp);
    await host.openFile(); const info = await host.inspect();
    const result = await host.apply(plan({ path: file, sourceXML: info.styles[0].xml }));
    const output = await openDeck(m.files.get(file));
    assert.equal(output.defaultID, result.id);
    assert.equal(readPart(output.styles[0].xml, 'firstRow').fill, '#0030C2');
    assert.deepEqual(m.files.get(result.backup), Buffer.from(original));
    assert.equal(m.fds.size, 0);
    assert.equal(context.Application.ActivePresentation.FullName, '/工具/编辑器.pptm');
  } finally { await rm(dir, { recursive: true, force: true }); }
});
