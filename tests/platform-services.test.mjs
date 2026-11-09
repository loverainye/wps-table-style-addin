import test from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { createPlatformServices, detectPlatform, macroApplication } from '../macro/platform-services.js';
import { createMacFileSystem } from '../macro/platform/macos.js';
import { createWindowsFileSystem } from '../macro/platform/windows.js';

const v4 = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;
function native() {
  const fs = {
    files: new Map(), readAsBinaryString(p) { return this.files.get(p); },
    writeAsBinaryString(p, data) { this.files.set(p, data); },
    existsSync(p) { return this.files.has(p); }, unlinkSync(p) { this.files.delete(p); },
    tmpdir: () => 'C:\\用户\\Temp\\',
  };
  return fs;
}

test('WPS Platform OS names are normalized without assuming Linux for unknown systems', () => {
  for (const [name, expected] of [['Linux','linux'], ['Darwin','macos'], ['Windows','windows'], ['win32','windows'], ['MacIntel','macos'], ['unknown','unsupported']]) {
    assert.equal(detectPlatform({ Platform: { OS: () => name } }), expected);
  }
  assert.equal(detectPlatform({}), null);
});

test('global macro FileSystem and JS randomUUID work on every OS with no FFI or proc reads', () => {
  for (const os of ['Linux', 'Darwin', 'Windows']) {
    let calls = 0;
    const fs = native(), app = Object.freeze({ get FileSystem() { throw new Error('not in macro host'); } });
    const root = { FileSystem: fs, Platform: { OS: () => os }, crypto: { randomUUID() { calls++; return randomUUID(); } }, ffi: { LoadLibrary() { throw new Error('must not load FFI'); } } };
    const services = createPlatformServices(app, root);
    services.files.writeAsBinaryString('file', '\0\xff中文');
    assert.equal(services.files.readAsBinaryString('file'), '\0\xff中文');
    assert.equal(services.files.Exists('file'), true);
    services.files.Remove('file'); assert.equal(services.files.Exists('file'), false);
    assert.equal(services.tempPath(), 'C:\\用户\\Temp\\');
    const ids = new Set(Array.from({ length: 200 }, () => services.uuid()));
    assert.equal(ids.size, 200); assert.equal(calls, 200);
    for (const id of ids) assert.match(id, v4);
    assert.equal(services.fileBackend, 'WPS FileSystem');
    assert.equal(services.uuidBackend, 'JS crypto.randomUUID');
  }
});

test('JS Crypto UUID generation uses the correct receiver and falls back to getRandomValues', () => {
  const crypto = { getRandomValues(bytes) { assert.equal(this, crypto); bytes.fill(255); return bytes; } };
  const services = createPlatformServices({ FileSystem: native(), Env: { GetTempPath: () => '/native/temp' } }, { crypto });
  assert.equal(services.uuid(), '{FFFFFFFF-FFFF-4FFF-BFFF-FFFFFFFFFFFF}');
  assert.equal(services.uuidBackend, 'JS crypto.getRandomValues');
  assert.equal(services.tempPath(), '/native/temp');
  const preferred = createPlatformServices({ FileSystem: native() }, { crypto: { randomUUID, getRandomValues() { throw new Error('randomUUID must win'); } } });
  assert.match(preferred.uuid(), v4);
});

test('without Crypto, UUIDs are generated in JS without files, FFI, or operating-system knowledge', () => {
  for (const os of ['Linux', 'Darwin', 'Windows', 'Other', '']) {
    let calls = 0;
    const fs = native(); fs.readFileString = () => { throw new Error('must not read system UUID files'); };
    const services = createPlatformServices({ FileSystem: fs }, {
      Platform: { OS: () => os }, Math: { random() { calls++; return Math.random(); } },
      ffi: { LoadLibrary() { throw new Error('UUID must not use FFI'); } },
    });
    for (let i = 0; i < 200; i++) assert.match(services.uuid(), v4);
    assert.equal(calls, 200 * 16); assert.equal(services.uuidBackend, 'JS Math.random');
  }
});

test('Darwin file fallback uses fixed-signature creat to avoid Apple arm64 variadic open ABI', () => {
  const calls = [];
  const api = { open: () => 3, creat(p, mode) { calls.push([p, mode]); return 3; }, close: () => 0, fsync: () => 0,
    write: (fd, ptr, size) => size, read: () => 0n, access: () => 0, unlink: () => 0, getenv: () => '/var/folders/wps/temp/' };
  const files = createMacFileSystem({ LoadLibrary(name, signature) {
    assert.equal(name, '/usr/lib/libSystem.B.dylib');
    assert.deepEqual(signature.open.parameters, ['string', 'int32']);
    assert.deepEqual(signature.creat.parameters, ['string', 'uint32']);
    return api;
  } });
  files.writeAsBinaryString('/文稿/中文.pptx', 'abc');
  assert.deepEqual(calls, [['/文稿/中文.pptx', 0o600]]);
  assert.equal(files.tempPath(), '/var/folders/wps/temp/');
  assert.throws(() => files.writeAsBinaryString('relative', ''), /绝对路径/);
});

function windowsMock() {
  const files = new Map(), handles = new Set(), fail = {};
  const api = {
    CreateFileW(path, access, share, security, creation) {
      if (fail.open || (creation === 3 && !files.has(path))) return { ToBigInt: () => 0xffffffffffffffffn };
      if (creation === 2) files.set(path, Buffer.alloc(0));
      const h = { path, at: 0, ToBigInt: () => 100n }; handles.add(h); return h;
    },
    ReadFile(h, pointer, length, count) {
      if (fail.read) return 0;
      const data = files.get(h.path), size = Math.min(1111, length, data.length - h.at);
      new Uint8Array(pointer).set(data.subarray(h.at, h.at + size)); new Uint32Array(count)[0] = size; h.at += size; return 1;
    },
    WriteFile(h, pointer, length, count) {
      if (fail.write) return 0;
      const size = fail.zeroWrite ? 0 : Math.min(1111, length);
      files.set(h.path, Buffer.concat([files.get(h.path), Buffer.from(pointer).subarray(0, size)]));
      new Uint32Array(count)[0] = size; return 1;
    },
    FlushFileBuffers: () => fail.flush ? 0 : 1,
    CloseHandle(h) { assert.ok(handles.delete(h)); return fail.close ? 0 : 1; },
    GetFileAttributesW: p => files.has(p) ? 128 : 0xffffffff,
    DeleteFileW: p => Number(files.delete(p)),
    GetTempPathW(length, pointer) { const p = 'C:\\用户\\Temp\\'; new Uint16Array(pointer).set([...p].map(c => c.charCodeAt(0))); return p.length; },
  };
  const root = { Platform: { PointerSize: () => 4 }, ffi: { LoadLibrary(name, sig) {
    assert.equal(name, 'kernel32.dll'); assert.equal(sig.CreateFileW.parameters[0], 'u16string');
    for (const decl of Object.values(sig)) assert.equal(decl.abi, 'stdcall');
    return api;
  } } };
  return { root, files, handles, fail };
}

test('Windows fallback preserves binary bytes, Chinese/UNC paths and partial I/O; closes failed handles', () => {
  const m = windowsMock(), fs = createWindowsFileSystem(m.root);
  const data = Buffer.from(Array.from({ length: 140000 }, (_, i) => i % 256));
  for (const path of ['C:\\文稿\\新文件.pptx', '\\\\server\\共享\\新文件.pptx']) {
    fs.writeAsBinaryString(path, data.toString('latin1'));
    assert.equal(fs.readAsBinaryString(path), data.toString('latin1'));
    assert.equal(fs.Exists(path), true); fs.Remove(path); assert.equal(fs.Exists(path), false);
  }
  assert.equal(fs.tempPath(), 'C:\\用户\\Temp\\');
  assert.equal(m.handles.size, 0);
  for (const path of ['relative', 'C:relative', 'C:\\a\0b', null]) assert.throws(() => fs.writeAsBinaryString(path, ''), /绝对路径/);
  for (const key of ['read', 'write', 'zeroWrite', 'flush', 'close', 'open']) {
    m.files.set('C:\\test', Buffer.from('abc'));
    m.fail[key] = true;
    assert.throws(() => key === 'read' ? fs.readAsBinaryString('C:\\test') : fs.writeAsBinaryString('C:\\test', 'abc'), /失败/);
    assert.equal(m.handles.size, 0, key); m.fail[key] = false;
  }
});

test('macro application facade preserves live document proxies and never adds properties to WPS', () => {
  let active = { Name: 'A' };
  const app = Object.freeze({ get ActivePresentation() { return active; }, FileDialog() { assert.equal(this, app); return 3; } });
  const root = { FileSystem: native(), crypto: { randomUUID }, Platform: { OS: () => 'Windows' } };
  const adapted = macroApplication(app, root);
  assert.equal(adapted.FileDialog(3), 3); active = { Name: 'B' };
  assert.equal(adapted.ActivePresentation, active); assert.equal(app.TableStylePlatform, undefined);
  assert.match(adapted.TableStylePlatform.uuid(), v4);
});
