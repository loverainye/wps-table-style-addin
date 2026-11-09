import test from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { createPlatformServices, macroApplication } from '../macro/platform-services.js';

const v4 = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/;
function native() {
  const fs = {
    files: new Map(), readAsBinaryString(p) { return this.files.get(p); },
    writeAsBinaryString(p, data) { this.files.set(p, data); },
    existsSync(p) { return this.files.has(p); }, unlinkSync(p) { this.files.delete(p); },
  };
  return fs;
}

function forbidSystemAPIs(root) {
  for (const name of ['ffi', 'FFI', 'Platform', 'process', 'navigator', 'Env']) {
    Object.defineProperty(root, name, { get() { throw new Error('must not access '+name); } });
  }
  return root;
}

test('missing or incomplete WPS file access reports the requirement without probing a system API', () => {
  for (const app of [{}, { FileSystem: { readAsBinaryString() {} } }, { FileSystem: { writeAsBinaryString() {} } }]) {
    assert.throws(() => createPlatformServices(app, forbidSystemAPIs({})), /原生二进制文件读写接口/);
  }
  const root = forbidSystemAPIs({ FileSystem: native() });
  assert.equal(createPlatformServices({ FileSystem: {} }, root).fileBackend, 'WPS FileSystem');
});

test('WPS files preserve binary bytes and receiver binding without platform or temporary directory access', () => {
  const fs = native(), services = createPlatformServices({ FileSystem: fs }, forbidSystemAPIs({}));
  const data = String.fromCharCode(...Array.from({length: 256}, (_, i) => i)).repeat(550);
  services.files.writeAsBinaryString('中文.pptx', data);
  assert.equal(services.files.readAsBinaryString('中文.pptx'), data);
  fs.writeAsBinaryString('text', Buffer.from('中文🙂').toString('latin1'));
  assert.equal(services.files.readFileString('text'), '中文🙂');
  assert.equal(services.tempPath, undefined);
});

test('file errors propagate from WPS without trying another file backend', () => {
  const fs = native();
  fs.readAsBinaryString = () => { throw new Error('native read failed'); };
  fs.writeAsBinaryString = () => { throw new Error('native write failed'); };
  const services = createPlatformServices({ FileSystem: fs }, forbidSystemAPIs({}));
  assert.throws(() => services.files.readAsBinaryString('file'), /native read failed/);
  assert.throws(() => services.files.writeAsBinaryString('file', ''), /native write failed/);
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

test('macro application facade preserves live document proxies and never adds properties to WPS', () => {
  let active = { Name: 'A' };
  const app = Object.freeze({ get ActivePresentation() { return active; }, FileDialog() { assert.equal(this, app); return 3; } });
  const root = { FileSystem: native(), crypto: { randomUUID }, Platform: { OS: () => 'Windows' } };
  const adapted = macroApplication(app, root);
  assert.equal(adapted.FileDialog(3), 3); active = { Name: 'B' };
  assert.equal(adapted.ActivePresentation, active); assert.equal(app.TableStylePlatform, undefined);
  assert.match(adapted.TableStylePlatform.uuid(), v4);
});
