import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import koffi from 'koffi';
import { createPlatformServices } from '../macro/platform-services.js';
import { buildMacro } from '../scripts/build.mjs';
import { fixture, plan } from './fixtures.mjs';
import { openDeck, readPart } from '../src/styles.js';

// Exercise the actual OS libraries in CI. Koffi is a test-only bridge; release
// PPTMs use WPS ffi.LoadLibrary and contain no Node or native npm dependencies.
const types = { pointer: 'void *', string: 'str', u16string: 'str16', int32: 'int32_t', uint32: 'uint32_t', int64: 'int64_t', uint64: 'uint64_t' };
function nativeRoot() {
  return {
    Platform: { OS: () => os.type(), PointerSize: () => koffi.sizeof('void *') },
    ffi: { LoadLibrary(name, signature) {
      const library = koffi.load(name), api = {};
      for (const [symbol, decl] of Object.entries(signature)) {
        const args = decl.parameters.map(t => types[t]);
        const fn = decl.abi === 'stdcall' ? library.func('__stdcall', symbol, types[decl.returnType], args) : library.func(symbol, types[decl.returnType], args);
        api[symbol] = (...values) => {
          const converted = values.map(v => v?.nativePointer || (Object.prototype.toString.call(v) === '[object ArrayBuffer]' ? new Uint8Array(v) : v));
          const result = fn(...converted);
          return decl.returnType === 'pointer' && result ? { nativePointer: result, ToBigInt: () => koffi.address(result) } : result;
        };
      }
      return api;
    } },
  };
}

test('real OS fallback roundtrips every byte, truncates files, supplies temp paths and unique UUIDs', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'wps-native-api-'));
  try {
    const services = createPlatformServices({}, nativeRoot()), fs = services.files;
    const file = path.join(dir, '中文空格 文件.pptx');
    const data = Buffer.from(Array.from({ length: 140000 }, (_, i) => i % 256));
    fs.writeAsBinaryString(file, data.toString('latin1'));
    assert.equal(fs.readAsBinaryString(file), data.toString('latin1'));
    fs.writeAsBinaryString(file, 'short\0\xff');
    assert.equal(fs.readAsBinaryString(file), 'short\0\xff');
    fs.writeAsBinaryString(file, ''); assert.equal(fs.readAsBinaryString(file), '');
    assert.equal(fs.Exists(file), true); fs.Remove(file); assert.equal(fs.Exists(file), false);
    assert.ok(path.isAbsolute(services.tempPath()));
    const ids = new Set(Array.from({ length: 200 }, () => services.uuid()));
    assert.equal(ids.size, 200);
    for (const id of ids) assert.match(id, /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
    assert.throws(() => fs.readAsBinaryString(file), /失败/);
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('actual bundled macro on this OS edits a PPTX through system FFI, verifies backup, and stays offline', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'wps-native-deck-'));
  try {
    const runtime = path.join(dir, 'macro.js'), file = path.join(dir, '目标文稿.pptx');
    await buildMacro(runtime, []);
    const root = nativeRoot(), original = await fixture();
    const sourceServices = createPlatformServices({}, root);
    sourceServices.files.writeAsBinaryString(file, original.reduce((s, b) => s + String.fromCharCode(b), ''));
    const app = { Presentations: { Count: 0 }, ActivePresentation: { FullName: path.join(dir, 'tool.pptm') } };
    const context = vm.createContext({ ...root, Application: app });
    vm.runInContext('globalThis.setImmediate = callback => Promise.resolve().then(callback);' + await readFile(runtime, 'utf8'), context);
    const adapted = context.NativeTableStyles.installRuntime(app);
    const host = new context.NativeTableStyles.NativeEditor(adapted, {}).host; host.path = file;
    const info = await host.inspect();
    const result = await host.apply(plan({ path: file, sourceXML: info.styles[0].xml }));
    const output = await openDeck(await readFile(file));
    assert.equal(output.defaultID, result.id);
    assert.equal(readPart(output.styles.find(s => s.id === result.id).xml, 'firstRow').fill, '#0030C2');
    assert.deepEqual(await readFile(result.backup), Buffer.from(original));
    assert.equal(vm.runInContext('typeof crypto', context), 'undefined');
  } finally { await rm(dir, { recursive: true, force: true }); }
});
