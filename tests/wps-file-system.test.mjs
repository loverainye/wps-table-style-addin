import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { buildUserFormMacro } from '../scripts/build.mjs';
import { fixture, plan } from './fixtures.mjs';
import { openDeck, readPart } from '../src/styles.js';

// A WPS FileSystem contract backed by real disk files for CI on all three OSes.
// This does not stand in for a WPS client integration test.
test('bundled UserForm macro edits a real PPTX through the WPS file contract on this OS', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'wps-files-'));
  try {
    const runtime = path.join(dir, 'runtime.js'), file = path.join(dir, '中文 目标文稿.pptx');
    await buildUserFormMacro(runtime);
    const bundle = await readFile(runtime, 'utf8');
    assert.doesNotMatch(bundle, /LoadLibrary|libc\.so|libSystem|kernel32|__WPS_LIBC_PATHS__|\/proc\/|GetTempPath/);
    const original = await fixture(); await writeFile(file, original);
    const writes = [], fs = {
      readAsBinaryString(p) { return readFileSync(p).toString('latin1'); },
      writeAsBinaryString(p, s) { writes.push(p); writeFileSync(p, Buffer.from(s, 'latin1')); },
    };
    for (const globalFiles of [true, false]) {
      await writeFile(file, original); writes.length = 0;
      const app = { Presentations: { Count: 0 }, ActivePresentation: { FullName: path.join(dir, 'tool.pptm') } };
      const globals = { Application: app };
      if (globalFiles) globals.FileSystem = fs; else app.FileSystem = fs;
      for (const name of ['ffi', 'FFI', 'Platform', 'process', 'Env']) {
        Object.defineProperty(globals, name, { get() { throw new Error('forbidden API '+name); } });
      }
      const context = vm.createContext(globals);
      vm.runInContext('globalThis.setImmediate = callback => Promise.resolve().then(callback);'+bundle, context);
      const runtimeAPI = context.NativeUserFormTableStyles, adapted = runtimeAPI.installRuntime(app);
      const host = new runtimeAPI.UserFormEditor(adapted, {}).host; host.path = file;
      const info = await host.inspect(), result = await host.apply(plan({path: file, sourceXML: info.styles[0].xml}));
      const output = await openDeck(await readFile(file));
      assert.equal(readPart(output.styles.find(s => s.id === result.id).xml, 'firstRow').fill, '#0030C2');
      assert.deepEqual(await readFile(result.backup), Buffer.from(original));
      assert.deepEqual(writes, [result.backup, file]);
      assert.equal(adapted.TableStylePlatform.fileBackend, globalFiles ? 'WPS FileSystem' : 'WPS Application.FileSystem');
      assert.equal(adapted.TableStylePlatform.uuidBackend, 'JS Math.random');
    }
  } finally { await rm(dir, {recursive: true, force: true}); }
});
