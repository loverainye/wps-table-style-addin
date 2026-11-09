import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';
import { fixture, plan } from './fixtures.mjs';
import { binaryString } from '../src/host.js';
import { FileHost } from '../src/file-host.js';
import { openDeck, readPart } from '../src/styles.js';

async function mock() {
  const path = '/docs/其它季度汇报.pptx', source = await fixture(), files = new Map([[path, source]]);
  let selection = path, cancel = false, failWrite = false;
  const macro = { FullName: '/tools/editor.pptm' };
  const app = {
    ActivePresentation: macro,
    Presentations: { Count: 1, Item: () => macro, Open() { throw new Error('must not open'); } },
    FileDialog: kind => {
      assert.equal(kind, 3);
      return { Filters: { Clear() {}, Add() {} }, Show: () => cancel ? 0 : -1, SelectedItems: { Item: () => selection } };
    },
    FileSystem: {
      Exists: p => files.has(p), Remove: p => files.delete(p),
      readAsBinaryString: p => binaryString(files.get(p)),
      writeAsBinaryString: (p, data) => {
        if (p === path && failWrite) { failWrite = false; files.set(p, new Uint8Array([0])); throw new Error('write failure'); }
        files.set(p, Uint8Array.from(data, c => c.charCodeAt(0)));
      },
    },
  };
  return { app, path, source, files, host: new FileHost(app), select(p) { selection = p; }, cancel() { cancel = true; }, fail() { failWrite = true; } };
}

test('macro file picker loads and edits a closed PPTX without opening it in WPS', async () => {
  const m = await mock();
  assert.equal(await m.host.inspect(), null);
  assert.equal(await m.host.call('openFile'), true);
  const info = await m.host.call('inspect');
  assert.equal(info.path, m.path);
  const result = await m.host.call('apply', plan({ path: m.path, sourceXML: info.styles[0].xml }));
  const output = await openDeck(m.files.get(m.path));
  assert.equal(output.defaultID, result.id);
  assert.equal(readPart(output.styles[0].xml, 'firstRow').fill, '#0030C2');
  assert.deepEqual(m.files.get(result.backup), m.source);
  assert.equal(m.app.ActivePresentation.FullName, '/tools/editor.pptm');
});

test('cancelling a macro file picker keeps the current target', async () => {
  const m = await mock(); await m.host.openFile(); m.cancel();
  assert.equal(await m.host.openFile(), false);
  assert.equal(m.host.path, m.path);
});

test('macro writes restore the original after a partial write without opening documents', async () => {
  const m = await mock(); await m.host.openFile(); const info = await m.host.inspect(); m.fail();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml })), /write failure/);
  assert.deepEqual(m.files.get(m.path), m.source);
  assert.equal(m.host.busy, false);
});

test('closed-file editor rejects selected-table scope without a native selection', async () => {
  const m = await mock(); await m.host.openFile(); const info = await m.host.inspect();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml, scope: 'selected' })), /选中表格/);
  assert.deepEqual(m.files.get(m.path), m.source);
});

test('macro launcher uses the embedded native form without web dialog or extraction APIs', async () => {
  const source = (await readFile(new URL('../macro/launcher.js', import.meta.url), 'utf8'))
    .replace('__TABLE_STYLES_VERSION__', '"test"');
  const calls = [];
  const context = {
    Application: {},
    TableStylesFormEditor: () => calls.push('initialize'),
    TableStyleForm: { Width: 700, Height: 762, OpenFile: { Caption: '打开 PPTX' }, Show: () => { calls.push('show'); return true; } },
  };
  vm.runInNewContext(source, context);
  assert.doesNotMatch(source, /^(?:var|let|const)\s/m, 'macro entry must contain only function declarations at module scope');
  assert.doesNotMatch(source, /ShowDialog|CreateTaskPane|writeFileString|GetTempPath/);
  assert.equal(context.TableStylesVersion(), 'test');
  assert.equal(context.OpenTableStyleEditor(), true);
  assert.deepEqual(calls, ['initialize', 'show']);
  delete context.TableStyleForm;
  assert.throws(() => context.OpenTableStyleEditor(), /缺少原生窗体/);
});
