import test from 'node:test';
import assert from 'node:assert/strict';
import { fixture, plan } from './fixtures.mjs';
import { WpsHost, binaryString } from '../src/host.js';
import { openDeck } from '../src/styles.js';

async function mock({ openFails = false, writeFails = false, masters = [] } = {}) {
  const source = await fixture({ masters }), path = '/docs/current.pptx', files = new Map([[path, source]]), calls = [];
  let doc, memory = source, failOpen = openFails, failWrite = writeFails;
  const app = { Env: { GetTempPath: () => '/tmp/' }, ActiveWindow: { View: { Slide: { SlideIndex: 2 }, GotoSlide: i => calls.push(['goto', i]) } } };
  const createDoc = () => ({
    FullName: path, Name: 'current.pptx', Path: '/docs', ReadOnly: 0, Saved: -1,
    SaveCopyAs: tmp => { files.set(tmp, memory); calls.push(['snapshot', tmp]); },
    Save: () => { files.set(path, memory); doc.Saved = -1; calls.push(['save']); },
    Close: () => { app.ActivePresentation = null; calls.push(['close']); },
  });
  doc = app.ActivePresentation = createDoc();
  app.FileSystem = {
    Exists: p => files.has(p),
    Remove: p => files.delete(p),
    readAsBinaryString: p => { if (!files.has(p)) throw new Error('missing ' + p); return binaryString(files.get(p)); },
    writeAsBinaryString: (p, value) => {
      if (p === path && failWrite) { failWrite = false; files.set(p, new Uint8Array([0])); throw new Error('磁盘写入失败'); }
      files.set(p, Uint8Array.from(value, c => c.charCodeAt(0) & 255));
    },
  };
  app.Presentations = { Open: p => {
    calls.push(['open', p]);
    if (failOpen) { failOpen = false; throw new Error('重新打开失败'); }
    memory = files.get(p); doc = app.ActivePresentation = createDoc();
  } };
  return { app, path, files, calls, source, host: new WpsHost(app), get memory() { return memory; }, set memory(value) { memory = value; }, reopen() { memory = files.get(path); doc = app.ActivePresentation = createDoc(); } };
}
test('inspect takes a temporary snapshot without saving or closing the current file', async () => {
  const m = await mock(), info = await m.host.inspect();
  assert.equal(info.path, m.path); assert.equal(info.tableCount, 2);
  assert.equal(m.calls.filter(c => c[0] === 'save' || c[0] === 'close').length, 0);
  assert.equal(m.files.size, 1);
  assert.deepEqual(m.files.get(m.path), m.source);
});
test('writes the same path with a verified backup and leaves the document open', async () => {
  const m = await mock(), info = await m.host.inspect();
  const result = await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  assert.equal(result.path, m.path);
  assert.deepEqual(m.files.get(result.backup), m.source);
  assert.notDeepEqual(m.files.get(m.path), m.source);
  assert.equal((await openDeck(m.files.get(m.path))).defaultID, result.id);
  assert.deepEqual(m.calls.filter(c => ['save', 'close', 'open', 'goto'].includes(c[0])), [['save']]);
  assert.equal(m.files.size, 2);
});
test('selected scope follows the selection at save time rather than the editor inspection', async () => {
  const m = await mock();
  let selectedShape = '5';
  m.app.ActiveWindow.View.Slide.SlideID = 256;
  m.app.ActiveWindow.Selection = { ShapeRange: { Count: 1, Item: () => ({ HasTable: true, Id: selectedShape }) } };
  const info = await m.host.inspect();
  assert.equal(info.targets[0].slideID, '256');
  m.app.ActiveWindow.View.Slide.SlideID = 300;
  await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml, scope: 'selected', targets: info.targets }));
  const deck = await openDeck(m.files.get(m.path));
  assert.equal(deck.tables[0].styleID, info.styles[0].id);
  assert.notEqual(deck.tables[1].styleID, info.styles[0].id);
});
test('cleared native selection cannot apply to a stale table selected before opening the editor', async () => {
  const m = await mock(), info = await m.host.inspect();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml, scope: 'selected', targets: [{ slideID: '256', shapeID: '5' }] })), /选中表格/);
  assert.equal(m.calls.some(c => c[0] === 'close' || c[0] === 'save'), false);
  assert.deepEqual(m.files.get(m.path), m.source);
});
test('restores the saved current file after a partial write without closing or reopening', async () => {
  const m = await mock({ writeFails: true }), info = await m.host.inspect();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml })), /已恢复修改前的文件/);
  assert.deepEqual(m.files.get(m.path), m.source);
  assert.equal(m.app.ActivePresentation.FullName, m.path);
  assert.equal(m.host.busy, false);
});
test('a different active document is rejected before saving or writing', async () => {
  const m = await mock();
  await assert.rejects(() => m.host.apply(plan({ path: '/other.pptx' })), /切换/);
  assert.equal(m.calls.length, 0); assert.equal(m.files.size, 1);
});
test('a stale style is rejected before closing the current file', async () => {
  const m = await mock();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: '<stale/>' })), /其他操作/);
  assert.equal(m.calls.some(c => c[0] === 'close' || c[0] === 'save'), false);
});
test('read-only or unsupported WPS APIs fail before any file mutation', async () => {
  const m = await mock(); m.app.ActivePresentation.ReadOnly = -1;
  await assert.rejects(() => m.host.inspect(), /只读/);
  m.app.ActivePresentation.ReadOnly = 0; delete m.app.FileSystem.writeAsBinaryString;
  await assert.rejects(() => m.host.inspect(), /二进制/);
  assert.equal(m.calls.length, 0);
});
test('preserves edits made to table text after the editor was opened', async () => {
  const m = await mock(), info = await m.host.inspect();
  const deck = await openDeck(m.files.get(m.path));
  const slide = await deck.zip.file('ppt/slides/slide1.xml').async('string');
  deck.zip.file('ppt/slides/slide1.xml', slide.replace('UNCHANGED TEXT 1', 'LATEST USER EDIT'));
  m.memory = await deck.zip.generateAsync({ type: 'uint8array' }); m.app.ActivePresentation.Saved = 0;
  await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  const output = await openDeck(m.files.get(m.path));
  assert.match(await output.zip.file('ppt/slides/slide1.xml').async('string'), /LATEST USER EDIT/);
});
test('does not overwrite edits made during the asynchronous apply operation', async () => {
  const m = await mock(), info = await m.host.inspect();
  const changed = await openDeck(m.source);
  changed.zip.file('customXml/user-edit.xml', '<latest/>');
  const latest = await changed.zip.generateAsync({ type: 'uint8array' });
  const save = m.app.ActivePresentation.Save;
  m.app.ActivePresentation.Save = () => { save(); m.files.set(m.path, latest); };
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml })), /内容发生变化/);
  assert.deepEqual(m.files.get(m.path), latest);
  assert.equal(m.calls.some(c => c[0] === 'close'), false);
});
test('never calls native close, open or slide navigation while saving styles', async () => {
  const m = await mock(), info = await m.host.inspect(), document = m.app.ActivePresentation;
  m.app.ActivePresentation.Close = () => { throw new Error('must not close'); };
  m.app.Presentations.Open = () => { throw new Error('must not open'); };
  await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  assert.equal(m.app.ActivePresentation, document);
  assert.notDeepEqual(m.files.get(m.path), m.source);
  assert.equal(m.calls.some(c => ['close', 'open', 'goto'].includes(c[0])), false);
});
test('repeated saves read the last written styles while native memory retains the old XML', async () => {
  const m = await mock(), info = await m.host.inspect();
  const first = await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  const written = await m.host.inspect(), style = written.styles.find(s => s.id === first.id);
  assert.ok(style);
  assert.deepEqual(m.memory, m.source);
  const second = await m.host.apply(plan({ path: m.path, id: style.id, name: style.name, sourceXML: style.xml, patches: { firstRow: { fill: '#228822' } } }));
  const deck = await openDeck(m.files.get(m.path));
  assert.equal(deck.defaultID, second.id);
  assert.notEqual(first.id, second.id);
  assert.equal(m.calls.filter(c => c[0] === 'save').length, 1);
  assert.deepEqual(m.memory, m.source);
});
test('refuses to overwrite pending styles when old native memory has new unsaved edits', async () => {
  const m = await mock(), info = await m.host.inspect();
  const first = await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  const written = await m.host.inspect(), style = written.styles.find(s => s.id === first.id), bytes = m.files.get(m.path);
  const memory = await openDeck(m.memory); memory.zip.file('customXml/latest.xml', '<latest/>');
  m.memory = await memory.zip.generateAsync({ type: 'uint8array' }); m.app.ActivePresentation.Saved = 0;
  await assert.rejects(() => m.host.apply(plan({ path: m.path, id: style.id, sourceXML: style.xml })), /仍保留旧内容/);
  assert.deepEqual(m.files.get(m.path), bytes);
  assert.equal(m.calls.filter(c => c[0] === 'save').length, 1);
});
test('a manual reopen restores normal saves and retains the already written styles', async () => {
  const m = await mock(), info = await m.host.inspect();
  await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  m.reopen();
  const loaded = await m.host.inspect(), style = loaded.styles.find(s => s.id === loaded.defaultID);
  await m.host.apply(plan({ path: m.path, id: style.id, name: style.name, sourceXML: style.xml, patches: {} }));
  assert.equal(m.calls.filter(c => c[0] === 'save').length, 2);
  assert.equal((await openDeck(m.files.get(m.path))).defaultID, style.id);
});

test('inspects and applies document-wide default font size with a verified backup', async () => {
  const m = await mock({ masters: ['<a:defPPr><a:defRPr sz="1800"/></a:defPPr>', null] });
  const info = await m.host.inspect();
  assert.deepEqual(info.fontDefaults, { masterCount: 2, points: null, mixed: true });
  const result = await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml, name: info.styles[0].name, patches: {}, scope: 'none', defaultFontSize: 26 }));
  assert.equal(result.changedMasters, 2);
  assert.equal(result.changedTables, 0);
  assert.equal(result.id, info.styles[0].id);
  assert.deepEqual(m.files.get(result.backup), m.source);
  assert.deepEqual((await m.host.inspect()).fontDefaults, { masterCount: 2, points: 26, mixed: false });
});

test('invalid font defaults stop before saving or closing the document', async () => {
  const m = await mock({ masters: [null] }), info = await m.host.inspect();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml, defaultFontSize: 0 })), /默认字号/);
  assert.equal(m.calls.some(c => c[0] === 'close' || c[0] === 'save'), false);
  assert.deepEqual(m.files.get(m.path), m.source);
});

test('a saved absolute FullName remains usable when WPS returns an empty Path', async () => {
  const m = await mock(); m.app.ActivePresentation.Path = '';
  const info = await m.host.inspect();
  assert.equal(info.path, m.path); assert.equal(info.tableCount, 2);
  assert.equal(m.calls.some(c => c[0] === 'save' || c[0] === 'close'), false);
});

test('unrecognized current documents report actual WPS path values before any mutation', async () => {
  const m = await mock();
  for (const fullName of ['current.pptx', '/tools/编辑器.pptm', 'https://example.test/current.pptx', '']) {
    m.app.ActivePresentation.FullName = fullName;
    await assert.rejects(() => m.host.inspect(), error => /FullName=/.test(error.message) && error.message.includes(fullName || '空'));
  }
  assert.equal(m.calls.length, 0);
});
