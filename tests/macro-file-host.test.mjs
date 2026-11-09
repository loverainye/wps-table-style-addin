import test from 'node:test';
import assert from 'node:assert/strict';
import { fixture, plan } from './fixtures.mjs';
import { binaryString } from '../src/host.js';
import { MacroFileHost } from '../macro/file-host.js';
import { openDeck, readPart } from '../src/styles.js';

async function mock(open = false) {
  const path = '/docs/当前文稿.pptx', original = await fixture(), files = new Map([[path, original]]), writes = [];
  const doc = { FullName: path, Saved: true, ReadOnly: false,
    Save() { throw new Error('must not save the in-memory presentation'); },
    SaveCopyAs() { throw new Error('must not create a snapshot'); },
    Close() { throw new Error('must not close'); },
  };
  const docs = open ? [doc] : [];
  let shapeID = '5', fail = false;
  const app = {
    ActivePresentation: open ? doc : { FullName: '/tools/editor.pptm' },
    Presentations: { get Count() { return docs.length; }, Item: i => docs[i-1], Open() { throw new Error('must not open'); } },
    ActiveWindow: { View: { Slide: { SlideID: '256' } }, Selection: { ShapeRange: { Count: 1, Item: () => ({ HasTable: true, get Id() { return shapeID; } }) } } },
    FileSystem: {
      readAsBinaryString: p => binaryString(files.get(p)),
      writeAsBinaryString(p, data) {
        writes.push(p);
        if (p === path && fail) { fail = false; files.set(p, new Uint8Array([0])); throw new Error('partial write'); }
        files.set(p, Uint8Array.from(data, c => c.charCodeAt(0)));
      },
    },
    get Env() { throw new Error('must not access temporary directories'); },
  };
  const host = new MacroFileHost(app); host.path = path;
  return { path, original, files, writes, app, doc, docs, host, select(id) { shapeID = id; }, fail() { fail = true; } };
}

test('macro edits an open saved file directly, keeps the window open and uses no snapshots or temp directory', async () => {
  const m = await mock(true), info = await m.host.inspect();
  const result = await m.host.apply(plan({ path: m.path, sourceXML: info.styles[0].xml }));
  const deck = await openDeck(m.files.get(m.path));
  assert.equal(readPart(deck.styles.find(s => s.id === result.id).xml, 'firstRow').fill, '#0030C2');
  assert.deepEqual(m.files.get(result.backup), m.original);
  assert.deepEqual(m.writes, [result.backup, m.path]);
  assert.equal(m.app.ActivePresentation, m.doc);
  assert.equal(m.doc.Saved, true);
});

test('unsaved or readonly open documents stop before any file is written', async () => {
  const m = await mock(true);
  for (const value of [false, 0, undefined]) {
    m.doc.Saved = value;
    await assert.rejects(() => m.host.inspect(), /未保存/);
    await assert.rejects(() => m.host.apply(plan({ path: m.path })), /未保存/);
  }
  m.doc.Saved = -1; m.doc.ReadOnly = true;
  await assert.rejects(() => m.host.apply(plan({ path: m.path })), /只读/);
  assert.deepEqual(m.writes, []);
  assert.deepEqual(m.files.get(m.path), m.original);
});

test('selected scope uses the selection at save time and rejects closed files', async () => {
  const m = await mock(true); await m.host.inspect(); m.select('999');
  await assert.rejects(() => m.host.apply(plan({ path: m.path, scope: 'selected' })), /选中/);
  assert.deepEqual(m.writes, []);
  m.select('5');
  const result = await m.host.apply(plan({ path: m.path, scope: 'selected' }));
  assert.equal(result.changedTables, 1);
  m.docs.length = 0;
  await assert.rejects(() => m.host.apply(plan({ path: m.path, scope: 'selected' })), /选中/);
});

test('closed files support repeated saves and recover the original after a partial target write', async () => {
  const m = await mock();
  for (const fill of ['#112233', '#445566']) {
    const info = await m.host.inspect();
    await m.host.apply(plan({ path: m.path, id: info.styles[0].id, sourceXML: info.styles[0].xml, patches: { firstRow: { fill } } }));
    const deck = await openDeck(m.files.get(m.path));
    assert.equal(readPart(deck.styles[0].xml, 'firstRow').fill, fill);
  }
  const before = m.files.get(m.path), info = await m.host.inspect(); m.fail();
  await assert.rejects(() => m.host.apply(plan({ path: m.path, id: info.styles[0].id, sourceXML: info.styles[0].xml })), /partial write/);
  assert.deepEqual(m.files.get(m.path), before);
  assert.equal(m.host.busy, false);
});

test('external modification or target switching during ZIP work stops before writing', async () => {
  for (const type of ['content', 'target', 'opened', 'unsaved']) {
    const m = await mock(type === 'unsaved'), pending = m.host.apply(plan({ path: m.path }));
    if (type === 'content') m.files.set(m.path, new Uint8Array([1,2,3]));
    if (type === 'target') m.host.path = '/docs/other.pptx';
    if (type === 'opened') m.docs.push(m.doc);
    if (type === 'unsaved') m.doc.Saved = false;
    await assert.rejects(() => pending, /修改|切换|打开|未保存/);
    assert.deepEqual(m.writes, []);
  }
});
