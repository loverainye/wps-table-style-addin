import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { fixture, oldID, styleXML, plan } from './fixtures.mjs';
import { openDeck, applyStyle, readPart, parseXML, editStyle, guid, A, P, xml, children, child, PARTS } from '../src/styles.js';

test('lists all 13 reference regions and independently saves their formatting', async () => {
  assert.deepEqual(PARTS, [
    ['wholeTbl', '整个表格'], ['band1H', '第一行条纹'], ['band2H', '第二行条纹'],
    ['band1V', '第一列条纹'], ['band2V', '第二列条纹'], ['lastCol', '最后一列'],
    ['firstCol', '第一列'], ['lastRow', '汇总行'], ['seCell', '最后一个汇总单元格'],
    ['swCell', '第一个汇总单元格'], ['firstRow', '标题行'], ['neCell', '最后一个标题单元格'],
    ['nwCell', '第一个标题单元格'],
  ]);
  const patches = Object.fromEntries(PARTS.map(([key], i) => [key, {
    fill: '#' + (0x123400 + i).toString(16).toUpperCase(), color: '#FFFFFF', bold: i % 2 ? 'off' : 'on',
    borders: { top: { mode: 'solid', color: '#654321', width: 1.25 } },
  }]));
  const output = await openDeck((await applyStyle(await openDeck(await fixture()), plan({ patches }))).bytes);
  const style = output.styles.find(s => s.id === output.defaultID);
  for (const [key] of PARTS) {
    const actual = readPart(style.xml, key);
    for (const property of ['fill', 'color', 'bold']) assert.equal(actual[property], patches[key][property], key + '.' + property);
    assert.equal(actual.borders.top.width, 1.25, key);
    assert.equal(actual.borders.top.color, '#654321', key);
  }
  assert.deepEqual(children(parseXML(style.xml).documentElement).map(n => n.localName), PARTS.map(([key]) => key));
});

test('reads table references with actual presentation slide IDs', async () => {
  const deck = await openDeck(await fixture());
  assert.equal(deck.defaultID, oldID);
  assert.deepEqual(deck.tables.map(t => [t.slideID, t.shapeID]), [['256', '5'], ['300', '5']]);
});
test('changes native definition and references with a fresh version-4 GUID', async () => {
  const deck = await openDeck(await fixture());
  const result = await applyStyle(deck, plan());
  assert.notEqual(result.id, oldID);
  assert.match(result.id, /^\{[A-F0-9]{8}-[A-F0-9]{4}-4[A-F0-9]{3}-[89AB][A-F0-9]{3}-[A-F0-9]{12}\}$/);
  const output = await openDeck(result.bytes);
  assert.equal(output.defaultID, result.id);
  assert.equal(readPart(output.styles.find(s => s.id === result.id).xml, 'firstRow').fill, '#0030C2');
  assert.ok(output.tables.every(t => t.styleID === result.id));
  assert.equal(output.styles.some(s => s.id === oldID), false);
  assert.ok(deck.tables.every(t => t.styleID === oldID));
  assert.match(await deck.zip.file('ppt/tableStyles.xml').async('string'), /C00000/);
  assert.doesNotMatch(await deck.zip.file('ppt/tableStyles.xml').async('string'), /0030C2/);
  assert.equal(await output.zip.file('customXml/item1.xml').async('string'), '<keep>untouched</keep>');
  assert.deepEqual(await output.zip.file('ppt/media/image1.bin').async('uint8array'), new Uint8Array([0, 255, 4, 128, 0]));
});
test('applies only to selected table and retains old definition for other tables', async () => {
  const output = await openDeck((await applyStyle(await openDeck(await fixture()), plan({ scope: 'selected', targets: [{ slideID: '300', shapeID: '5' }] }))).bytes);
  assert.equal(output.tables[0].styleID, oldID);
  assert.notEqual(output.tables[1].styleID, oldID);
  assert.ok(output.styles.some(s => s.id === oldID));
});
test('default-only apply leaves slide XML byte-identical and existing references intact', async () => {
  const deck = await openDeck(await fixture());
  const result = await applyStyle(deck, plan({ scope: 'none' }));
  const output = await openDeck(result.bytes);
  assert.equal(output.defaultID, result.id);
  assert.ok(output.tables.every(t => t.styleID === oldID));
  assert.equal(await output.zip.file('ppt/slides/slide1.xml').async('string'), await deck.zip.file('ppt/slides/slide1.xml').async('string'));
});
test('existing style with no definition edits keeps GUID', async () => {
  const deck = await openDeck(await fixture());
  const result = await applyStyle(deck, plan({ name: deck.styles[0].name, patches: {} }));
  assert.equal(result.id, oldID); assert.equal(result.modified, false);
});
test('clears only local colors, borders and bold/italic, keeping text, size and margins', async () => {
  const out = await openDeck((await applyStyle(await openDeck(await fixture({ overrides: true })), plan({ clearOverrides: true }))).bytes);
  const s = await out.zip.file('ppt/slides/slide1.xml').async('string');
  assert.match(s, /UNCHANGED TEXT 1/); assert.match(s, /sz="1800"/); assert.match(s, /marL="1234"/);
  assert.doesNotMatch(s, /solidFill|a:lnL|b="1"/);
});
test('preserves unsupported style effects and theme transforms when another property changes', () => {
  const original = styleXML.replace('</a:tblStyle>', '<a:extLst><a:ext uri="test"><x:keep xmlns:x="urn:test"/></a:ext></a:extLst></a:tblStyle>');
  const output = editStyle(original, 'A & B <样式>', guid(), { firstRow: { bold: 'off' } });
  assert.match(output, /x:keep/); assert.match(output, /A &amp; B &lt;样式&gt;/);
  assert.equal(readPart(output, 'firstRow').fill, '#C00000');
  assert.equal(readPart(output, 'firstRow').bold, 'off');
});
test('writes valid table-style child order and explicit border removal', () => {
  const s = editStyle(styleXML, 'border', guid(), { firstRow: { fill: '#112233', bold: 'on', borders: { top: { mode: 'none' } } } });
  const root = parseXML(s).documentElement, row = child(root, 'firstRow');
  assert.deepEqual(children(row).map(n => n.localName), ['tcTxStyle', 'tcStyle']);
  assert.deepEqual(children(child(row, 'tcStyle')).map(n => n.localName), ['tcBdr', 'fill']);
  assert.match(s, /<a:noFill\/>/);
});
test('adds tableStyles relationship and content type when missing', async () => {
  const d = await openDeck(await fixture({ styles: false }));
  const out = await openDeck((await applyStyle(d, plan({ create: true, baseXML: styleXML, scope: 'all' }))).bytes);
  assert.match(await out.zip.file('ppt/_rels/presentation.xml.rels').async('string'), /relationships\/tableStyles/);
  assert.match(await out.zip.file('[Content_Types].xml').async('string'), /presentationml.tableStyles/);
});
test('rejects invalid scope, selection, color, DTD, unsigned non-PPTX and invalid border width', async () => {
  const d = await openDeck(await fixture());
  await assert.rejects(() => applyStyle(d, plan({ scope: 'selected', targets: [] })), /选中/);
  await assert.rejects(() => applyStyle(d, plan({ scope: 'oops' })), /范围/);
  await assert.rejects(() => applyStyle(d, plan({ patches: { firstRow: { fill: 'red' } } })), /颜色/);
  assert.throws(() => parseXML('<!DOCTYPE x><x/>'), /DTD/);
  await assert.rejects(() => openDeck(new Uint8Array([1, 2, 3])));
  for (const width of [-1, 0.1, 0.24]) {
    assert.throws(() => editStyle(styleXML, 'x', guid(), { wholeTbl: { borders: { top: { mode: 'solid', width } } } }), /线宽/);
  }
  for (const width of [0.25, 1.25, 20]) {
    const output = editStyle(styleXML, 'x', guid(), { wholeTbl: { borders: { top: { mode: 'solid', color: '#123456', width } } } });
    assert.equal(readPart(output, 'wholeTbl').borders.top.width, width);
  }
});
test('roundtrips a real Pandoc deck without changing unrelated package entries', async () => {
  const data = await readFile(new URL('fixtures/pandoc-basic.pptx', import.meta.url));
  const initial = await openDeck(data);
  const template = await applyStyle(initial, { create: true, baseXML: styleXML, name: '参考样式', patches: {}, scope: 'all', setDefault: true });
  const deck = await openDeck(template.bytes), s = deck.styles.find(s => !s.builtin);
  const result = await applyStyle(deck, { id: s.id, name: s.name, patches: { firstRow: { fill: '#7F1542' } }, scope: 'linked', setDefault: true });
  const output = await openDeck(result.bytes);
  assert.ok(output.tables.length > 0);
  for (const [name, f] of Object.entries(deck.zip.files)) {
    if (f.dir || name === 'ppt/tableStyles.xml' || deck.tables.some(t => t.path === name) || name === '[Content_Types].xml' || name === 'ppt/_rels/presentation.xml.rels') continue;
    assert.deepEqual(await output.zip.file(name).async('uint8array'), await f.async('uint8array'), name);
  }
});

test('reads uniform, unset and mixed defaults across master text levels', async () => {
  const style = '<a:defPPr><a:defRPr sz="1850"/></a:defPPr>';
  assert.deepEqual((await openDeck(await fixture({ masters: [style, style] }))).fontDefaults, { masterCount: 2, points: 18.5, mixed: false });
  assert.deepEqual((await openDeck(await fixture({ masters: ['', null] }))).fontDefaults, { masterCount: 2, points: null, mixed: false });
  assert.equal((await openDeck(await fixture({ masters: [style + '<a:lvl2pPr><a:defRPr sz="2200"/></a:lvl2pPr>'] }))).fontDefaults.mixed, true);
});

test('sets all master fallback levels without touching table sizes, style GUID or title/body formatting', async () => {
  const deck = await openDeck(await fixture({ masters: ['<a:lvl2pPr marL="123"><a:defRPr sz="1200" b="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:defRPr></a:lvl2pPr>', ''] }));
  const originalMaster = await deck.zip.file('ppt/slideMasters/slideMaster1.xml').async('string');
  const result = await applyStyle(deck, plan({ name: deck.styles[0].name, patches: {}, scope: 'none', setDefault: false, defaultFontSize: 24.5 }));
  const out = await openDeck(result.bytes);
  assert.deepEqual(out.fontDefaults, { masterCount: 2, points: 24.5, mixed: false });
  assert.equal(result.changedMasters, 2);
  assert.equal(result.changedTables, 0);
  assert.equal(result.id, oldID);
  assert.equal(result.modified, false);
  assert.equal(out.stylesXML, deck.stylesXML);
  assert.equal(await deck.zip.file('ppt/slideMasters/slideMaster1.xml').async('string'), originalMaster);
  for (const table of deck.tables) {
    assert.equal(await out.zip.file(table.path).async('string'), await deck.zip.file(table.path).async('string'));
    assert.match(await out.zip.file(table.path).async('string'), /sz="1800"/);
  }
  const before = parseXML(originalMaster).documentElement;
  const after = parseXML(await out.zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).documentElement;
  for (const tag of ['titleStyle', 'bodyStyle']) assert.equal(xml(child(child(after, 'txStyles', P), tag, P)), xml(child(child(before, 'txStyles', P), tag, P)));
  const other = child(child(after, 'txStyles', P), 'otherStyle', P);
  assert.deepEqual(children(other).map(n => n.localName), ['defPPr', ...Array.from({ length: 9 }, (_, i) => `lvl${i + 1}pPr`)]);
  assert.equal(child(other, 'lvl2pPr').getAttribute('marL'), '123');
  const run = child(child(other, 'lvl2pPr'), 'defRPr');
  assert.equal(run.getAttribute('b'), '1');
  assert.equal(child(child(run, 'solidFill'), 'srgbClr').getAttribute('val'), '123456');
});

test('creates missing master text styles before extensions, retaining other resources', async () => {
  const deck = await openDeck(await fixture({ masters: [null] }));
  const out = await openDeck((await applyStyle(deck, plan({ defaultFontSize: 20 }))).bytes);
  assert.equal(out.fontDefaults.points, 20);
  const master = parseXML(await out.zip.file('ppt/slideMasters/slideMaster1.xml').async('string')).documentElement;
  assert.deepEqual(children(master, undefined, P).map(n => n.localName), ['cSld', 'txStyles', 'extLst']);
  assert.equal(child(child(master, 'extLst', P), 'ext', P).getAttribute('uri'), 'keep');
});

test('leaves master XML byte-identical when font setting is not requested', async () => {
  const deck = await openDeck(await fixture({ masters: [null, '<a:defPPr><a:defRPr sz="1800"/></a:defPPr>'] }));
  const result = await applyStyle(deck, plan());
  assert.equal(result.changedMasters, 0);
  const out = await openDeck(result.bytes);
  for (const path of Object.keys(deck.zip.files).filter(n => n.startsWith('ppt/slideMasters/') && !deck.zip.files[n].dir)) assert.equal(await out.zip.file(path).async('string'), await deck.zip.file(path).async('string'));
});

test('rejects invalid font sizes and a deck without masters without mutating source', async () => {
  const deck = await openDeck(await fixture({ masters: [''] }));
  const original = await deck.zip.file('ppt/slideMasters/slideMaster1.xml').async('string');
  for (const size of [0, -1, 4001, NaN, Infinity, '18', null, 18.123]) await assert.rejects(() => applyStyle(deck, plan({ defaultFontSize: size })), /默认字号/);
  assert.equal(await deck.zip.file('ppt/slideMasters/slideMaster1.xml').async('string'), original);
  const noMasters = await openDeck(await fixture());
  await assert.rejects(() => applyStyle(noMasters, plan({ defaultFontSize: 18 })), /母版/);
});

test('updates font defaults in a real Pandoc deck while preserving all slide and layout entries', async () => {
  const deck = await openDeck(await readFile(new URL('fixtures/pandoc-basic.pptx', import.meta.url)));
  const s = deck.styles[0];
  const result = await applyStyle(deck, { id: s.id, name: s.name, scope: 'none', patches: {}, defaultFontSize: 22 });
  const out = await openDeck(result.bytes);
  assert.ok(result.changedMasters > 0);
  assert.equal(out.fontDefaults.points, 22);
  assert.equal(result.id, s.id);
  for (const [path, f] of Object.entries(deck.zip.files)) {
    if (f.dir || path.startsWith('ppt/slideMasters/') || path === '[Content_Types].xml' || path === 'ppt/_rels/presentation.xml.rels') continue;
    assert.deepEqual(await out.zip.file(path).async('uint8array'), await f.async('uint8array'), path);
  }
});
