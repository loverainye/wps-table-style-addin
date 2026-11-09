import test from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { readFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { execFileSync } from 'node:child_process';
import { fixture } from './fixtures.mjs';
import { binaryString } from '../src/host.js';
import { openDeck, readPart } from '../src/styles.js';
import { buildMacro } from '../scripts/build.mjs';
import { NativeEditor, colorValue } from '../macro/runtime.js';

const geometry = JSON.parse(execFileSync('python3', ['-c', 'import sys,json;sys.path.insert(0,"macro");from form import definitions;print(json.dumps(definitions()))'], { cwd: new URL('..', import.meta.url), encoding: 'utf8' }));

function formMock() {
  const form = { Width: 700, Height: 762, Hide() {}, Show() {} };
  for (const item of geometry) {
    let text = '', index = -1;
    form[item.name] = {
      Caption: item.caption, Enabled: true, Value: false, Font: {}, options: [],
      Clear() { this.options = []; text = ''; index = -1; },
      AddItem(value) { this.options.push(value); },
      get Text() { return text; }, set Text(v) { text = String(v); index = this.options.indexOf(text); },
      get ListIndex() { throw new Error("WPS ComboBox has no ListIndex"); }, set ListIndex(v) { throw new Error("WPS ComboBox has no ListIndex"); },
    };
  }
  return form;
}

async function mock(options) {
  const file = '/docs/其它汇报.pptx', original = await fixture(options), files = new Map([[file, original]]);
  const app = {
    ActivePresentation: { FullName: '/tools/editor.pptm' },
    Presentations: { Count: 0, Open() { throw new Error('must not open target document'); } },
    FileDialog: () => ({ Filters: { Clear() {}, Add() {} }, Show: () => -1, SelectedItems: { Item: () => file } }),
    FileSystem: {
      readFileString: () => randomUUID(),
      readAsBinaryString: p => binaryString(files.get(p)),
      writeAsBinaryString: (p, data) => files.set(p, Uint8Array.from(data, c => c.charCodeAt(0))),
      Exists: p => files.has(p), Remove: p => files.delete(p),
    },
  };
  const form = formMock(), editor = new NativeEditor(app, form, geometry); editor.initialize();
  return { file, original, files, app, form, editor };
}

test('native form opens a closed PPTX, edits its XML and saves repeatedly without a web or document window', async () => {
  const m = await mock(), { editor: e, form: f } = m;
  assert.equal(f.SaveStyle.Enabled, false);
  await e.dispatch('open'); assert.equal(f.ErrorText.Caption, '');
  assert.equal(f.SaveStyle.Enabled, true);
  f.Fill.Text = '#124C88'; await e.dispatch('fields');
  assert.equal(f.Preview_0_0.BackColor, 0x124C88);
  await e.dispatch('save'); assert.equal(f.ErrorText.Caption, '');
  let output = await openDeck(m.files.get(m.file));
  assert.equal(readPart(output.styles.find(x => x.id === output.defaultID).xml, 'firstRow').fill, '#124C88');
  f.Fill.Text = '#C04020'; await e.dispatch('save'); assert.equal(f.ErrorText.Caption, '');
  output = await openDeck(m.files.get(m.file));
  assert.equal(readPart(output.styles.find(x => x.id === output.defaultID).xml, 'firstRow').fill, '#C04020');
  assert.deepEqual([...m.files].find(([name]) => name.includes('备份-'))[1], m.original);
  assert.equal(m.app.ActivePresentation.FullName, '/tools/editor.pptm');
});

test('native form validates unfinished color input and preserves edits when changing regions', async () => {
  const { editor: e, form: f } = await mock(); await e.dispatch('open');
  const original = f.Fill.Text;
  f.Fill.Text = '#0000FF'; await e.dispatch('fields');
  f.Fill.Text = original; await e.dispatch('fields');
  assert.equal(e.patches.firstRow, undefined, 'reverting input must discard the previous transient value');
  f.Fill.Text = '#00'; f.Part.Text = f.Part.options[0]; await e.dispatch('part');
  assert.match(f.ErrorText.Caption, /颜色/); assert.equal(e.part, 'firstRow');
  assert.equal(f.Part.Text, f.Part.options[10], 'invalid input must keep the original region selected');
  f.Fill.Text = '#001122'; f.Part.Text = f.Part.options[0]; await e.dispatch('part');
  assert.equal(e.patches.firstRow.fill, '#001122');
  f.Part.Text = f.Part.options[10]; await e.dispatch('part');
  assert.equal(f.Fill.Text, '#001122');
});

test('native form creates new styles, handles border controls and default font size', async () => {
  const { editor: e, form: f, files, file } = await mock(); await e.dispatch('open');
  f.Preset.Text = f.Preset.options[0]; await e.dispatch('new');
  assert.equal(e.selected.create, true); assert.equal(f.StyleName.Text, '我的绿色');
  f.Mode_top.Text = f.Mode_top.options[1]; f.Color_top.Text = '#AA3300'; f.Width_top.Text = '2';
  await e.dispatch('allBorders');
  assert.equal(f.Color_insideV.Text, '#AA3300');
  await e.dispatch('save'); assert.equal(f.ErrorText.Caption, '');
  const deck = await openDeck(files.get(file));
  const style = deck.styles.find(x => x.id === deck.defaultID);
  assert.equal(style.name, '我的绿色');
  assert.equal(readPart(style.xml, 'firstRow').borders.insideV.color, '#AA3300');
  assert.equal(f.SetDefaultFont.Enabled, false, 'file without masters cannot set a default size');
});

test('WPS Text-only selectors support new, duplicate, palette and custom hex colors without ListIndex', async () => {
  const { editor: e, form: f, files, file } = await mock();
  await e.dispatch('open');
  assert.equal(f.Scope.Text, '所有表格');
  assert.equal(f.Mode_top.Text, '继承');
  f.Preset.Text = '红色'; await e.dispatch('new');
  assert.equal(f.ErrorText.Caption, '');
  assert.equal(f.StyleName.Text, '我的红色');
  assert.equal(f.Preview_0_0.BackColor, 0xC00000);
  assert.equal(f.Preview_0_0.ForeColor, 0xFFFFFF);
  f.FillChoice.Text = '#385723'; await e.dispatch('fillChoice');
  assert.equal(f.Fill.Text, '#385723');
  assert.equal(f.Preview_0_0.BackColor, 0x385723);
  f.Fill.Text = '#1264A8'; await e.dispatch('fields');
  f.TextColor.Text = '#F0F1F2'; await e.dispatch('fields');
  assert.equal(f.Preview_0_0.BackColor, 0x1264A8);
  assert.equal(f.Preview_0_0.ForeColor, 0xF0F1F2);
  f.Bold.Text = '不加粗'; f.Italic.Text = '斜体';
  await e.dispatch('duplicate');
  assert.equal(f.ErrorText.Caption, '');
  assert.equal(f.StyleName.Text, '我的红色 副本');
  await e.dispatch('save'); assert.equal(f.ErrorText.Caption, '');
  const deck = await openDeck(files.get(file)), style = deck.styles.find(s => s.id === deck.defaultID);
  const header = readPart(style.xml, 'firstRow');
  assert.equal(style.name, '我的红色 副本');
  assert.equal(header.fill, '#1264A8'); assert.equal(header.color, '#F0F1F2');
  assert.equal(header.bold, 'off'); assert.equal(header.italic, 'on');
});

test('native macro bundle runs with no DOM, timers, crypto or ShowDialog supplied by WPS', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'wps-native-macro-'));
  try {
    const bundlePath = path.join(dir, 'runtime.js'); await buildMacro(bundlePath);
    const bundle = await readFile(bundlePath, 'utf8');
    assert.doesNotMatch(bundle, /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/, 'embedded macro source must remain valid XML 1.0 text');
    const source = (await readFile(new URL('../macro/form-code.js', import.meta.url), 'utf8'))
      .replace('__TABLE_STYLES_NATIVE_BUNDLE__', () => bundle)
      .replace('__TABLE_STYLES_NATIVE_GEOMETRY__', () => JSON.stringify(geometry));
    const m = await mock(), context = vm.createContext({ Platform: { OS: () => 'Linux' }, Application: m.app, TableStyleForm: m.form });
    assert.equal(vm.runInContext('typeof DOMParser + ":" + typeof crypto + ":" + typeof setImmediate', context), 'undefined:undefined:undefined');
    new vm.Script(source).runInContext(context);
    context.TableStyleForm_Initialize();
    assert.equal(m.form.ErrorText.Caption, '');
    const e = context.TableStylesFormEditor();
    await e.dispatch('open'); assert.equal(m.form.ErrorText.Caption, '');
    m.form.Fill.Text = '#B00020'; await e.dispatch('save'); assert.equal(m.form.ErrorText.Caption, '');
    const deck = await openDeck(m.files.get(m.file));
    assert.equal(readPart(deck.styles.find(s => s.id === deck.defaultID).xml, 'firstRow').fill, '#B00020');
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('default macro entry initializes before Show and handles synchronous native Change events once', async () => {
  const dir = await mkdtemp(path.join(tmpdir(), 'wps-native-entry-'));
  try {
    const bundlePath = path.join(dir, 'runtime.js'); await buildMacro(bundlePath);
    const bundle = await readFile(bundlePath, 'utf8');
    const launcher = (await readFile(new URL('../macro/launcher.js', import.meta.url), 'utf8'))
      .replace('__TABLE_STYLES_VERSION__', '"test"');
    assert.equal(launcher.match(/function\s+(\w+)\s*\(/)[1], 'OpenTableStyleEditor');
    const source = (await readFile(new URL('../macro/form-code.js', import.meta.url), 'utf8'))
      .replace('__TABLE_STYLES_NATIVE_BUNDLE__', () => bundle)
      .replace('__TABLE_STYLES_NATIVE_GEOMETRY__', () => JSON.stringify(geometry));
    const events = execFileSync('python3', ['-c', 'import sys;sys.path.insert(0,"macro");from form import definitions,event_code;print(event_code(definitions()))'],
      { cwd: new URL('..', import.meta.url), encoding: 'utf8' });
    const m = await mock(), messages = [];
    const context = vm.createContext({ Platform: { OS: () => 'Linux' }, Application: m.app, TableStyleForm: m.form, MsgBox: s => messages.push(s) });
    let clears = 0, shows = 0;
    const clear = m.form.Part.Clear;
    m.form.Part.Clear = function () {
      assert.equal(++clears, 1, 'Change must not recursively reinitialize controls');
      clear.call(this);
      context.TableStyleForm_Part_Change();
    };
    m.form.Show = () => {
      assert.ok(context.TableStylesFormEditor.cached, 'entry initializes even without Initialize event');
      assert.ok(m.form.Part.options.length > 0);
      shows++;
      m.form.Visible = true;
      context.TableStyleForm_Initialize();
    };
    new vm.Script(launcher + '\n' + source + '\n' + events).runInContext(context);
    assert.equal(context.OpenTableStyleEditor(), true);
    assert.equal(context.OpenTableStyleEditor(), true);
    assert.equal(clears, 1);
    assert.equal(shows, 2);
    assert.deepEqual(messages, []);
    assert.equal(context.DiagnoseTableStyleEditor().includes('编辑器已初始化：true'), true);
    assert.equal(messages.length, 1, 'diagnostic entry displays its result');
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('macro startup displays missing resources, initialization errors and invisible forms', async () => {
  const source = (await readFile(new URL('../macro/launcher.js', import.meta.url), 'utf8'))
    .replace('__TABLE_STYLES_VERSION__', '"test"');
  for (const scenario of ['missing', 'empty', 'initialization', 'invisible']) {
    const messages = [], context = vm.createContext({ MsgBox: s => messages.push(s) });
    if (scenario !== 'missing') context.TableStyleForm = {
      Show() {}, Visible: false, Width: scenario === 'empty' ? 0 : 700, Height: 762,
      OpenFile: { Caption: '打开 PPTX' },
    };
    context.TableStylesFormEditor = () => {
      if (scenario === 'initialization') throw new Error('控件初始化失败');
    };
    new vm.Script(source).runInContext(context);
    assert.equal(context.OpenTableStyleEditor(), false);
    assert.equal(messages.length, 1);
    assert.match(messages[0], { missing: /缺少原生窗体/, empty: /未正确载入/, initialization: /控件初始化失败/, invisible: /不可见/ }[scenario]);
  }
});

test('native color input accepts theme colors and rejects invalid values', () => {
  assert.equal(colorValue('c00000'), '#C00000');
  assert.equal(colorValue('theme:accent1'), 'theme:accent1');
  assert.equal(colorValue('无填充', true), 'none');
  assert.throws(() => colorValue('无填充'), /颜色/);
});

test('native form writes optional master default size and keeps the toolbar visible when zoomed out', async () => {
  const { editor: e, form: f, files, file } = await mock({ masters: [null] });
  await e.dispatch('open');
  assert.equal(f.SetDefaultFont.Enabled, true);
  f.SetDefaultFont.Value = true; f.DefaultFontSize.Text = '21.5';
  await e.dispatch('save'); assert.equal(f.ErrorText.Caption, '');
  assert.equal((await openDeck(files.get(file))).fontDefaults.points, 21.5);
  for (let i = 0; i < 5; i++) await e.dispatch('zoomOut');
  assert.equal(f.ZoomLevel.Caption, '50%');
  assert.ok(f.Width >= 700);
});
