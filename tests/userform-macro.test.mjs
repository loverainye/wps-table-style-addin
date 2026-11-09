import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { fixture } from './fixtures.mjs';
import { binaryString } from '../src/host.js';
import { openDeck, readPart } from '../src/styles.js';
import { buildUserFormMacro } from '../scripts/build.mjs';

const root = new URL('..', import.meta.url), dir = await mkdtemp(path.join(tmpdir(), 'wps-userform-bundle-'));
let project;
try {
  const file = path.join(dir, 'runtime.js'); await buildUserFormMacro(file, []);
  const bundle = await readFile(file, 'utf8');
  assert.doesNotMatch(bundle, /SaveCopyAs|UserPicture|GotoSlide|SlideShowSettings|SlideColorSurface/);
  project = JSON.parse(execFileSync('python3', ['-c',
    'import sys,json;sys.path.insert(0,"scripts");import package_userform_macro as p;'+
    'from userform_color_picker import definitions as colors;'+
    'print(json.dumps({"modules":[n.findtext("codetext") for n in p.build_project(open(sys.argv[1]).read()).findall("codemodule")],"items":p.definitions(),"colors":colors()}))', file],
  { cwd: root, encoding: 'utf8', maxBuffer: 4*1024*1024 }));
} finally { await rm(dir, { recursive: true, force: true }); }

function controls(items) {
  return Object.fromEntries(items.map(i => [i.name, {
    Left: i.x, Top: i.y, Width: i.w, Height: i.h, Caption: i.caption, Text: i.text || '',
    Value: false, Enabled: true, Font: { Size: i.font }, options: [],
    BackColor: i.background ? parseInt(i.background.slice(1), 16) : undefined,
    Clear() { this.options = []; this.Text = ''; }, AddItem(s) { this.options.push(s); },
    get ListIndex() { throw new Error('no ListIndex in WPS'); },
    set ListIndex(v) { throw new Error('no ListIndex in WPS'); },
  }]));
}

async function mock() {
  const path = '/docs/原生窗体测试.pptx', original = await fixture(), files = new Map([[path, original]]), writes = [];
  const doc = { FullName: path, Saved: true, ReadOnly: false };
  for (const key of ['Slides', 'SlideShowSettings', 'Save', 'SaveCopyAs', 'Close'])
    Object.defineProperty(doc, key, { get() { throw new Error('forbidden document API '+key); } });
  const app = {
    ActivePresentation: doc, Presentations: { Count: 1, Item: () => doc },
    FileDialog: () => ({ Filters: { Clear() {}, Add() {} }, Show: () => -1, SelectedItems: { Item: () => path } }),
  };
  Object.defineProperty(app, 'Env', { get() { throw new Error('no temporary directories'); } });
  const form = Object.assign(controls(project.items), { Width: 700, Height: 762, Visible: false, shows: 0,
    Show() { this.shows++; this.Visible = true; context.TableStyleForm_Initialize(); }, Hide() { this.Visible = false; } });
  const colorForm = Object.assign(controls(project.colors), { Width: 420, Height: 526, Visible: false,
    Show() { this.Visible = true; }, Hide() { this.Visible = false; } });
  const globals = {
    Application: app, TableStyleForm: form, TableStyleColorForm: colorForm,
    Platform: { OS: () => 'Linux' },
    FileSystem: { readAsBinaryString: p => binaryString(files.get(p)), writeAsBinaryString(p, s) {
      writes.push(p); files.set(p, Uint8Array.from(s, c => c.charCodeAt(0)));
    } },
  };
  Object.defineProperty(globals, 'ffi', { get() { throw new Error('native FileSystem must not load FFI'); } });
  const context = vm.createContext(globals);
  const clear = form.Part.Clear;
  form.Part.Clear = function () { clear.call(this); context.TableStyleForm_Part_Change(); };
  for (const module of project.modules) new vm.Script(module).runInContext(context);
  return { context, form, colorForm, path, files, writes, original, doc };
}

test('released UserForm modules launch lazily and save directly with global WPS files and pure JS UUIDs', async () => {
  const m = await mock(), { context: c, form: f } = m;
  assert.equal(c.TableStylesNativeRuntime.cached, undefined);
  assert.equal(c.TableStylesSlideEditor, undefined); assert.equal(c.TsOpenUserForm_Click, undefined);
  assert.equal(c.OpenTableStyleEditor(), true); assert.equal(f.Part.options.length, 13);
  await c.TableStyleForm_OpenFile_Click(); assert.equal(f.ErrorText.Caption, '');
  f.Preset.Text = '红色'; await c.TableStyleForm_NewStyle_Click();
  const id = c.TableStylesFormEditor.cached.selected.id;
  await c.TableStyleForm_DuplicateStyle_Click();
  assert.notEqual(c.TableStylesFormEditor.cached.selected.id, id);
  for (const edge of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
    assert.equal(f['Width_'+edge].Text, '1');
    await c['TableStyleForm_StepUp_Width_'+edge+'_Click'](); assert.equal(f['Width_'+edge].Text, '1.25');
    for (let i = 0; i < 8; i++) await c['TableStyleForm_StepDown_Width_'+edge+'_Click']();
    assert.equal(f['Width_'+edge].Text, '0.25');
  }
  f.Fill.Text = '#A81234'; await c.TableStyleForm_Fill_Change();
  await c.TableStyleForm_SaveStyle_Click(); assert.equal(f.ErrorText.Caption, '');
  const deck = await openDeck(m.files.get(m.path));
  assert.equal(readPart(deck.styles.find(s => s.id === deck.defaultID).xml, 'firstRow').fill, '#A81234');
  assert.equal(m.writes.length, 2); assert.equal(m.writes[1], m.path);
  assert.equal(f.Visible, true); assert.equal(f.shows, 1);
  assert.match(c.DiagnoseTableStyleEditor(), /WPS FileSystem.*\nUUID：JS Math.random/);
  assert.equal(vm.runInContext('typeof crypto', c), 'undefined');
});

test('all eight colors open the native palette and save without touching a slide, temp file or system API', async () => {
  const m = await mock(), { context: c, form: f, colorForm: p } = m;
  c.OpenTableStyleEditor(); await c.TableStyleForm_OpenFile_Click();
  for (const [i, field] of ['Fill', 'TextColor', 'Color_left', 'Color_right', 'Color_top', 'Color_bottom', 'Color_insideH', 'Color_insideV'].entries()) {
    const before = f[field].Text;
    await c['TableStyleForm_Pick'+field+'_Click']();
    assert.equal(p.Visible, true); assert.equal(f.SaveStyle.Enabled, false);
    await c.TableStyleColorForm_Cancel_Click(); assert.equal(f[field].Text, before);
    await c['TableStyleForm_Pick'+field+'_Click']();
    const hex = '#'+(0x1256A0+i).toString(16).toUpperCase();
    p.Hex.Text = hex; await c.TableStyleColorForm_Hex_Change();
    await c.TableStyleColorForm_Apply_Click();
    assert.equal(f[field].Text, hex); assert.equal(p.Visible, false);
    assert.equal(f['Pick'+field].BackColor, parseInt(hex.slice(1), 16));
  }
  await c.TableStyleForm_PickFill_Click();
  await c.TableStyleColorForm_Hue_MouseDown(1, 0, 192, 10);
  await c.TableStyleColorForm_Hue_MouseUp(0, 0, 192, 10);
  await c.TableStyleColorForm_Plane_MouseDown(1, 0, 384, 0);
  await c.TableStyleColorForm_Plane_MouseUp(0, 0, 384, 0);
  await c.TableStyleColorForm_Apply_Click(); assert.equal(f.Fill.Text, '#00FFFF');
  assert.deepEqual(m.writes, [], 'color picking must never create a file');
  await c.TableStyleForm_SaveStyle_Click(); assert.equal(f.ErrorText.Caption, '');
  const deck = await openDeck(m.files.get(m.path)), part = readPart(deck.styles.find(s => s.id === deck.defaultID).xml, 'firstRow');
  assert.equal(part.fill, '#00FFFF'); assert.equal(part.borders.insideV.color, '#1256A7');
  assert.equal(f.shows, 1); assert.equal(f.Visible, true);
});
