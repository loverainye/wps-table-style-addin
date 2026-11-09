import test from 'node:test';
import assert from 'node:assert/strict';
import { randomUUID } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { fixture } from './fixtures.mjs';
import { binaryString } from '../src/host.js';
import { openDeck, readPart } from '../src/styles.js';
import { buildSlideMacro } from '../scripts/build.mjs';
import { SlideEditor } from '../macro/slide-runtime.js';
import { NativeColorPicker, hsvToHex } from '../macro/color-picker.js';
import { SlideColorSurface, gradientBitmap, PLANE, HUE } from '../macro/slide-color-surface.js';

const root = new URL('..', import.meta.url);
const python = code => execFileSync('python3', ['-c', 'import sys;sys.path.insert(0,"macro");'+code], { cwd: root, encoding: 'utf8' });
const items = JSON.parse(python('import json;from slide import definitions;print(json.dumps(definitions()))'));
const eventCode = python('from slide import definitions,event_code;print(event_code(definitions()))');
const pickerItems = JSON.parse(python('import json;from color_picker import definitions;print(json.dumps(definitions()))'));
const pickerBItems = JSON.parse(python('import json;from color_picker import definitions;print(json.dumps(definitions(3)))'));
const pickerEvents = python('from color_picker import definitions,event_code;print(event_code(definitions(2)+definitions(3)))');
const userformItems = JSON.parse(python('import json;from userform import definitions;print(json.dumps(definitions()))'));
const userformColorItems = JSON.parse(python('import json;from userform_color_picker import definitions;print(json.dumps(definitions()))'));

async function combinedProject(m, withSlideControls = false) {
  const dir = await mkdtemp(path.join(tmpdir(), 'wps-combined-macro-'));
  try {
    const bundlePath = path.join(dir, 'runtime.js');
    await buildSlideMacro(bundlePath, []);
    const modules = JSON.parse(execFileSync('python3', ['-c',
      'import sys,json;sys.path.insert(0,"scripts");import package_slide_macro as p;'+
      's=(p.ROOT/"macro/slide-code.js").read_text().replace("__TABLE_STYLES_VERSION__",json.dumps("test"));'+
      'project=p.build_project(s,p.runtime_code(open(sys.argv[1]).read()),p.event_code(p.definitions())+p.color_events(p.color_definitions(2)+p.color_definitions(3)));'+
      'print(json.dumps([n.findtext("codetext") for n in project.findall("codemodule")]))', bundlePath],
    { cwd: root, encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 }));
    const form = Object.assign(controls(userformItems), { Width: 700, Height: 762, Visible: false, shows: 0,
      Show() { this.shows++; this.Visible = true; context.TableStyleForm_Initialize(); },
      Hide() { this.Visible = false; },
    });
    const colorForm = Object.assign(controls(userformColorItems), { Width: 420, Height: 526, Visible: false, shows: 0,
      Show() { this.shows++; this.Visible = true; }, Hide() { this.Visible = false; },
    });
    const messages = [], globals = withSlideControls ? Object.fromEntries([
      ...items.map(i => [i.native, m.form[i.name]]), ...pickerItems.map(i => [i.native, m.pickerControls[i.name]]),
      ...pickerBItems.map(i => [i.native, m.pickerBControls[i.name]]),
    ]) : {};
    const context = vm.createContext({ Application: m.app, TableStyleForm: form, TableStyleColorForm: colorForm, ...globals, MsgBox: s => messages.push(s) });
    const clear = form.Part.Clear;
    form.Part.Clear = function () { clear.call(this); context.TableStyleForm_Part_Change(); };
    for (const module of modules) new vm.Script(module).runInContext(context);
    return { context, form, colorForm, messages };
  } finally { await rm(dir, { recursive: true, force: true }); }
}

test('combined PPTM opens UserForm from its fourth-page button without any slide controls or eager initialization', async () => {
  const m = await mock(), { context: c, form: f, messages } = await combinedProject(m);
  assert.equal(c.TableStylesNativeRuntime.cached, undefined);
  assert.equal(c.TableStylesFormEditor.cached, undefined);
  assert.equal(c.TsOpenUserForm_Click(), true);
  assert.equal(f.Visible, true); assert.equal(f.shows, 1);
  assert.equal(f.Part.options.length, 13);
  assert.equal(f.Width_left.Text, '1'); assert.equal(f.StepUp_Width_left.Enabled, false);
  assert.equal(c.TableStylesSlideEditor.cached, undefined);
  const editor = c.TableStylesFormEditor.cached, runtime = c.TableStylesNativeRuntime();
  assert.equal(c.TableStylesNativeRuntime(), runtime);
  await c.TableStyleForm_OpenFile_Click(); assert.equal(f.ErrorText.Caption, '');
  f.Preset.Text = '红色'; await c.TableStyleForm_NewStyle_Click();
  f.Fill.Text = '#123456'; await c.TableStyleForm_Fill_Change();
  await c.TableStyleForm_CloseEditor_Click(); assert.equal(f.Visible, false);
  assert.equal(c.TsOpenUserForm_Click(), true); assert.equal(f.Fill.Text, '#123456');
  assert.equal(c.TableStylesFormEditor.cached, editor); assert.equal(f.shows, 2);
  assert.deepEqual(messages, []);
});

test('UserForm events save colors, copies, regions, font size and quarter-point borders with separate interface drafts', async () => {
  const m = await mock(), { context: c, form: f } = await combinedProject(m, true);
  assert.equal(c.OpenTableStyleEditor(), true);
  const slide = c.TableStylesSlideEditor.cached;
  await slide.dispatch('open');
  m.form.Fill.Text = '#789ABC'; await slide.dispatch('fields');
  assert.equal(c.TsOpenUserForm_Click(), true);
  const editor = c.TableStylesFormEditor.cached;
  assert.notEqual(editor, slide); assert.equal(editor instanceof c.TableStylesNativeRuntime().UserFormEditor, true);
  await c.TableStyleForm_OpenFile_Click(); assert.equal(f.ErrorText.Caption, '');
  f.Preset.Text = '蓝色'; await c.TableStyleForm_NewStyle_Click();
  assert.equal(f.StyleName.Text, '我的蓝色');
  const originalID = editor.selected.id;
  await c.TableStyleForm_DuplicateStyle_Click(); assert.notEqual(editor.selected.id, originalID);
  assert.equal(f.StyleName.Text, '我的蓝色 副本');
  f.Fill.Text = '#C00000'; await c.TableStyleForm_Fill_Change();
  assert.equal(f.Fill.Text, '#C00000'); assert.equal(f.Preview_0_0.BackColor, 0xC00000);
  f.TextColor.Text = '#FFFFFF'; await c.TableStyleForm_TextColor_Change();
  for (const edge of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
    assert.equal(f['Width_'+edge].Text, '1');
    await c['TableStyleForm_StepUp_Width_'+edge+'_Click'](); assert.equal(f['Width_'+edge].Text, '1.25');
    await c['TableStyleForm_Width_'+edge+'_KeyUp'](40, 0);
    await c['TableStyleForm_StepDown_Width_'+edge+'_Click'](); assert.equal(f['Width_'+edge].Text, '0.75');
    for (let i = 0; i < 5; i++) await c['TableStyleForm_StepDown_Width_'+edge+'_Click']();
    assert.equal(f['Width_'+edge].Text, '0.25');
    await c['TableStyleForm_Width_'+edge+'_KeyUp'](38, 0); assert.equal(f['Width_'+edge].Text, '0.5');
    assert.equal(f['Mode_'+edge].Text, '线条');
    f['Color_'+edge].Text = '#112233'; await c['TableStyleForm_Color_'+edge+'_Change']();
  }
  f.Part.Text = f.Part.options[0]; await c.TableStyleForm_Part_Change();
  f.Fill.Text = '#CCDDEE'; await c.TableStyleForm_Fill_Change();
  f.Part.Text = f.Part.options[10]; await c.TableStyleForm_Part_Change();
  assert.equal(f.Fill.Text, '#C00000');
  f.SetDefaultFont.Value = true; await c.TableStyleForm_SetDefaultFont_Click();
  const beforeFont = Number(f.DefaultFontSize.Text);
  await c.TableStyleForm_StepUp_DefaultFontSize_Click(); assert.equal(Number(f.DefaultFontSize.Text), beforeFont+1);
  await c.TableStyleForm_SaveStyle_Click(); assert.equal(f.ErrorText.Caption, '');
  const deck = await openDeck(m.files.get(m.file)), style = deck.styles.find(s => s.id === deck.defaultID);
  const part = readPart(style.xml, 'firstRow');
  assert.equal(part.fill, '#C00000'); assert.equal(part.color, '#FFFFFF');
  assert.equal(readPart(style.xml, 'wholeTbl').fill, '#CCDDEE');
  for (const edge of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
    assert.equal(part.borders[edge].width, 0.5); assert.equal(part.borders[edge].color, '#112233');
  }
  assert.equal(deck.fontDefaults.points, beforeFont+1);
  assert.equal(m.form.Fill.Text, '#789ABC');
  assert.equal([...m.files.keys()].filter(p => p.includes('表格样式备份')).length, 1);
  // An explicit inherit choice must survive later captures and saving.
  f.Mode_left.Text = '继承'; await c.TableStyleForm_Mode_left_Change();
  await c.TableStyleForm_SaveStyle_Click(); assert.equal(f.ErrorText.Caption, '');
  const again = await openDeck(m.files.get(m.file));
  assert.equal(readPart(again.styles.find(s => s.id === again.defaultID).xml, 'firstRow').borders.left.mode, 'inherit');
  // Edit the original style that the slide interface still has open.
  f.Styles.Text = f.Styles.options[editor.allStyles().findIndex(s => s.id === slide.selected.id)];
  await c.TableStyleForm_Styles_Change();
  f.Fill.Text = '#009966'; await c.TableStyleForm_Fill_Change();
  await c.TableStyleForm_SaveStyle_Click(); assert.equal(f.ErrorText.Caption, '');
  // The other interface sees a stale definition and must not overwrite it.
  await slide.dispatch('save'); assert.match(m.form.ErrorText.Caption, /样式已更改/);
});

test('all eight UserForm colors use an independent native dialog without reading or navigating any PPT page', async () => {
  const m = await mock(), { context: c, form: f, colorForm: p } = await combinedProject(m);
  m.show.page = 4;
  Object.defineProperty(m.app.ActivePresentation, 'Slides', { get() { throw new Error('do not access any PPT slide'); } });
  m.showWindow.View.GotoSlide = () => { throw new Error('do not navigate PPT pages'); };
  m.app.ActivePresentation.SlideShowSettings.Run = () => { throw new Error('do not start playback'); };
  c.TsOpenUserForm_Click(); await c.TableStyleForm_OpenFile_Click();
  const e = c.TableStylesFormEditor.cached, id = e.selected.id;
  assert.equal(f.FillChoice, undefined); assert.equal(f.TextColorChoice, undefined);
  const invoke = (suffix, ...args) => c['TableStyleColorForm_'+suffix](...args);
  for (const [i, field] of ['Fill', 'TextColor', 'Color_left', 'Color_right', 'Color_top', 'Color_bottom', 'Color_insideH', 'Color_insideV'].entries()) {
    const before = f[field].Text;
    await c['TableStyleForm_Pick'+field+'_Click']();
    assert.equal(f.Visible, true); assert.equal(p.Visible, true); assert.equal(m.show.page, 4);
    assert.equal(f.SaveStyle.Enabled, false); assert.equal(f.PickFill.Enabled, false);
    assert.equal(c.TableStylesSlideEditor.cached, undefined, 'the UserForm picker must not initialize the slide editor');
    await invoke('Cancel_Click');
    assert.equal(p.Visible, false); assert.equal(f.Visible, true); assert.equal(m.show.page, 4); assert.equal(f[field].Text, before);
    assert.equal(f.SaveStyle.Enabled, true);
    await c['TableStyleForm_Pick'+field+'_Click']();
    const hex = '#'+(0x1264A8+i).toString(16).toUpperCase();
    p.Hex.Text = hex; await invoke('Hex_Change');
    await invoke('Hex_KeyUp', 13, 0);
    assert.equal(p.Current.BackColor, parseInt(hex.slice(1), 16));
    await invoke('Apply_Click');
    assert.equal(f[field].Text, hex); assert.ok(f[field].options.includes(hex));
    assert.equal(f['Pick'+field].BackColor, parseInt(hex.slice(1), 16));
    assert.equal(f.Visible, true); assert.equal(p.Visible, false); assert.equal(m.show.page, 4);
    assert.equal(e.selected.id, id); assert.equal(f.StyleName.Text, e.selected.name);
    if (field.startsWith('Color_')) assert.equal(f['Mode_'+field.slice(6)].Text, '线条');
  }
  await c.TableStyleForm_SaveStyle_Click(); assert.equal(f.ErrorText.Caption, '');
  const deck = await openDeck(m.files.get(m.file)), part = readPart(deck.styles.find(s => s.id === deck.defaultID).xml, 'firstRow');
  assert.equal(part.fill, '#1264A8'); assert.equal(part.color, '#1264A9');
  assert.equal(part.borders.insideV.color, '#1264AF');
  // Pointer events must also reach this picker without initializing the slide editor.
  await c.TableStyleForm_PickFill_Click();
  await invoke('Hue_MouseDown', 1, 0, 192, 10); await invoke('Hue_MouseUp', 0, 0, 192, 10);
  await invoke('Plane_MouseDown', 1, 0, 384, 0); await invoke('Plane_MouseUp', 0, 0, 384, 0);
  await invoke('Apply_Click'); assert.equal(f.Fill.Text, '#00FFFF'); assert.equal(f.Visible, true);
  assert.equal(f.shows, 1, 'the main form must never close or reopen during color picking');
  assert.equal(m.show.frames.length, 0);
});

test('UserForm picker restores main controls after open failure, supports close retry, and handles the modal window X', async () => {
  const m = await mock(), { context: c, form: f, colorForm: p } = await combinedProject(m);
  m.show.page = 4; c.TsOpenUserForm_Click(); await c.TableStyleForm_OpenFile_Click();
  const show = p.Show, hide = p.Hide;
  p.Show = () => { throw new Error('dialog refused'); };
  await c.TableStyleForm_PickFill_Click();
  assert.equal(f.Visible, true); assert.equal(m.show.page, 4);
  assert.match(f.ErrorText.Caption, /dialog refused/); assert.equal(f.SaveStyle.Enabled, true);
  assert.equal(c.TableStylesFormEditor.cached.picker.form.visible, false);
  p.Show = show; await c.TableStyleForm_PickFill_Click();
  assert.equal(p.Visible, true); assert.equal(f.SaveStyle.Enabled, false);
  p.Hide = () => { throw new Error('close refused'); };
  await c.TableStyleColorForm_Cancel_Click();
  assert.equal(p.Visible, true); assert.equal(c.TableStylesFormEditor.cached.picker.form.visible, true);
  assert.match(p.ErrorText.Caption, /close refused/);
  p.Hide = hide; await c.TableStyleColorForm_Cancel_Click();
  assert.equal(f.Visible, true); assert.equal(m.show.page, 4);
  assert.equal(p.Visible, false); assert.equal(f.SaveStyle.Enabled, true);
  // Native modal Show returns after its window X closes, without a Cancel event.
  p.Show = function () { this.Visible = true; this.Visible = false; };
  await c.TableStyleForm_PickFill_Click();
  assert.equal(f.SaveStyle.Enabled, true); assert.equal(f.PickFill.Enabled, true);
  assert.equal(c.TableStylesFormEditor.cached.picker.form.visible, false);
  assert.deepEqual(m.files.get(m.file), m.original);
});

test('native picker computes precise pointer colors, limits gradient repaints, and validates packaged HEX/RGB events', async () => {
  const m = await mock(), { context: c, form: f, colorForm: p } = await combinedProject(m);
  c.TsOpenUserForm_Click(); await c.TableStyleForm_OpenFile_Click();
  let paints = 0;
  for (const item of userformColorItems.filter(i => i.paint_cell === 'plane')) {
    let color = p[item.name].BackColor;
    Object.defineProperty(p[item.name], 'BackColor', { get: () => color, set(value) { paints++; color = value; } });
  }
  await c.TableStyleForm_PickFill_Click();
  assert.equal(paints, 0, 'red gradient is already embedded in the native resource');
  await c.TableStyleColorForm_Hue_MouseDown(1, 0, 128, 10);
  await c.TableStyleColorForm_Hue_MouseMove(0, 0, 144, 10);
  assert.equal(paints, 0, 'dragging updates preview without repainting hundreds of labels');
  await c.TableStyleColorForm_Hue_MouseUp(0, 0, 144, 10);
  assert.equal(paints, 32*24);
  assert.equal(p.PlaneCell_0_31.BackColor, parseInt(hsvToHex(225, 1, 1).slice(1), 16));
  assert.equal(p.PlaneCell_23_31.BackColor, 0);
  p.Plane.Width = 768; p.Plane.Height = 432;
  await c.TableStyleColorForm_Plane_MouseDown(1, 0, 321.25, 147.75);
  await c.TableStyleColorForm_Plane_MouseUp(0, 0, 321.25, 147.75);
  assert.equal(p.Hex.Text, hsvToHex(225, 321.25/768, 1-147.75/432), 'pointer precision must not be rounded to color cells');
  assert.equal(paints, 32*24, 'saturation/value selection keeps the same background');
  const before = f.Fill.Text;
  p.Hex.Text = '#12'; await c.TableStyleColorForm_Hex_Change(); await c.TableStyleColorForm_Apply_Click();
  assert.match(p.ErrorText.Caption, /六位/); assert.equal(p.Visible, true); assert.equal(f.Fill.Text, before);
  p.R.Text = '256'; p.G.Text = '0'; p.B.Text = '0';
  await c.TableStyleColorForm_R_Change(); await c.TableStyleColorForm_Apply_Click();
  assert.match(p.ErrorText.Caption, /0–255/); assert.equal(p.Visible, true);
  p.R.Text = '18'; p.G.Text = '52'; p.B.Text = '86';
  await c.TableStyleColorForm_B_Change(); await c.TableStyleColorForm_B_KeyUp(13, 0);
  await c.TableStyleColorForm_Apply_Click();
  assert.equal(f.Fill.Text, '#123456'); assert.equal(p.Visible, false); assert.equal(f.SaveStyle.Enabled, true);
  await c.TableStyleForm_PickTextColor_Click();
  await c.TableStyleColorForm_Swatch4_Click(); assert.equal(p.Hex.Text, '#C00000');
  await c.TableStyleColorForm_Apply_Click(); assert.equal(f.TextColor.Text, '#C00000');
});

test('native UserForm and slide color pickers keep their own controls, gradients and callbacks', async () => {
  const m = await mock(), { context: c, form: f, colorForm: p } = await combinedProject(m, true);
  m.show.page = 4; c.TsOpenUserForm_Click(); await c.TableStyleForm_OpenFile_Click();
  const invoke = suffix => c[(m.show.page === 2 ? 'TsPicker' : 'TsPickerB')+suffix]();
  await c.TableStyleForm_PickFill_Click();
  p.Hex.Text = '#112233'; await c.TableStyleColorForm_Hex_Change(); await c.TableStyleColorForm_Hex_KeyUp(13, 0);
  const nativeColor = p.PlaneCell_0_31.BackColor;
  await c.TableStyleColorForm_Apply_Click();
  assert.equal(f.Fill.Text, '#112233'); assert.equal(m.show.page, 4);
  m.showWindow.View.GotoSlide(1, -1); c.OpenTableStyleEditor();
  await c.TsOpenFile_Click(); await c.TsPickFill_Click();
  m.getPickerControls(m.show.page).Hex.Text = '#EF5511'; await invoke('Hex_Change'); await invoke('Apply_Click');
  assert.equal(m.form.Fill.Text, '#EF5511'); assert.equal(m.show.page, 1);
  const before = m.show.frames.length;
  m.showWindow.View.GotoSlide(4, -1); await c.TableStyleForm_PickTextColor_Click();
  assert.equal(m.show.frames.length, before+1, 'only the explicit move to the launcher changes the PPT page');
  assert.equal(p.PlaneCell_0_31.BackColor, nativeColor);
  await c.TableStyleColorForm_Cancel_Click(); assert.equal(m.show.page, 4);
  assert.equal(f.Fill.Text, '#112233'); assert.equal(m.form.Fill.Text, '#EF5511');
});

function controls(definitions = items) {
  return Object.fromEntries(definitions.map(item => [item.name, {
    Name: item.native, Left: item.x, Top: item.y, Width: item.w, Height: item.h,
    Caption: item.caption, Text: item.text || '', Value: false, Visible: !item.hidden, Enabled: true, Font: { Size: item.font }, options: [], addCalls: 0,
    BackColor: item.background ? parseInt(item.background.slice(1), 16) : undefined,
    Move(x, y, w, h) { this.Left = x; this.Top = y; if (w !== undefined) this.Width = w; if (h !== undefined) this.Height = h; },
    Clear() { this.options = []; this.Text = ''; },
    AddItem(label) { this.addCalls++; this.options.push(label); },
    get ListIndex() { throw new Error('WPS does not expose ListIndex'); },
    set ListIndex(value) { throw new Error('WPS does not expose ListIndex'); },
  }]));
}

async function mock({ playing = true } = {}) {
  const file = '/docs/幻灯片版测试.pptx', original = await fixture({ masters: [null] });
  const files = new Map([[file, original]]), form = controls();
  let zoom = 85, fit = 0;
  const canvasPane = { ViewType: 1, Activate() { window.ActivePane = canvasPane; } };
  const window = { ViewType: 9, ActivePane: { ViewType: 11 }, Activate() { app.ActiveWindow = window; },
    Panes: { Count: 1, Item: () => canvasPane } };
  const view = {
    get Zoom() { return zoom; },
    set Zoom(value) { if (window.ViewType === 1 && window.ActivePane === canvasPane) { zoom = value; fit = 0; } },
    get ZoomToFit() { return fit; },
    set ZoomToFit(value) { if (window.ViewType === 1 && window.ActivePane === canvasPane) { fit = value; if (value === -1) zoom = 62; } },
  };
  window.View = view;
  const show = { playing, starts: 0, frames: [], frame: null, page: 1 };
  const pickerControls = controls(pickerItems), pickerBControls = controls(pickerBItems);
  const groups = { 1: form, 2: pickerControls, 3: pickerBControls, 4: {} };
  const getPickerControls = page => groups[page];
  for (const group of Object.values(groups)) for (const control of Object.values(group)) {
    Object.defineProperty(control, 'Visible', { get: () => true, set() { throw new Error('do not toggle native visibility'); } });
  }
  const shapesByPage = { 1: [], 2: [], 3: [], 4: [] }, slides = {};
  for (const page of [1, 2, 3, 4]) {
    for (const c of Object.values(groups[page])) shapesByPage[page].push({ Name: c.Name, Left: c.Left, Top: c.Top, Width: c.Width, Height: c.Height, Fill: { paints: [] } });
    if (page === 2 || page === 3) for (const Name of ['TsColorPlaneImage', 'TsColorHueImage']) {
      shapesByPage[page].push({ Name, get Visible() { return -1; }, set Visible(v) { throw new Error('do not toggle shapes'); },
        Fill: { paints: [], UserPicture(path) {
          assert.notEqual(page, show.page, 'stage gradients on a page that is not being displayed');
          this.paints.push(path); if (!files.has(path)) throw new Error('missing BMP');
        } }, Delete() { throw new Error('do not delete packaged shapes'); },
      });
    }
    const shapes = shapesByPage[page];
    slides[page] = { Shapes: { AddPicture() { throw new Error('do not add pictures'); }, AddShape() { throw new Error('do not add shapes'); },
      get Count() { return shapes.length; }, Item(name) {
        const found = typeof name === 'number' ? shapes[name - 1] : shapes.find(s => s.Name === name);
        if (!found) throw new Error('missing shape '+name); return found;
      },
    } };
  }
  const shapes = Object.values(shapesByPage).flat();
  const paletteForm = { plane: PLANE, hue: HUE, ...pickerControls, shows: 0, hides: 0, Show() { this.shows++; }, Hide() { this.hides++; } };
  const showWindow = { View: {
    get CurrentShowPosition() { return show.page; },
    Exit() { throw new Error('color selection must remain in playback'); },
    GotoSlide(index, reset) {
      assert.equal(show.playing, true); assert.equal(reset, -1);
      assert.notEqual(index, show.page, 'never reload the current slide');
      assert.ok([1, 2, 3, 4].includes(index));
      const snapshot = group => Object.fromEntries(Object.entries(group).map(([name, c]) =>
        [name, { Visible: c.Visible, Caption: c.Caption, Text: c.Text, BackColor: c.BackColor, Left: c.Left, Top: c.Top, Width: c.Width, Height: c.Height, FontSize: c.Font.Size }]));
      show.page = index;
      show.frame = { page: index, controls: snapshot(groups[index]),
        shapes: shapesByPage[index].map(s => ({ Name: s.Name, Left: s.Left, Top: s.Top, paint: s.Fill.paints.at(-1) })) };
      show.frames.push(show.frame);
    },
  } };
  const reads = [];
  const app = {
    Env: { GetTempPath: () => '/tmp/' },
    ActivePresentation: { Slides: { Item: page => slides[page] }, FullName: '/tools/编辑器.pptm', Windows: { Item: () => window },
      get SlideShowWindow() { return show.playing ? showWindow : null; },
      SlideShowSettings: { Run() { show.starts++; show.playing = true; return showWindow; } },
    }, ActiveWindow: { View: { Zoom: 10 } },
    Presentations: { Count: 0, Item() { throw new Error('no open documents'); }, Open() { throw new Error('do not open'); } },
    FileDialog: () => ({ Filters: { Clear() {}, Add() {} }, Show: () => -1, SelectedItems: { Item: () => file } }),
    FileSystem: {
      readFileString: () => randomUUID(),
      readAsBinaryString: p => { reads.push(p); return binaryString(files.get(p)); },
      writeAsBinaryString: (p, data) => files.set(p, Uint8Array.from(data, c => c.charCodeAt(0))),
      Exists: p => files.has(p), Remove: p => files.delete(p),
    },
  };
  const editor = new SlideEditor(app, form, () => new NativeColorPicker(paletteForm)); editor.initialize();
  return { file, original, files, form, app, editor, view, window, paletteForm, pickerControls, pickerBControls, getPickerControls, shapes, shapesByPage, show, showWindow, reads };
}

test('slide controls open, create, copy, recolor and save an external PPTX without changing the tool', async () => {
  const m = await mock(), e = m.editor, c = m.form;
  assert.equal(c.Scope.options.length, 3);
  assert.equal(c.SaveStyle.Enabled, false);
  await e.dispatch('open'); assert.equal(c.ErrorText.Caption, '');
  c.Preset.Text = '红色'; await e.dispatch('new');
  assert.equal(c.ErrorText.Caption, '');
  assert.equal(c.StyleName.Text, '我的红色');
  const originalID = e.selected.id;
  await e.dispatch('duplicate'); assert.equal(c.ErrorText.Caption, '');
  assert.notEqual(e.selected.id, originalID);
  e.setColor('Fill', '#9A3726');
  assert.equal(c.Preview_0_0.BackColor, 0x9A3726);
  c.TextColor.Text = '#FFFFFF'; await e.dispatch('fields');
  c.SetDefaultFont.Value = true; c.DefaultFontSize.Text = '22';
  await e.dispatch('save'); assert.equal(c.ErrorText.Caption, '');
  let deck = await openDeck(m.files.get(m.file));
  let style = deck.styles.find(s => s.id === deck.defaultID);
  assert.equal(readPart(style.xml, 'firstRow').fill, '#9A3726');
  assert.equal(deck.fontDefaults.points, 22);
  assert.equal(deck.tables.every(t => t.styleID === style.id), true);
  assert.equal([...m.files.keys()].filter(p => p.includes('表格样式备份')).length, 1);
  assert.equal(m.app.ActivePresentation.FullName, '/tools/编辑器.pptm');
  e.setColor('Fill', '#166534'); await e.dispatch('save');
  assert.equal(c.ErrorText.Caption, '');
  deck = await openDeck(m.files.get(m.file)); style = deck.styles.find(s => s.id === deck.defaultID);
  assert.equal(readPart(style.xml, 'firstRow').fill, '#166534');
});

test('slide editor refuses an open target and preserves the current target when open is cancelled', async () => {
  const m = await mock();
  await m.editor.dispatch('open');
  const target = m.editor.host.path;
  m.app.FileDialog = () => ({ Filters: { Clear() {}, Add() {} }, Show: () => 0 });
  await m.editor.dispatch('open'); assert.equal(m.editor.host.path, target);
  m.app.Presentations = { Count: 1, Item: () => ({ FullName: m.file }) };
  await m.editor.dispatch('save'); assert.match(m.form.ErrorText.Caption, /先关闭目标/);
  assert.deepEqual(m.files.get(m.file), m.original);
  assert.equal(m.files.size, 1);
});

test('packaged document events resolve slide globals, initialize once and dispatch first open without UserForm', async () => {
  const m = await mock(), dir = await mkdtemp(path.join(tmpdir(), 'wps-slide-macro-'));
  try {
    const bundlePath = path.join(dir, 'runtime.js');
    await buildSlideMacro(bundlePath, []);
    const bundle = await readFile(bundlePath, 'utf8');
    const source = (await readFile(new URL('../macro/slide-code.js', import.meta.url), 'utf8'))
      .replace('__TABLE_STYLES_VERSION__', '"test"');
    const runtime = 'function TableStylesNativeRuntime() {\n'+bundle+'\nreturn NativeSlideTableStyles;\n}';
    const messages = [];
    const globals = Object.fromEntries([...items.map(i => [i.native, m.form[i.name]]), ...pickerItems.map(i => [i.native, m.pickerControls[i.name]]), ...pickerBItems.map(i => [i.native, m.pickerBControls[i.name]])]);
    const context = vm.createContext({ Application: m.app, ...globals, MsgBox: s => messages.push(s) });
    // Native controls can emit Change while Clear/AddItem is running.
    const clear = m.form.Preset.Clear;
    m.form.Preset.Clear = function () { clear.call(this); context.TsFill_Change(); };
    for (const module of [source, eventCode+'\n'+pickerEvents, runtime]) new vm.Script(module).runInContext(context);
    await context.TsOpenFile_Click();
    assert.equal(m.form.ErrorText.Caption, '');
    assert.equal(m.form.SaveStyle.Enabled, true);
    const editor = context.TableStylesSlideEditor.cached;
    assert.equal(context.OpenTableStyleEditor(), true);
    assert.equal(context.TableStylesSlideEditor.cached, editor);
    await context.TsNewStyle_Click(); assert.equal(m.form.StyleName.Text, '我的绿色');
    await context.TsPickFill_Click(); assert.equal(m.show.page, 2, m.form.ErrorText.Caption);
    assert.equal(m.show.frame.controls.Plane.Visible, true);
    m.pickerControls.Hex.Text = '#C00000'; context.TsPickerHex_Change(); await context.TsPickerHex_KeyUp(13, 0);
    assert.equal(m.show.page, 3); await context.TsPickerBApply_Click();
    assert.equal(m.form.Fill.Text, '#C00000');
    await context.TsPickColor_left_Click();
    m.pickerControls.Hex.Text = '#AB2783'; context.TsPickerHex_Change(); await context.TsPickerApply_Click();
    assert.equal(m.form.Color_left.Text, '#AB2783');
    assert.equal(m.form.Mode_left.Text, '线条');
    await context.TsSaveStyle_Click(); assert.equal(m.form.ErrorText.Caption, '');
    assert.equal([...m.files.keys()].filter(p => p.includes('备份')).length, 1);
    assert.equal(messages.length, 0);
  } finally { await rm(dir, { recursive: true, force: true }); }
});

test('border width, color and dash edits turn inherited borders into explicit lines and survive saving', async () => {
  const { editor: e, form: c, files, file } = await mock();
  await e.dispatch('open');
  assert.equal(c.Mode_left.Text, '继承');
  c.Width_left.Text = '2'; await e.dispatch('border:left');
  assert.equal(c.Mode_left.Text, '线条');
  c.Width_right.Text = '3.25'; await e.dispatch('border:right');
  e.setColor('Color_left', '#E91E63');
  c.Dash_bottom.Text = '虚线'; await e.dispatch('border:bottom');
  // Save must also capture edits whose native Change event has not fired yet.
  c.Width_top.Text = '2.5';
  await e.dispatch('save'); assert.equal(c.ErrorText.Caption, '');
  const deck = await openDeck(files.get(file)), s = deck.styles.find(s => s.id === deck.defaultID);
  const borders = readPart(s.xml, 'firstRow').borders;
  assert.equal(borders.left.width, 2); assert.equal(borders.left.color, '#E91E63');
  assert.equal(borders.right.width, 3.25); assert.equal(borders.top.width, 2.5);
  assert.equal(borders.bottom.mode, 'solid'); assert.equal(borders.bottom.dash, 'dash');
  assert.equal(c.Width_left.Text, '2'); assert.equal(c.Width_right.Text, '3.25');
  assert.equal(c.Width_top.Text, '2.5');
  assert.ok(c.Color_left.options.includes('#E91E63'), 'custom colors must be added before selecting in native noneditable combo');
  c.Mode_left.Text = '继承'; await e.dispatch('mode:left'); await e.dispatch('save');
  assert.equal(c.ErrorText.Caption, '');
  assert.equal(readPart((await openDeck(files.get(file))).styles.find(s => s.id === e.selected.id).xml, 'firstRow').borders.left.mode, 'inherit');
});

test('all eight color fields can use the palette and cancel preserves the previous value', async () => {
  const { editor: e, form: c, paletteForm: p } = await mock();
  await e.dispatch('open');
  for (const [i, name] of ['Fill', 'TextColor', 'Color_left', 'Color_right', 'Color_top', 'Color_bottom', 'Color_insideH', 'Color_insideV'].entries()) {
    const before = c[name].Text;
    await e.dispatch('pick:'+name); e.pickerAction('swatch', 4); e.pickerAction('cancel');
    assert.equal(c[name].Text, before);
    await e.dispatch('pick:'+name);
    const hex = '#'+(0x126400+i).toString(16).toUpperCase();
    p.Hex.Text = hex; e.pickerAction('hex'); e.pickerAction('apply');
    assert.equal(c[name].Text, hex); assert.ok(c[name].options.includes(hex));
    assert.equal(c['Pick'+name].BackColor, parseInt(hex.slice(1), 16));
  }
});

test('typing or choosing local field values does not disable and recreate native controls', async () => {
  const { editor: e, form: c } = await mock(); await e.dispatch('open');
  const changes = [];
  for (const name of ['Fill', 'Width_left', 'Color_left', 'Dash_left']) {
    let enabled = c[name].Enabled;
    Object.defineProperty(c[name], 'Enabled', { get: () => enabled, set(value) { changes.push([name, value]); enabled = value; } });
  }
  c.Width_left.Text = '2'; await e.dispatch('border:left');
  c.Fill.Text = '#C00000'; await e.dispatch('fields');
  assert.equal(c.Width_left.Text, '2'); assert.deepEqual(changes, []);
});

test('invalid widths stop before writing', async () => {
  const m = await mock(), { editor: e, form: c, files, file, original } = m;
  await e.dispatch('open'); c.Width_left.Text = '0'; await e.dispatch('border:left');
  await e.dispatch('save'); assert.match(c.ErrorText.Caption, /0.25–20/);
  assert.deepEqual(files.get(file), original); assert.equal(files.size, 1);

});

test('BMP color surfaces use correct channels, padding, orientation and HSV endpoints', () => {
  const data = gradientBitmap(3, 3, 120), v = new DataView(data.buffer);
  assert.equal(String.fromCharCode(...data.subarray(0, 2)), 'BM');
  assert.equal(v.getUint32(2, true), data.length);
  assert.equal(v.getInt32(18, true), 3); assert.equal(v.getInt32(22, true), 3);
  const rgb = (x, y) => [...data.subarray(54 + (2 - y) * 12 + x * 3, 57 + (2 - y) * 12 + x * 3)].reverse();
  assert.deepEqual(rgb(0, 0), [255, 255, 255]); assert.deepEqual(rgb(2, 0), [0, 255, 0]);
  assert.deepEqual(rgb(0, 2), [0, 0, 0]); assert.deepEqual(rgb(2, 2), [0, 0, 0]);
  assert.deepEqual(rgb(1, 1), [64, 128, 64]);
});

function surfaceFor(m, report) {
  return new SlideColorSurface(m.app, m.app.ActivePresentation, m.getPickerControls, report);
}

test('picker pages preserve static controls and shape geometry without visibility changes or shape creation', async () => {
  const m = await mock(), surface = surfaceFor(m), p = new NativeColorPicker(surface);
  await p.open('#74C445', () => {});
  assert.equal(m.show.page, 2); assert.equal(m.show.frame.controls.Current.BackColor, 0x74C445);
  assert.equal(m.form.OpenFile.Visible, true, 'main controls stay intact on their own page');
  assert.equal(m.show.frame.controls.OpenFile, undefined);
  for (const x of [40, 160, 320]) {
    p.dispatch('pointer', { surface: 'hue', phase: 'down', button: 1, x, y: 12 });
    await p.dispatch('pointer', { surface: 'hue', phase: 'up', button: 0, x, y: 12 });
    assert.equal(surface.ErrorText.Caption.trim(), '');
  }
  assert.deepEqual(m.show.frames.map(f => f.page), [2, 3, 2, 3]);
  assert.equal(m.shapes.filter(s => ['TsColorPlaneImage', 'TsColorHueImage'].includes(s.Name)).length, 4); assert.equal(m.files.size, 1);
  for (const group of [m.pickerControls, m.pickerBControls]) {
    for (const name of ['Original', 'Current']) assert.equal(group[name].Caption, '\u00a0');
  }
  await p.dispatch('cancel'); assert.equal(m.show.page, 1);
  assert.ok(m.show.frame.controls.OpenFile); assert.equal(m.show.frame.controls.Plane, undefined);
});

test('opening a color page preserves playback and never activates the edit canvas', async () => {
  const m = await mock(); await m.editor.dispatch('open');
  const surface = surfaceFor(m); m.editor.pickerFactory = () => new NativeColorPicker(surface);
  await m.editor.dispatch('pick:Fill');
  assert.equal(m.show.playing, true); assert.equal(m.editor.busy, false); assert.equal(surface.visible, true);
  assert.equal(m.window.ViewType, 9); assert.notEqual(m.app.ActiveWindow, m.window);
  assert.equal(m.show.starts, 0); assert.equal(m.show.page, 2); assert.equal(m.form.ErrorText.Caption, '');
});

test('opening a picker from editing starts the show then enters its dedicated page', async () => {
  const m = await mock({ playing: false }), surface = surfaceFor(m);
  await new NativeColorPicker(surface).open('#74C445', () => {});
  assert.equal(m.show.starts, 1); assert.equal(m.show.page, 2); assert.equal(m.window.ViewType, 9);
});

test('mouse release switches color pages and apply or cancel returns to the intact main form', async () => {
  const m = await mock(); await m.editor.dispatch('open');
  const surface = surfaceFor(m); m.editor.pickerFactory = () => new NativeColorPicker(surface);
  await m.editor.dispatch('pick:Fill');
  const before = m.form.Fill.Text, frame = m.show.frame;
  for (const [phase, x, y] of [['down', 220, 12], ['move', 100, 12]]) {
    m.editor.pickerAction('pointer', { surface: 'hue', phase, button: 1, x, y }, 2);
    assert.equal(m.show.frame, frame, 'navigation must not interrupt dragging');
  }
  await m.editor.pickerAction('pointer', { surface: 'hue', phase: 'up', button: 0, x: 100, y: 12 }, 2);
  assert.equal(m.show.page, 3); assert.equal(m.show.frame.controls.Current.BackColor, surface.Current.BackColor);
  await m.editor.pickerAction('cancel', null, 3);
  assert.equal(m.form.Fill.Text, before); assert.equal(m.show.page, 1);
  await m.editor.dispatch('pick:Fill');
  surface.Hex.Text = '#74C445'; m.editor.pickerAction('hex', null, 2);
  await m.editor.pickerAction('commit', null, 2); assert.equal(m.show.page, 3);
  await m.editor.pickerAction('apply', null, 3);
  assert.equal(m.show.page, 1); assert.equal(m.show.frame.controls.Fill.Text, '#74C445');
  assert.equal(m.show.frame.controls.PickFill.BackColor, 0x74C445);
});

test('typing keeps focus on the current page; Enter switches once and keeps incomplete or invalid inputs for correction', async () => {
  const m = await mock(), surface = surfaceFor(m), p = new NativeColorPicker(surface);
  await p.open('#C00000', () => {});
  const initial = m.show.frame;
  for (const text of ['#', '#12', '#1264A8']) {
    surface.Hex.Text = text; p.dispatch('hex'); assert.equal(m.show.frame, initial);
  }
  await p.dispatch('commit'); assert.equal(m.show.page, 3);
  assert.equal(m.show.frame.controls.Current.BackColor, 0x1264A8);
  surface.Hex.Text = '#12'; p.dispatch('hex'); await p.dispatch('commit');
  assert.equal(m.show.page, 2); assert.match(m.show.frame.controls.ErrorText.Caption, /六位/);
  assert.equal(surface.Hex.Text, '#12'); assert.equal(p.color, '#1264A8');
  surface.R.Text = '256'; p.dispatch('rgb'); await p.dispatch('commit');
  assert.match(m.show.frame.controls.ErrorText.Caption, /0–255/); assert.equal(surface.R.Text, '256');
  surface.R.Text = '18'; surface.G.Text = '52'; surface.B.Text = '86'; p.dispatch('rgb');
  const frame = m.show.frame;
  await p.dispatch('commit'); assert.notEqual(m.show.frame, frame);
  assert.equal(m.show.frame.controls.Current.BackColor, 0x123456);
});

test('only events from the displayed picker page are accepted; navigation callbacks cannot edit the draft', async () => {
  const m = await mock(); await m.editor.dispatch('open');
  const surface = surfaceFor(m); m.editor.pickerFactory = () => new NativeColorPicker(surface);
  await m.editor.dispatch('pick:Fill');
  const p = m.editor.picker, original = p.color;
  m.pickerBControls.Hex.Text = '#AB2783'; m.editor.pickerAction('hex', null, 3);
  assert.equal(p.color, original);
  const nativeGoto = m.showWindow.View.GotoSlide;
  m.showWindow.View.GotoSlide = (page, reset) => {
    assert.equal(surface.syncing, true);
    m.editor.pickerAction('apply', null, 2); m.editor.pickerAction('cancel', null, 3);
    nativeGoto(page, reset);
  };
  surface.Hex.Text = '#1264A8'; m.editor.pickerAction('hex', null, 2);
  await m.editor.pickerAction('commit', null, 2);
  assert.equal(m.show.page, 3); assert.equal(surface.visible, true); assert.equal(p.color, '#1264A8');
  m.pickerControls.Hex.Text = '#123456'; m.editor.pickerAction('hex', null, 2);
  assert.equal(p.color, '#1264A8');
  await m.editor.pickerAction('cancel', null, 3);
  assert.equal(surface.acceptEvent(3), false); assert.equal(m.show.page, 1);
});

test('both picker pages generate Enter and mouse events carrying the correct page number', () => {
  const actions = [], context = vm.createContext({ TableStylesColorAction: (...args) => actions.push(args) });
  new vm.Script(pickerEvents).runInContext(context);
  for (const [prefix, page] of [['TsPicker', 2], ['TsPickerB', 3]]) {
    for (const field of ['Hex', 'R', 'G', 'B']) {
      context[prefix+field+'_KeyUp'](65, 0); assert.equal(actions.length, 0);
      for (const code of [13, 16777220, 16777221]) context[prefix+field+'_KeyUp'](code, 0);
      assert.deepEqual(actions.splice(0), [['commit', null, page], ['commit', null, page], ['commit', null, page]]);
    }
    context[prefix+'Hue_MouseUp'](0, 0, 10, 12);
    assert.equal(actions[0][0], 'pointer'); assert.equal(actions[0][2], page);
    assert.equal(actions[0][1].phase, 'up'); actions.length = 0;
  }
});

test('failed navigation reports the error and preserves editor colors, playback and target bytes', async () => {
  const m = await mock(); await m.editor.dispatch('open'); m.editor.setColor('Fill', '#AB2783');
  const errors = [], surface = surfaceFor(m, e => errors.push(e.message));
  m.editor.pickerFactory = () => new NativeColorPicker(surface);
  m.showWindow.View.GotoSlide = () => { throw new Error('方法执行失败'); };
  await m.editor.dispatch('pick:Fill');
  assert.match(surface.lastRefreshError, /无法切换色盘页面.*方法执行失败/);
  assert.equal(errors.length, 1); assert.equal(m.editor.busy, false); assert.equal(surface.syncing, false);
  assert.equal(m.form.Fill.Text, '#AB2783'); assert.equal(m.show.page, 1); assert.equal(m.show.playing, true);
  assert.deepEqual(m.files.get(m.file), m.original); assert.equal(m.files.size, 1);
});

test('navigation reads the actual current page after a manual page change', async () => {
  const m = await mock(), surface = surfaceFor(m), p = new NativeColorPicker(surface);
  m.show.page = 2;
  await p.open('#C00000', () => {});
  assert.equal(m.show.page, 3, 'do not reload page 2 even if the user entered it manually');
});

test('a failed return keeps the picker interactive and allows cancel to be retried', async () => {
  const m = await mock(), errors = [], surface = surfaceFor(m, e => errors.push(e.message));
  const p = new NativeColorPicker(surface); await p.open('#C00000', () => {});
  const nativeGoto = m.showWindow.View.GotoSlide;
  m.showWindow.View.GotoSlide = () => { throw new Error('temporary failure'); };
  await p.dispatch('cancel');
  assert.equal(surface.visible, true); assert.equal(surface.acceptEvent(2), true);
  assert.equal(surface.refreshTask, null); assert.equal(errors.length, 1);
  m.showWindow.View.GotoSlide = nativeGoto; await p.dispatch('cancel');
  assert.equal(m.show.page, 1); assert.equal(surface.visible, false);
});

test('dragging the gradient uses actual control dimensions without moving controls, including Qt Button=0', async () => {
  const m = await mock();
  for (const group of [m.pickerControls, m.pickerBControls]) for (const c of Object.values(group)) {
    c.Left *= 2; c.Top *= 2; c.Width *= 2; c.Height *= 2;
    c.Move = () => { throw new Error('picker controls must remain fixed'); };
  }
  const surface = surfaceFor(m), p = new NativeColorPicker(surface), applied = [];
  await p.open('#74C445', c => applied.push(c));
  const plane = surface.Plane, initial = p.color;
  p.dispatch('pointer', { surface: 'plane', control: 'Plane', phase: 'down', button: 1, x: 100, y: 20 });
  p.dispatch('pointer', { surface: 'plane', control: 'Plane', phase: 'move', button: 0, x: 400, y: 140 });
  assert.equal(p.hsv[1], 400 / plane.Width); assert.equal(p.hsv[2], 1 - 140 / plane.Height);
  assert.notEqual(p.color, initial); assert.equal(m.show.frames.length, 1);
  const selected = p.color;
  await p.dispatch('pointer', { surface: 'plane', control: 'Plane', phase: 'up', button: 0, x: 400, y: 140 });
  assert.equal(p.color, selected); assert.equal(m.show.page, 3);
  await p.dispatch('apply'); assert.deepEqual(applied, [selected]);
});

test('real numeric step actions change decimal widths and optional default size, respect bounds and persist', async () => {
  const m = await mock(), e = m.editor; await e.dispatch('open');
  for (const edge of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
    assert.equal(e.c('Width_' + edge).Text, '1');
    await e.dispatch('step:Width_' + edge + ':+1'); assert.equal(e.c('Width_' + edge).Text, '1.25');
    await e.dispatch('step:Width_' + edge + ':-1'); assert.equal(e.c('Width_' + edge).Text, '1');
    await e.dispatch('step:Width_' + edge + ':-1');
    assert.equal(e.c('Width_' + edge).Text, '0.75');
    await e.dispatch('step:Width_' + edge + ':+1'); assert.equal(e.c('Width_' + edge).Text, '1');
    e.c('Width_' + edge).Text = '0.25'; await e.dispatch('step:Width_' + edge + ':-1');
    assert.equal(e.c('Width_' + edge).Text, '0.25');
    e.c('Width_' + edge).Text = '2'; await e.dispatch('step:Width_' + edge + ':+1');
    assert.equal(e.c('Width_' + edge).Text, '2.25'); assert.equal(e.c('Mode_' + edge).Text, '线条');
    await e.dispatch('step:Width_' + edge + ':-1'); assert.equal(e.c('Width_' + edge).Text, '2');
    await e.dispatch('step:Width_' + edge + ':+1');
  }
  e.c('Width_right').Text = '20'; await e.dispatch('step:Width_right:+1'); assert.equal(e.c('Width_right').Text, '20');
  e.c('Width_bottom').Text = '0.25'; await e.dispatch('step:Width_bottom:-1'); assert.equal(e.c('Width_bottom').Text, '0.25');
  e.c('Width_bottom').Text = '0.1'; await e.dispatch('step:Width_bottom:-1'); assert.equal(e.c('Width_bottom').Text, '0.25');
  e.c('Width_right').Text = '2.25'; await e.dispatch('border:right');
  e.c('Width_bottom').Text = '2.25'; await e.dispatch('border:bottom');
  const oldSize = e.c('DefaultFontSize').Text;
  await e.dispatch('step:DefaultFontSize:+1'); assert.equal(e.c('DefaultFontSize').Text, oldSize);
  e.c('SetDefaultFont').Value = true; await e.dispatch('font');
  e.c('DefaultFontSize').Text = '18'; await e.dispatch('step:DefaultFontSize:+1');
  assert.equal(e.c('DefaultFontSize').Text, '19');
  await e.dispatch('save'); assert.equal(e.c('ErrorText').Caption, '');
  const deck = await openDeck(m.files.get(m.file)), style = deck.styles.find(s => s.id === deck.defaultID);
  for (const edge of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) assert.equal(readPart(style.xml, 'firstRow').borders[edge].width, 2.25);
  assert.equal(deck.fontDefaults.points, 19);
});

test('initial dropdown population is bounded and full colors load once on demand without losing the value', async () => {
  const m = await mock(), e = m.editor;
  const initialCalls = Object.values(m.form).reduce((n, c) => n + c.addCalls, 0);
  assert.ok(initialCalls < 130, String(initialCalls)); assert.equal(m.form.Color_left.options.length, 1);
  await e.dispatch('open');
  e.setColor('Color_left', '#AB2783'); const calls = m.form.Color_left.addCalls;
  await e.dispatch('colors:Color_left');
  assert.equal(m.form.Color_left.Text, '#AB2783');
  assert.ok(m.form.Color_left.options.includes('theme:accent6'));
  assert.ok(m.form.Color_left.addCalls > calls); assert.equal(m.form.Color_right.options.length, 1);
  const loadedCalls = m.form.Color_left.addCalls; await e.dispatch('colors:Color_left');
  assert.equal(m.form.Color_left.addCalls, loadedCalls);
});

test('opening validates and inspects one file read, then discards the parse cache for future inspections', async () => {
  const m = await mock(); await m.editor.dispatch('open');
  assert.equal(m.reads.filter(p => p === m.file).length, 1);
  assert.equal(m.editor.host.openInspection, null);
  await m.editor.host.inspect(); assert.equal(m.reads.filter(p => p === m.file).length, 2);
});

test('cancelling the first file chooser does not load the runtime or initialize color dropdowns', async () => {
  const m = await mock(); let shows = 0, runtimes = 0;
  const source = (await readFile(new URL('../macro/slide-code.js', import.meta.url), 'utf8')).replace('__TABLE_STYLES_VERSION__', '"test"');
  m.app.FileDialog = () => ({ Filters: { Clear() {}, Add() {} }, Show() { shows++; return 0; } });
  // Deliberately omit all native globals: the chooser must precede even their resolution.
  const context = vm.createContext({ Application: m.app, TableStylesNativeRuntime() { runtimes++; throw new Error('must not initialize'); } });
  for (const code of [source, eventCode]) new vm.Script(code).runInContext(context);
  const calls = m.form.Fill.addCalls; await context.TsOpenFile_Click();
  assert.equal(shows, 1); assert.equal(runtimes, 0); assert.equal(m.form.Fill.addCalls, calls);
  assert.equal(context.TableStylesSlideEditor.cached, undefined); assert.equal(m.form.OpenFile.Caption, '打开 PPTX');
});

test('the fixed editor preserves input and ignores delayed main-page events while its picker is open', async () => {
  const m = await mock(); await m.editor.dispatch('open');
  m.form.StyleName.Text = '固定布局样式'; await m.editor.dispatch('fields');
  const geometry = Object.values(m.form).map(c => [c.Left, c.Top, c.Width, c.Height, c.Font.Size]);
  const surface = surfaceFor(m); m.editor.pickerFactory = () => new NativeColorPicker(surface);
  await m.editor.dispatch('pick:Fill'); assert.equal(m.show.page, 2);
  assert.equal(m.editor.acceptEvent(1), false);
  surface.Hex.Text = '#123456'; m.editor.pickerAction('hex', null, 2);
  await m.editor.pickerAction('apply', null, 2);
  assert.equal(m.show.page, 1); assert.equal(m.show.frame.controls.Fill.Text, '#123456');
  assert.equal(m.form.StyleName.Text, '固定布局样式'); assert.equal(m.editor.acceptEvent(1), true);
  assert.deepEqual(Object.values(m.form).map(c => [c.Left, c.Top, c.Width, c.Height, c.Font.Size]), geometry);
});
