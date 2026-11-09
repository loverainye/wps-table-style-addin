import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { FileHost } from '../src/file-host.js';
import { PARTS, EDGES, BORDER_MIN, BORDER_MAX, guid, preset, readPart, editStyle, parseXML, colorForPreview } from '../src/styles.js';
import { macroApplication } from './platform-services.js';
import { setUUIDProvider } from '../src/uuid.js';

// JSA is a JavaScript runtime, without browser DOM or Web Crypto. Supply the
// XML implementation with the file; platform services supply files and UUIDs.
// These globals are installed only when the macro is run, never at module scope.
export function installRuntime(app) {
  const hostApp = macroApplication(app);
  globalThis.DOMParser = class {
    parseFromString(text, type) {
      return new DOMParser({ onError(level, message) {
        if (level !== 'warning') throw new Error('文稿中的 XML 无法解析：' + message);
      } }).parseFromString(text, type);
    }
  };
  globalThis.XMLSerializer = XMLSerializer;
  setUUIDProvider(() => hostApp.TableStylePlatform.uuid());
  return hostApp;
}

const THEMES = { dk1: '#202020', lt1: '#FFFFFF', dk2: '#203864', lt2: '#E7E6E6', accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47', tx1: '#202020', tx2: '#203864', bg1: '#FFFFFF', bg2: '#E7E6E6', hlink: '#0563C1', folHlink: '#954F72' };
const FONT_MODES = ['inherit', 'on', 'off'];
const BORDER_MODES = ['inherit', 'solid', 'none', 'unsupported'];
const DASHES = ['solid', 'dash', 'dot', 'dashDot', 'lgDash'];
const SCOPES = ['all', 'linked', 'none', 'selected'];
const FLAGS = ['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'];
// Native control properties take 0xRRGGBB. Their serialized XML uses ARGB;
// passing that XML value through the macro API turns the preview black.
const rgb = value => parseInt(value.slice(1), 16);

export const previewColor = value => colorForPreview(value, THEMES, '#FFFFFF');

export function colorValue(text, fill = false) {
  const s = String(text).trim();
  if (s === '继承' || s === 'inherit') return 'inherit';
  if (fill && ['无填充', 'none'].includes(s)) return 'none';
  if (s === '保留现有' || s === 'unsupported') return 'unsupported';
  if (/^#?[0-9a-f]{6}$/i.test(s)) return '#' + s.replace(/^#/, '').toUpperCase();
  if (/^theme:(accent[1-6]|dk[12]|lt[12]|tx[12]|bg[12]|hlink|folHlink)$/.test(s)) return s;
  throw new Error('颜色请输入六位十六进制值，例如 #C00000；也可选择“继承”或主题颜色。');
}
const colorText = v => ({ inherit: '继承', none: '无填充', unsupported: '保留现有' }[v] || v);

export class NativeEditor {
  constructor(app, form, geometry = []) {
    this.form = form; this.host = new FileHost(app); this.geometry = geometry;
    this.context = null; this.selected = null; this.newStyles = []; this.drafts = new Map();
    this.patches = {}; this.part = 'firstRow'; this.loaded = {}; this.busy = false; this.updating = false; this.zoom = 100;
    this.choices = {};
  }
  c(name) { return this.form[name]; }
  text(name) { return String(this.c(name).Text || ''); }
  index(name) {
    const index = (this.choices[name] || []).indexOf(this.text(name));
    if (index < 0) throw new Error('请从下拉列表选择有效选项：' + name);
    return index;
  }
  select(name, index) {
    const text = (this.choices[name] || [])[index];
    if (text === undefined) throw new Error('下拉选项无效：' + name);
    this.c(name).Text = text;
  }
  checked(name) { return this.c(name).Value === true || Number(this.c(name).Value) === 1; }
  list(name, labels, index = 0) {
    const c = this.c(name); c.Clear();
    this.choices[name] = [...labels];
    for (const label of labels) c.AddItem(label);
    if (labels.length) this.select(name, index); else c.Text = '';
  }
  status(error = '') { this.c('ErrorText').Caption = error; }
  initialize() {
    this.updating = true;
    try {
      this.list('Preset', ['绿色', '红色', '蓝色', '空白']);
      this.list('Part', PARTS.map(x => x[1]), PARTS.findIndex(x => x[0] === this.part));
      this.list('Scope', ['所有表格', '使用此样式的表格', '不更改现有表格', '当前选中的表格']);
      this.list('Bold', ['继承', '加粗', '不加粗']); this.list('Italic', ['继承', '斜体', '不斜体']);
      const colors = ['继承', '#FFFFFF', '#20242A', '#385723', '#C00000', '#2455A4', ...Object.keys(THEMES).map(k => 'theme:' + k), '保留现有'];
      this.list('FillChoice', ['继承', '无填充', ...colors.slice(1)]); this.list('TextColorChoice', colors);
      for (const [edge] of EDGES) {
        this.list('Mode_' + edge, ['继承', '线条', '无边框', '保留现有']);
        this.list('Dash_' + edge, ['实线', '虚线', '点线', '点划线', '长虚线']);
        this.c('Color_' + edge).Text = '#FFFFFF'; this.c('Width_' + edge).Text = '1';
      }
      this.c('SetDefault').Value = true; this.c('FirstPlaceholder').Caption = '';
      this.c('DefaultFontSize').Text = '18';
      this.status(''); this.enabled();
    } finally { this.updating = false; }
  }
  enabled() {
    const editable = !!this.selected?.xml && !this.selected.builtin;
    for (const name of ['StyleName', 'Part', 'Fill', 'FillChoice', 'TextColor', 'TextColorChoice', 'Bold', 'Italic', 'AllBorders']) this.c(name).Enabled = editable && !this.busy;
    for (const [edge] of EDGES) for (const prefix of ['Mode_', 'Color_', 'Width_', 'Dash_']) this.c(prefix + edge).Enabled = editable && !this.busy;
    for (const name of ['Styles', 'NewStyle', 'Preset', 'SetDefault', 'Scope', 'ClearOverrides', ...FLAGS.map(x => 'Flag_' + x)]) this.c(name).Enabled = !!this.context && !this.busy;
    this.c('DuplicateStyle').Enabled = editable && !this.busy;
    this.c('SaveStyle').Enabled = !!this.selected && !this.busy;
    this.c('SetDefaultFont').Enabled = !!this.context?.fontDefaults.masterCount && !this.busy;
    this.c('DefaultFontSize').Enabled = this.c('SetDefaultFont').Enabled && this.checked('SetDefaultFont');
    this.c('OpenFile').Enabled = !this.busy; this.c('UseCurrent').Enabled = !this.busy;
    this.c('FirstPlaceholder').Caption = !this.context ? '点击“打开 PPTX”选择要修改的文件。' : this.selected?.builtin ? '此为内置样式。要修改颜色，请新建一个自定义样式。' : '';
  }
  allStyles() { return [...(this.context?.styles || []), ...this.newStyles]; }
  editedXML() { return this.selected?.xml ? editStyle(this.selected.xml, this.text('StyleName') || this.selected.name, this.selected.id, this.patches) : null; }
  flags() { return Object.fromEntries(FLAGS.map(key => [key, this.checked('Flag_' + key)])); }
  capture(strict = true) {
    if (!this.selected?.xml || this.updating) return;
    const patch = {};
    const attempt = fn => { try { fn(); } catch (e) { if (strict) throw e; } };
    for (const [name, key, fill] of [['Fill', 'fill', true], ['TextColor', 'color', false]]) {
      if (this.text(name) !== this.loaded[name]) attempt(() => {
        const value = colorValue(this.text(name), fill); if (value !== 'unsupported') patch[key] = value;
      });
    }
    for (const [name, key] of [['Bold', 'bold'], ['Italic', 'italic']]) {
      if (this.index(name) !== this.loaded[name]) patch[key] = FONT_MODES[this.index(name)];
    }
    for (const [edge] of EDGES) {
      const names = ['Mode_', 'Color_', 'Width_', 'Dash_'].map(p => p + edge);
      const values = [this.index(names[0]), this.text(names[1]), this.text(names[2]), this.index(names[3])];
      if (JSON.stringify(values) === this.loaded[edge]) continue;
      attempt(() => {
        const mode = BORDER_MODES[values[0]]; if (mode === 'unsupported') return;
        const color = mode === 'solid' ? colorValue(values[1]) : '#FFFFFF', width = Number(values[2]);
        if (mode === 'solid' && (!Number.isFinite(width) || width < BORDER_MIN || width > BORDER_MAX || ['unsupported', 'inherit'].includes(color))) throw new Error('边框需为 0.25–20 磅，颜色需为十六进制值或主题颜色。');
        patch.borders ||= {}; patch.borders[edge] = { mode, color, width, dash: DASHES[values[3]] };
      });
    }
    const base = JSON.parse(JSON.stringify(this.loaded.patch || {}));
    this.patches[this.part] = { ...base, ...patch, ...(patch.borders ? { borders: { ...base.borders, ...patch.borders } } : {}) };
    // Empty patches would unnecessarily issue a new style ID on every save.
    if (!Object.keys(this.patches[this.part]).length) delete this.patches[this.part];
  }
  saveDraft() {
    this.capture();
    if (this.selected) this.drafts.set(this.selected.id, { patches: this.patches, name: this.text('StyleName') });
  }
  selectStyle(index) {
    this.selected = this.allStyles()[index] || null;
    this.updating = true;
    try {
      if (!this.selected) { this.enabled(); return; }
      const draft = this.drafts.get(this.selected.id); this.patches = draft?.patches || {};
      this.c('StyleName').Text = draft?.name || this.selected.name;
      this.c('SetDefault').Value = this.selected.id.toUpperCase() === this.context.defaultID.toUpperCase();
      const f = this.context.tables.find(t => t.styleID.toUpperCase() === this.selected.id.toUpperCase())?.flags || { firstRow: true, bandRow: true };
      for (const key of FLAGS) this.c('Flag_' + key).Value = !!f[key];
      this.fields(); this.preview(); this.enabled();
    } finally { this.updating = false; }
  }
  populate(id) {
    const all = this.allStyles();
    const found = all.findIndex(s => s.id.toUpperCase() === String(id).toUpperCase());
    const index = found < 0 ? 0 : found;
    this.updating = true;
    try { this.list('Styles', all.map(s => s.name + (s.id.toUpperCase() === this.context.defaultID.toUpperCase() ? ' · 默认' : '')), index); }
    finally { this.updating = false; }
    this.selectStyle(index);
  }
  async refresh(preferred) {
    this.context = await this.host.inspect(); this.drafts.clear(); this.newStyles = [];
    if (!this.context) { this.selected = null; this.enabled(); return; }
    this.c('SetDefaultFont').Value = false;
    this.c('DefaultFontSize').Text = String(this.context.fontDefaults.points || 18);
    this.populate(preferred || this.context.defaultID || this.context.styles[0]?.id);
  }
  fields() {
    if (!this.selected?.xml) return;
    const p = readPart(this.editedXML(), this.part);
    this.c('Fill').Text = colorText(p.fill); this.c('TextColor').Text = colorText(p.color);
    this.c('FillChoice').Text = colorText(p.fill); this.c('TextColorChoice').Text = colorText(p.color);
    this.select('Bold', FONT_MODES.indexOf(p.bold)); this.select('Italic', FONT_MODES.indexOf(p.italic));
    this.loaded = { Fill: this.text('Fill'), TextColor: this.text('TextColor'), Bold: this.index('Bold'), Italic: this.index('Italic') };
    this.loaded.patch = JSON.parse(JSON.stringify(this.patches[this.part] || {}));
    for (const [edge] of EDGES) {
      const b = p.borders[edge];
      this.select('Mode_' + edge, BORDER_MODES.indexOf(b.mode));
      this.c('Color_' + edge).Text = b.color && b.color !== 'unsupported' ? colorText(b.color) : '#FFFFFF';
      this.c('Width_' + edge).Text = String(b.width || 1);
      this.select('Dash_' + edge, DASHES.indexOf(b.dash || 'solid'));
      this.loaded[edge] = JSON.stringify([this.index('Mode_' + edge), this.text('Color_' + edge), this.text('Width_' + edge), this.index('Dash_' + edge)]);
    }
  }
  preview() {
    const style = this.editedXML(), f = this.flags();
    const parts = style ? Object.fromEntries(PARTS.map(([key]) => [key, readPart(style, key)])) : {};
    for (let ri = 0; ri < 4; ri++) for (let ci = 0; ci < 4; ci++) {
      const rules = ['wholeTbl'];
      if (f.bandRow) rules.push((ri - (f.firstRow ? 1 : 0)) % 2 === 0 ? 'band1H' : 'band2H');
      if (f.bandCol) rules.push((ci - (f.firstCol ? 1 : 0)) % 2 === 0 ? 'band1V' : 'band2V');
      if (f.lastCol && ci === 3) rules.push('lastCol');
      if (f.firstCol && ci === 0) rules.push('firstCol');
      if (f.lastRow && ri === 3) { rules.push('lastRow'); if (f.lastCol && ci === 3) rules.push('seCell'); if (f.firstCol && ci === 0) rules.push('swCell'); }
      if (f.firstRow && ri === 0) { rules.push('firstRow'); if (f.lastCol && ci === 3) rules.push('neCell'); if (f.firstCol && ci === 0) rules.push('nwCell'); }
      const result = { fill: '#FFFFFF', color: '#20242A', bold: 'off', italic: 'off' };
      for (const key of rules) for (const k of Object.keys(result)) if (parts[key] && !['inherit', 'unsupported'].includes(parts[key][k])) result[k] = parts[key][k];
      const cell = this.c('Preview_' + ri + '_' + ci);
      cell.BackColor = rgb(colorForPreview(result.fill, THEMES));
      cell.ForeColor = rgb(colorForPreview(result.color, THEMES, '#20242A'));
      try { cell.Font.Bold = result.bold === 'on'; cell.Font.Italic = result.italic === 'on'; } catch { /* Older JSA fonts have fewer properties. */ }
    }
  }
  add(duplicate = false) {
    this.saveDraft();
    const choices = ['green', 'red', 'blue', 'blank'], labels = ['我的绿色', '我的红色', '我的蓝色', '我的空白'];
    const base = duplicate ? this.editedXML() : preset(choices[this.index('Preset')]);
    if (!base) throw new Error('请先选择一个可编辑的自定义样式。');
    const id = guid(), name = duplicate ? this.text('StyleName') + ' 副本' : labels[this.index('Preset')];
    this.newStyles.push({ id, name, xml: editStyle(base, name, id, {}), create: true, builtin: false });
    this.populate(id); this.c('SetDefault').Value = true; this.select('Scope', 0);
  }
  async save() {
    this.capture();
    const s = this.selected; if (!s) throw new Error('请先打开 PPTX 并选择或新建样式。');
    const plan = { path: this.context.path, id: s.id, name: this.text('StyleName'), sourceXML: s.xml || null,
      baseXML: s.create ? s.xml : undefined, create: !!s.create, patches: this.patches,
      flags: this.flags(), scope: SCOPES[this.index('Scope')], setDefault: this.checked('SetDefault'), clearOverrides: this.checked('ClearOverrides') };
    if (this.checked('SetDefaultFont')) plan.defaultFontSize = Number(this.text('DefaultFontSize'));
    const result = await this.host.apply(plan);
    await this.refresh(result.id);
  }
  scale(delta) {
    this.zoom = Math.max(50, Math.min(200, this.zoom + delta));
    const ratio = this.zoom / 100;
    for (const item of this.geometry) {
      if (item.toolbar) continue;
      const c = this.c(item.name);
      c.Left = item.x * ratio; c.Top = 48 + (item.y - 48) * ratio;
      c.Width = item.w * ratio; c.Height = item.h * ratio;
      try { c.Font.Size = 10 * ratio; } catch { /* Geometry remains scalable. */ }
    }
    this.c('ZoomLevel').Caption = this.zoom + '%';
    this.form.Width = Math.max(700, 700 * ratio); this.form.Height = 48 + 714 * ratio;
  }
  async dispatch(action) {
    if (this.updating || this.busy) return;
    // Disabling a native edit on every Change commits the edit and steals its
    // focus. Local changes are synchronous and must keep their controls alive.
    const local = ['fields', 'fillChoice', 'textColorChoice', 'flags', 'font', 'zoomIn', 'zoomOut'].includes(action);
    this.status(''); this.busy = true;
    if (!local) this.enabled();
    try {
      if (action === 'open') { if (await this.host.openFile()) await this.refresh(); }
      else if (action === 'current') { this.host.path = null; await this.refresh(); }
      else if (action === 'save') await this.save();
      else if (action === 'new') this.add();
      else if (action === 'duplicate') this.add(true);
      else if (action === 'style') { this.saveDraft(); this.selectStyle(this.index('Styles')); }
      else if (action === 'part') {
        this.capture(); this.part = PARTS[this.index('Part')][0];
        this.updating = true; try { this.fields(); } finally { this.updating = false; }
      }
      else if (action === 'fields') { this.capture(false); this.preview(); }
      else if (action === 'fillChoice' || action === 'textColorChoice') {
        const name = action === 'fillChoice' ? 'Fill' : 'TextColor';
        this.c(name).Text = this.text(name + 'Choice'); this.capture(false); this.preview();
      }
      else if (action === 'flags') this.preview();
      else if (action === 'allBorders') {
        this.capture(); const top = readPart(this.editedXML(), this.part).borders.top;
        if (top.mode === 'unsupported') throw new Error('请先设置“上”边框，再应用到全部边框。');
        this.patches[this.part] ||= {}; this.patches[this.part].borders = Object.fromEntries(EDGES.map(([e]) => [e, { ...top }]));
        this.updating = true; try { this.fields(); } finally { this.updating = false; }
        this.preview();
      }
      else if (action === 'zoomIn') this.scale(10);
      else if (action === 'zoomOut') this.scale(-10);
      else if (action === 'close') this.form.Hide();
    } catch (e) {
      this.updating = true;
      try {
        if (action === 'style') this.select('Styles', this.allStyles().indexOf(this.selected));
        if (action === 'part') this.select('Part', PARTS.findIndex(x => x[0] === this.part));
      } finally { this.updating = false; }
      this.status(e.message || String(e));
    }
    finally { this.busy = false; if (!local || action === 'font') this.enabled(); }
  }
}

export function createEditor(app, form, geometry) {
  const editor = new NativeEditor(installRuntime(app), form, geometry);
  editor.initialize();
  return editor;
}
