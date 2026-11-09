import { NativeEditor, installRuntime, colorValue, previewColor } from './runtime.js';
import { FileHost } from '../src/file-host.js';
import { openDeck, EDGES, BORDER_MIN, BORDER_MAX, BORDER_STEP, readPart } from '../src/styles.js';
import { NativeColorPicker, contrastColor, SWATCHES } from './color-picker.js';

export { installRuntime, NativeColorPicker };
export { SlideColorSurface } from './slide-color-surface.js';
export { UserFormEditor } from './userform-runtime.js';
export { UserFormColorSurface } from './userform-color-surface.js';

const colorFields = ['Fill', 'TextColor', ...EDGES.map(([edge]) => 'Color_' + edge)];
const colorLabel = value => ({ inherit: '继承', none: '无填充', unsupported: '保留现有' }[value] || value);

// The tool stays active on its own slide. Edit closed external files so that
// saving WPS's in-memory presentation cannot overwrite the updated ZIP parts.
class SlideFileHost extends FileHost {
  requireClosed(path) {
    if (this.opened(path)) throw new Error('请先关闭目标 PPTX，再从本页“打开 PPTX”加载文件。');
  }
  async openFile() {
    const picker = this.app.FileDialog(3);
    picker.Title = '打开需要编辑表格样式的 PPTX';
    picker.AllowMultiSelect = false;
    picker.Filters.Clear(); picker.Filters.Add('PowerPoint 演示文稿', '*.pptx');
    if (Number(picker.Show()) !== -1) return false;
    return this.openPath(String(picker.SelectedItems.Item(1)));
  }
  async openPath(path) {
    if (!/\.pptx$/i.test(path)) throw new Error('请选择 .pptx 文件，宏工具本身保留为 .pptm。');
    this.requireClosed(path);
    const deck = await openDeck(this.read(path));
    this.path = path;
    this.openInspection = deck;
    return true;
  }
  async inspect() {
    if (!this.path) return null;
    this.requireClosed(this.path);
    if (this.openInspection) {
      const deck = this.openInspection; this.openInspection = null;
      return this.describe(deck, this.path, this.path.split(/[\\/]/).pop(), []);
    }
    return super.inspect();
  }
  async apply(plan) {
    this.requireClosed(plan.path);
    return super.apply(plan);
  }
}

export class SlideEditor extends NativeEditor {
  constructor(app, controls, pickerFactory) {
    // UseCurrent is deliberately absent from the slide, but the shared form
    // controller keeps its Enabled property in sync.
    const form = { UseCurrent: { Enabled: false } };
    for (const name of Object.keys(controls)) Object.defineProperty(form, name, { get: () => controls[name] });
    Object.defineProperty(form, 'FillChoice', { get: () => controls.Fill });
    Object.defineProperty(form, 'TextColorChoice', { get: () => controls.TextColor });
    super(app, form);
    this.host = new SlideFileHost(app);
    this.presentation = app.ActivePresentation;
    this.pickerFactory = pickerFactory;
    this.colorCatalog = {}; this.loadedColorLists = new Set();
    this.explicitBorderModes = new Set();
  }
  initialize() {
    this.deferColors = true;
    super.initialize();
    this.updating = true;
    try {
      this.list('Scope', ['所有表格', '使用此样式的表格', '不更改现有表格']);
      const colors = [...SWATCHES, ...this.colorCatalog.TextColor.filter(v => v.startsWith('theme:'))];
      this.list('Fill', ['继承', '无填充', ...colors, '保留现有']);
      this.list('TextColor', ['继承', ...colors, '保留现有']);
      for (const [edge] of EDGES) this.list('Color_' + edge, colors);
    } finally { this.updating = false; this.deferColors = false; }
  }
  list(name, labels, index = 0) {
    const field = ({ FillChoice: 'Fill', TextColorChoice: 'TextColor' })[name] || name;
    if (this.deferColors && colorFields.includes(field)) {
      this.colorCatalog[field] = labels;
      return super.list(name, labels.length ? [labels[index] || labels[0]] : []);
    }
    return super.list(name, labels, index);
  }
  loadColors(field) {
    if (!colorFields.includes(field) || this.loadedColorLists.has(field)) return;
    this.updating = true;
    try {
      const current = this.text(field), options = [...new Set([...(this.colorCatalog[field] || []), current])];
      super.list(field, options, options.indexOf(current));
      this.loadedColorLists.add(field);
    } finally { this.updating = false; }
  }
  enabled() {
    super.enabled();
    for (const field of colorFields) this.c('Pick' + field).Enabled = !!this.selected?.xml && !this.selected.builtin && !this.busy;
    for (const field of [...EDGES.map(([e]) => 'Width_' + e), 'DefaultFontSize']) {
      for (const direction of ['Up', 'Down']) this.c('Step' + direction + '_' + field).Enabled = this.c(field).Enabled;
    }
  }
  addColorOption(name, value) {
    const choices = this.choices[name] ||= [];
    if (!choices.includes(value)) { this.c(name).AddItem(value); choices.push(value); }
  }
  fields() {
    this.explicitBorderModes.clear();
    if (this.selected?.xml) {
      const p = readPart(this.editedXML(), this.part);
      this.addColorOption('Fill', colorLabel(p.fill));
      this.addColorOption('TextColor', colorLabel(p.color));
      for (const [edge] of EDGES) {
        const color = p.borders[edge].color;
        if (color && color !== 'unsupported') this.addColorOption('Color_' + edge, colorLabel(color));
      }
    }
    super.fields(); this.updateColorButtons();
  }
  preview() { super.preview(); this.updateColorButtons(); }
  updateColorButtons() {
    for (const field of colorFields) {
      let value = this.text(field);
      try { value = colorValue(value, field === 'Fill'); } catch { value = '#FFFFFF'; }
      const hex = previewColor(value), button = this.c('Pick' + field);
      button.BackColor = parseInt(hex.slice(1), 16);
      button.Caption = '\u00a0';
      button.ForeColor = contrastColor(hex);
    }
  }
  capture(strict = true) {
    // A final native edit may commit when Save receives focus, before its
    // Change handler runs. Detect that case as well as normal border events.
    if (!this.updating && this.selected?.xml) for (const [edge] of EDGES) {
      const loaded = this.loaded[edge] && JSON.parse(this.loaded[edge]);
      if (!loaded || this.explicitBorderModes.has(edge)) continue;
      const mode = this.index('Mode_' + edge);
      const changed = this.text('Color_' + edge) !== loaded[1] || this.text('Width_' + edge) !== loaded[2] || this.index('Dash_' + edge) !== loaded[3];
      if (changed && mode !== 1 && mode === loaded[0]) this.setBorderLine(edge);
    }
    return super.capture(strict);
  }
  setBorderLine(edge) {
    const previous = this.updating; this.updating = true;
    try { this.select('Mode_' + edge, 1); }
    finally { this.updating = previous; }
  }
  setColor(field, color) {
    if (!colorFields.includes(field)) throw new Error('未知的颜色区域。');
    const hex = colorValue(color);
    this.updating = true;
    try {
      this.addColorOption(field, hex); this.c(field).Text = hex;
      if (field.startsWith('Color_')) {
        const edge = field.slice(6);
        this.explicitBorderModes.delete(edge); this.select('Mode_' + edge, 1);
      }
    } finally { this.updating = false; }
    this.capture(false); this.preview();
  }
  async openColor(field) {
    if (!this.selected?.xml || this.selected.builtin) return;
    if (!this.pickerFactory) throw new Error('缺少文件内的原生色盘资源。');
    this.picker ||= this.pickerFactory();
    let value;
    try { value = colorValue(this.text(field), field === 'Fill'); } catch { value = '#C00000'; }
    await this.picker.open(previewColor(value), hex => this.setColor(field, hex));
  }
  pickerAction(action, value, page) {
    if (page !== undefined && !this.picker?.form.acceptEvent(page)) return;
    return this.picker?.dispatch(action, value);
  }
  stepNumber(field, direction) {
    const width = EDGES.some(([e]) => field === 'Width_' + e);
    if ((!width && field !== 'DefaultFontSize') || !this.c(field).Enabled) return;
    const value = Number(this.text(field));
    if (!Number.isFinite(value) || !this.text(field).trim()) throw new Error('请先输入有效的数值。');
    const next = Math.max(width ? BORDER_MIN : 1, Math.min(width ? BORDER_MAX : 4000, Math.round((value + Number(direction) * (width ? BORDER_STEP : 1)) * 100) / 100));
    this.updating = true;
    try { this.c(field).Text = String(next); if (width) this.setBorderLine(field.slice(6)); }
    finally { this.updating = false; }
    this.capture(false);
  }
  async openSelected(path) {
    if (this.busy || this.updating) return;
    this.busy = true; this.enabled();
    this.c('FirstPlaceholder').Caption = '正在读取文稿…';
    try { if (await this.host.openPath(path)) await this.refresh(); }
    catch (error) { this.status(error.message || String(error)); }
    finally { this.busy = false; this.enabled(); }
  }
  acceptEvent(page) {
    if (Number(page) !== 1 || this.picker?.form.visible || this.picker?.form.syncing) return false;
    try {
      const current = Number(this.presentation.SlideShowWindow?.View.CurrentShowPosition);
      return !Number.isFinite(current) || current === 1;
    } catch { return true; }
  }
  async dispatch(action) {
    if (['current', 'close', 'zoomIn', 'zoomOut', 'fit'].includes(action)) return;
    if (this.busy || this.updating) return;
    if (action.startsWith('colors:')) return this.loadColors(action.slice(7));
    if (action.startsWith('step:')) {
      try { this.status(''); const [, field, direction] = action.split(':'); this.stepNumber(field, direction); }
      catch (error) { this.status(error.message || String(error)); }
      return;
    }
    if (action.startsWith('mode:')) this.explicitBorderModes.add(action.slice(5));
    if (action.startsWith('border:')) {
      const edge = action.slice(7); this.explicitBorderModes.delete(edge); this.setBorderLine(edge);
    }
    if (action.startsWith('border:') || action.startsWith('mode:')) return super.dispatch('fields');
    if (!action.startsWith('pick:')) return super.dispatch(action);
    try {
      this.status(''); this.busy = true;
      await this.openColor(action.slice(5));
    } catch (error) { this.status(error.message || String(error)); }
    finally { this.busy = false; }
  }
}
