import { NativeEditor, colorValue, previewColor } from './runtime.js';
import { EDGES, BORDER_MIN, BORDER_MAX, BORDER_STEP, readPart } from '../src/styles.js';
import { SWATCHES } from './color-picker.js';

const colorFields = ['Fill', 'TextColor', ...EDGES.map(([edge]) => 'Color_' + edge)];
const colorLabel = value => ({ inherit: '继承', none: '无填充', unsupported: '保留现有' }[value] || value);

// The two interfaces keep separate drafts but share the same ZIP/XML runtime.
export class UserFormEditor extends NativeEditor {
  constructor(app, form, geometry, pickerFactory) {
    super(app, form, geometry);
    this.explicitBorderModes = new Set();
    this.pickerFactory = pickerFactory;
  }
  c(name) { return super.c(({ FillChoice: 'Fill', TextColorChoice: 'TextColor' })[name] || name); }
  list(name, labels, index = 0) { return super.list(({ FillChoice: 'Fill', TextColorChoice: 'TextColor' })[name] || name, labels, index); }
  initialize() {
    super.initialize();
    this.updating = true;
    try {
      const colors = [...SWATCHES, ...this.choices.TextColor.filter(v => v.startsWith('theme:'))];
      this.list('Fill', ['继承', '无填充', ...colors, '保留现有']);
      this.list('TextColor', ['继承', ...colors, '保留现有']);
      for (const [edge] of EDGES) this.list('Color_' + edge, colors, colors.indexOf('#FFFFFF'));
      this.updateColorButtons();
    } finally { this.updating = false; }
  }
  enabled() {
    const busy = this.busy;
    this.busy = busy || !!this.picker?.form.visible;
    try { super.enabled(); }
    finally { this.busy = busy; }
    for (const field of [...EDGES.map(([e]) => 'Width_' + e), 'DefaultFontSize']) {
      for (const direction of ['Up', 'Down']) this.c('Step' + direction + '_' + field).Enabled = this.c(field).Enabled;
    }
    this.c('CloseEditor').Enabled = !this.busy && !this.picker?.form.visible;
    for (const field of colorFields) this.c('Pick' + field).Enabled = this.c(field).Enabled;
  }
  addColorOption(name, value) {
    const options = this.choices[name] ||= [];
    if (!options.includes(value)) { this.c(name).AddItem(value); options.push(value); }
  }
  fields() {
    this.explicitBorderModes.clear();
    if (this.selected?.xml) {
      const part = readPart(this.editedXML(), this.part);
      this.addColorOption('Fill', colorLabel(part.fill));
      this.addColorOption('TextColor', colorLabel(part.color));
      for (const [edge] of EDGES) {
        const color = part.borders[edge].color;
        if (color && color !== 'unsupported') this.addColorOption('Color_' + edge, colorLabel(color));
      }
    }
    super.fields(); this.updateColorButtons();
  }
  preview() { super.preview(); this.updateColorButtons(); }
  updateColorButtons() {
    for (const field of colorFields) {
      let color;
      try { color = colorValue(this.text(field), field === 'Fill'); } catch { color = '#FFFFFF'; }
      const button = this.c('Pick' + field);
      button.BackColor = parseInt(previewColor(color).slice(1), 16);
      button.Caption = '\u00a0';
    }
  }
  setColor(field, color) {
    if (!colorFields.includes(field)) throw new Error('未知的颜色区域。');
    const hex = colorValue(color);
    this.updating = true;
    try {
      this.addColorOption(field, hex); this.c(field).Text = hex;
      if (field.startsWith('Color_')) {
        this.explicitBorderModes.delete(field.slice(6)); this.select('Mode_' + field.slice(6), 1);
      }
    } finally { this.updating = false; }
    this.capture(false); this.preview();
  }
  async openColor(field) {
    if (!colorFields.includes(field)) throw new Error('未知的颜色区域。');
    if (!this.selected?.xml || this.selected.builtin) return;
    if (!this.pickerFactory) throw new Error('缺少文件内的色盘资源。');
    this.pickerOpening = true;
    try {
      this.capture(false);
      this.picker ||= this.pickerFactory();
      let value;
      try { value = colorValue(this.text(field), field === 'Fill'); } catch { value = '#C00000'; }
      await this.picker.open(previewColor(value), hex => this.setColor(field, hex));
    } finally { this.pickerOpening = false; }
  }
  pickerAction(action, value, page) {
    if (page !== undefined && !this.picker?.form.acceptEvent(page)) return;
    return this.picker?.dispatch(action, value);
  }
  setBorderLine(edge) {
    const previous = this.updating; this.updating = true;
    try { this.select('Mode_' + edge, 1); }
    finally { this.updating = previous; }
  }
  capture(strict = true) {
    if (!this.updating && this.selected?.xml) for (const [edge] of EDGES) {
      const loaded = this.loaded[edge] && JSON.parse(this.loaded[edge]);
      if (!loaded || this.explicitBorderModes.has(edge)) continue;
      const mode = this.index('Mode_' + edge);
      const changed = this.text('Color_' + edge) !== loaded[1] || this.text('Width_' + edge) !== loaded[2] || this.index('Dash_' + edge) !== loaded[3];
      if (changed && mode === loaded[0] && mode !== 1) this.setBorderLine(edge);
    }
    return super.capture(strict);
  }
  stepNumber(field, direction) {
    const width = EDGES.some(([e]) => field === 'Width_' + e);
    if ((!width && field !== 'DefaultFontSize') || !this.c(field).Enabled) return;
    const value = Number(this.text(field));
    if (!Number.isFinite(value) || !this.text(field).trim()) throw new Error('请先输入有效的数值。');
    const next = Math.max(width ? BORDER_MIN : 1, Math.min(width ? BORDER_MAX : 4000,
      Math.round((value + Number(direction) * (width ? BORDER_STEP : 1)) * 100) / 100));
    this.updating = true;
    try {
      this.c(field).Text = String(next);
      if (width) { this.explicitBorderModes.delete(field.slice(6)); this.setBorderLine(field.slice(6)); }
    } finally { this.updating = false; }
    this.capture(false);
  }
  async dispatch(action) {
    if (this.busy || this.updating || this.picker?.form.visible || this.picker?.form.syncing || ['zoomIn', 'zoomOut', 'fit'].includes(action)) return;
    if (action.startsWith('pick:')) {
      try { this.status(''); this.busy = true; this.enabled(); await this.openColor(action.slice(5)); }
      catch (error) { this.status(error.message || String(error)); }
      finally { this.busy = false; this.enabled(); }
      return;
    }
    if (action.startsWith('step:')) {
      try { this.status(''); const [, field, direction] = action.split(':'); this.stepNumber(field, direction); }
      catch (error) { this.status(error.message || String(error)); }
      return;
    }
    if (action.startsWith('mode:')) this.explicitBorderModes.add(action.slice(5));
    if (action.startsWith('border:')) {
      const edge = action.slice(7); this.explicitBorderModes.delete(edge); this.setBorderLine(edge);
    }
    return super.dispatch(action.startsWith('border:') || action.startsWith('mode:') ? 'fields' : action);
  }
}
