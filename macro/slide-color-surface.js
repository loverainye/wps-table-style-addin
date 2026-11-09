import { binaryString } from '../src/host.js';
import { hsvToHex } from './color-picker.js';
import { guid } from '../src/uuid.js';

export const PLANE = { x: 296, y: 150, w: 560, h: 280 };
// The swatches occupy their own row; the hue strip is a separate full-width row.
export const HUE = { x: 296, y: 500, w: 560, h: 24 };

// Standard 24-bit BMP. The macro needs no browser, extra executable or image library.
export function gradientBitmap(width, height, hue, strip = false) {
  const stride = (width * 3 + 3) & ~3;
  const data = new Uint8Array(54 + stride * height), view = new DataView(data.buffer);
  data[0] = 66; data[1] = 77;
  view.setUint32(2, data.length, true); view.setUint32(10, 54, true);
  view.setUint32(14, 40, true); view.setInt32(18, width, true); view.setInt32(22, height, true);
  view.setUint16(26, 1, true); view.setUint16(28, 24, true);
  view.setUint32(34, stride * height, true);
  const base = [1, 3, 5].map(i => parseInt(hsvToHex(hue, 1, 1).slice(i, i + 2), 16));
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const s = x / (width - 1), v = 1 - y / (height - 1);
    const rgb = strip ? [1, 3, 5].map(i => parseInt(hsvToHex((1 - s) * 360, 1, 1).slice(i, i + 2), 16))
      : base.map(c => Math.round(v * (255 + (c - 255) * s)));
    const at = 54 + (height - 1 - y) * stride + x * 3;
    data[at] = rgb[2]; data[at + 1] = rgb[1]; data[at + 2] = rgb[0];
  }
  return data;
}

// Editor and picker controls live on separate slides. Alternate two prebuilt
// picker slides to redraw ordinary shapes through a real page change. Never
// toggle native Visible or reload the page hosting the active mouse control.
export class SlideColorSurface {
  constructor(app, presentation, getControls, reportError) {
    this.app = app; this.presentation = presentation; this.getControls = getControls;
    this.reportError = reportError; this.plane = PLANE; this.hue = HUE;
    this.activePage = 2; this.displayPage = null; this.visible = false;
    this.syncing = false; this.refreshTask = null; this.lastRefreshError = '';
    this.paintedHue = new Map(); this.prepared = new Set(); this.controlGroups = new Map();
    for (const name of Object.keys(this.controlsFor(2))) {
      Object.defineProperty(this, name, { get: () => this.controlsFor(this.activePage)[name] });
    }
  }
  controlsFor(page) {
    if (!this.controlGroups.has(page)) this.controlGroups.set(page, this.getControls(page));
    return this.controlGroups.get(page);
  }
  slide(page = this.activePage) { return this.presentation.Slides.Item(page); }
  shape(name, page = this.activePage) { return this.slide(page).Shapes.Item(name); }
  acceptEvent(page) { return this.visible && !this.syncing && Number(page) === this.displayPage; }
  showView() {
    let view;
    try { view = this.presentation.SlideShowWindow?.View; } catch { /* No current show. */ }
    if (!view) view = this.presentation.SlideShowSettings.Run()?.View;
    if (!view || typeof view.GotoSlide !== 'function') throw new Error('当前 WPS 无法切换本工具的放映画面。');
    return view;
  }
  snapshot() {
    const group = this.controlsFor(this.activePage);
    const state = {};
    for (const name of ['Hex', 'R', 'G', 'B']) state[name] = { Text: group[name].Text };
    for (const name of ['Original', 'Current']) state[name] = {
      Caption: '\u00a0', BackColor: group[name].BackColor, ForeColor: group[name].ForeColor,
    };
    state.ErrorText = { Caption: group.ErrorText.Caption };
    return state;
  }
  refresh() {
    if (this.refreshTask) return this.refreshTask;
    this.refreshTask = Promise.resolve().then(() => {
      const previous = this.activePage;
      this.syncing = true;
      try {
        const view = this.showView();
        let current;
        try { current = Number(view.CurrentShowPosition); } catch { /* Older hosts. */ }
        if (![1, 2, 3].includes(current)) current = this.displayPage || 1;
        if (!this.visible) {
          if (current !== 1) view.GotoSlide(1, -1);
          this.displayPage = null;
        } else {
          const target = current === 2 ? 3 : 2;
          const state = this.snapshot(), controls = this.controlsFor(target);
          // Update only the page that is not currently on screen. Native Change
          // callbacks emitted by these assignments are ignored while syncing.
          for (const [name, values] of Object.entries(state)) {
            for (const [key, value] of Object.entries(values)) controls[name][key] = value;
          }
          this.paintPage(target);
          this.activePage = target;
          view.GotoSlide(target, -1);
          this.displayPage = target;
        }
        this.lastRefreshError = '';
        return true;
      } catch (error) {
        this.activePage = previous;
        this.visible = this.displayPage !== null;
        throw error;
      } finally { this.syncing = false; }
    }).catch(error => {
      const message = '无法切换色盘页面：' + (error.message || String(error));
      this.ErrorText.Caption = message; this.lastRefreshError = message;
      try { if (this.reportError) this.reportError(new Error(message)); }
      catch { /* The inline error remains if the host cannot show a message. */ }
      return false;
    }).then(result => { this.refreshTask = null; return result; });
    return this.refreshTask;
  }
  native(step, work) {
    try { return work(); }
    catch (error) { throw new Error(step + '：' + (error.message || String(error))); }
  }
  paint(name, bytes, page) {
    return this.native('更新渐变图片失败', () => {
      const dir = String(this.app.Env.GetTempPath()).replace(/[\\/]?$/, '/');
      const path = dir + 'wps-table-color-' + guid().slice(1, -1) + '.bmp';
      const fs = this.app.FileSystem;
      try {
        fs.writeAsBinaryString(path, binaryString(bytes));
        this.shape(name, page).Fill.UserPicture(path);
      } finally { try { if (fs.Exists(path)) fs.Remove(path); } catch { /* Best effort temp cleanup. */ } }
    });
  }
  paintPage(page) {
    const hue = Math.round(this.hsv[0] * 10) / 10;
    if (!this.prepared.has(page)) {
      this.paint('TsColorHueImage', gradientBitmap(this.hue.w, 2, 0, true), page);
      this.prepared.add(page);
    }
    if (this.paintedHue.get(page) !== hue) {
      this.paint('TsColorPlaneImage', gradientBitmap(280, 140, hue), page);
      this.paintedHue.set(page, hue);
    }
  }
  pointerPosition(event) {
    const name = event.surface === 'plane' ? 'Plane' : 'Hue';
    const area = this.controlsFor(this.activePage)[name];
    const w = Number(area.Width), h = Number(area.Height);
    const x = Number(event.x) / w, y = Number(event.y) / h;
    this.lastPointer = { control: name, phase: event.phase, button: event.button, rawX: event.x, rawY: event.y, width: w, height: h, x, y };
    if (!(w > 0 && h > 0)) throw new Error('无法读取取色区域尺寸。');
    return { x, y };
  }
  render(hsv, color, final = true) {
    this.hsv = [...hsv];
    if (final && this.visible) return this.refresh();
  }
  renderError() { if (this.visible) return this.refresh(); }
  Show() { this.visible = true; return this.refresh(); }
  Hide() { this.visible = false; return this.refresh(); }
}
