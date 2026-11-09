export const SWATCHES = ['#FFFFFF', '#000000', '#E7E6E6', '#7F7F7F', '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050',
  '#00B0F0', '#0070C0', '#002060', '#7030A0', '#385723', '#A9D18E', '#F4B183', '#F8CBAD', '#B4C6E7', '#D9E1F2'];

export function hsvToHex(h, s, v) {
  const f = n => {
    const k = (n + h / 60) % 6;
    return Math.round(255 * (v - v * s * Math.max(0, Math.min(k, 4 - k, 1)))).toString(16).padStart(2, '0');
  };
  return ('#' + f(5) + f(3) + f(1)).toUpperCase();
}

export function hexToHSV(hex) {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255);
  const v = Math.max(r, g, b), min = Math.min(r, g, b), d = v - min;
  const h = !d ? 0 : v === r ? ((g - b) / d + 6) % 6 : v === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h * 60, v ? d / v : 0, v];
}

export const contrastColor = hex => {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16));
  return r * 0.299 + g * 0.587 + b * 0.114 > 155 ? 0x20242A : 0xFFFFFF;
};

export class NativeColorPicker {
  constructor(form) { this.form = form; this.updating = false; this.hsv = [0, 1, 1]; this.color = '#C00000'; this.dragging = null; }
  c(name) { return this.form[name]; }
  open(initial, callback) {
    this.callback = callback; this.source = 'hex'; this.dragging = null;
    this.c('Original').Caption = '\u00a0';
    this.c('Original').BackColor = parseInt(initial.slice(1), 16);
    this.c('ErrorText').Caption = '\u00a0';
    this.setColor(initial);
    return this.form.Show();
  }
  setColor(hex, source, exactHSV, final = true) {
    this.updating = true;
    try {
      this.color = hex.toUpperCase();
      const hsv = hexToHSV(this.color);
      // Hue is meaningful for the picker even when selecting black/white.
      this.hsv = exactHSV || [hsv[1] && hsv[2] ? hsv[0] : this.hsv[0], hsv[1], hsv[2]];
      if (source !== 'hex') this.c('Hex').Text = this.color;
      for (const [i, channel] of ['R', 'G', 'B'].entries()) {
        if (source !== 'rgb') this.c(channel).Text = String(parseInt(this.color.slice(1 + i * 2, 3 + i * 2), 16));
      }
      this.c('Current').BackColor = parseInt(this.color.slice(1), 16);
      this.c('Current').ForeColor = contrastColor(this.color);
      this.c('Current').Caption = '\u00a0';
      this.finalRender = final && source !== 'hex' && source !== 'rgb';
      return this.renderGrid();
    } finally { this.updating = false; }
  }
  renderGrid() {
    if (this.form.render) return this.form.render(this.hsv, this.color, this.finalRender);
  }
  pointer(event) {
    const { surface, phase, button, x, y } = event;
    if (!['plane', 'hue'].includes(surface) || !['down', 'move', 'up'].includes(phase)) return;
    if (phase === 'down') {
      if (!(Number(button) & 1)) return;
      this.dragging = surface;
    } else if (this.dragging !== surface) return;
    // Qt-based WPS can report Button=0 for MouseMove while the left button is
    // held. Keep the drag started by MouseDown until MouseUp arrives.
    const geometry = surface === 'plane' ? this.form.plane : this.form.hue;
    const position = this.form.pointerPosition?.(event) || { x: Number(x) / geometry.w, y: Number(y) / geometry.h };
    const nx = Math.max(0, Math.min(1, position.x));
    const ny = Math.max(0, Math.min(1, position.y));
    if (!Number.isFinite(nx) || !Number.isFinite(ny)) return;
    if (surface === 'plane') { this.hsv[1] = nx; this.hsv[2] = 1 - ny; }
    else this.hsv[0] = (1 - nx) * 360;
    const selected = [...this.hsv];
    this.source = 'hex';
    const result = this.setColor(hsvToHex(...selected), undefined, selected, phase === 'up');
    if (phase === 'up') this.dragging = null;
    return result;
  }
  input(source, strict = false) {
    if (source === 'rgb') {
      const rgb = ['R', 'G', 'B'].map(key => String(this.c(key).Text).trim());
      if (!rgb.every(s => /^\d{1,3}$/.test(s) && Number(s) <= 255)) {
        if (strict) throw new Error('RGB 请输入 0–255 的整数。');
        return;
      }
      return '#' + rgb.map(s => Number(s).toString(16).padStart(2, '0')).join('').toUpperCase();
    }
    const hex = String(this.c('Hex').Text).trim();
    if (!/^#?[\da-f]{6}$/i.test(hex)) {
      if (strict) throw new Error('请输入六位十六进制颜色，例如 #C00000。');
      return;
    }
    return '#' + hex.replace(/^#/, '').toUpperCase();
  }
  dispatch(action, value) {
    if (this.updating) return;
    try {
      this.c('ErrorText').Caption = '\u00a0';
      if (action === 'pointer') return this.pointer(value);
      if (action === 'cancel') { this.dragging = null; return this.form.Hide(); }
      if (action === 'apply') {
        const color = this.input(this.source, true);
        this.callback(color); this.dragging = null; return this.form.Hide();
      }
      if (action === 'commit') return this.setColor(this.input(this.source, true));
      if (action === 'hex' || action === 'rgb') {
        this.source = action;
        const hex = this.input(action); if (hex) return this.setColor(hex, action);
        return;
      }
      this.source = 'hex';
      if (action === 'swatch') this.setColor(SWATCHES[value]);
    } catch (error) {
      this.c('ErrorText').Caption = error.message || String(error);
      return this.form.renderError?.();
    }
  }
}
