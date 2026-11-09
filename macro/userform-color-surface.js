import { hsvToHex } from './color-picker.js';

// Every display element belongs to TableStyleColorForm. The existing HSV
// controller supplies exact pointer colors; these labels draw its background.
export class UserFormColorSurface {
  constructor(form, geometry, onClose) {
    this.nativeForm = form;
    this.plane = geometry.plane; this.hue = geometry.hue;
    this.columns = geometry.columns; this.rows = geometry.rows;
    this.onClose = onClose; this.shown = false; this.syncing = false;
    this.paintedHue = 0; // The embedded resource already contains a red plane.
    this.cells = null;
    for (const name of ['Original', 'Current', 'Hex', 'R', 'G', 'B', 'ErrorText']) {
      Object.defineProperty(this, name, { get: () => form[name] });
    }
  }
  get visible() { return this.shown && this.nativeForm.Visible !== false && this.nativeForm.Visible !== 0; }
  pointerPosition(event) {
    const control = this.nativeForm[event.surface === 'plane' ? 'Plane' : 'Hue'];
    const width = Number(control.Width), height = Number(control.Height);
    if (!(width > 0 && height > 0)) throw new Error('无法读取 UserForm 取色区域尺寸。');
    return { x: Number(event.x)/width, y: Number(event.y)/height };
  }
  render(hsv, color, final = true) {
    if (!final) return;
    const hue = Math.round(hsv[0]*10)/10;
    if (this.paintedHue === hue) return;
    // Hue dragging updates the chosen color immediately; repaint this plane
    // only on release or Enter, avoiding hundreds of native calls per move.
    this.cells ||= Array.from({ length: this.rows }, (_, row) =>
      Array.from({ length: this.columns }, (_, col) => this.nativeForm['PlaneCell_'+row+'_'+col]));
    this.syncing = true;
    try {
      for (let row = 0; row < this.rows; row++) for (let col = 0; col < this.columns; col++) {
        this.cells[row][col].BackColor = parseInt(hsvToHex(hue, col/(this.columns-1), 1-row/(this.rows-1)).slice(1), 16);
      }
      this.paintedHue = hue;
    } finally { this.syncing = false; }
  }
  Show() {
    this.shown = true;
    try {
      // A modal native Show returns when Apply, Cancel, or the window X closes
      // it. Keep main-form state intact through that native event loop.
      this.nativeForm.Show();
      if (this.nativeForm.Visible === false || this.nativeForm.Visible === 0) this.shown = false;
    } catch (error) { this.shown = false; throw error; }
  }
  Hide() {
    this.nativeForm.Hide(); this.shown = false;
    this.onClose?.();
  }
}
