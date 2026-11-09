import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { NativeColorPicker, hsvToHex, hexToHSV, SWATCHES } from '../macro/color-picker.js';

const items = JSON.parse(execFileSync('python3', ['-c', 'import sys,json;sys.path.insert(0,"macro");from color_picker import definitions;print(json.dumps(definitions()))'],
  { cwd: new URL('..', import.meta.url), encoding: 'utf8' }));
function mock() {
  const form = { plane: { w: 560, h: 280 }, hue: { w: 560, h: 24 }, shows: 0, hides: 0, Show() { this.shows++; }, Hide() { this.hides++; } };
  for (const item of items) form[item.name] = { Text: '', Caption: item.caption };
  const picker = new NativeColorPicker(form), applied = [];
  picker.open('#C00000', color => applied.push(color));
  return { form, picker, applied };
}

test('color conversion roundtrips all palette colors and primary HSV endpoints', () => {
  assert.equal(hsvToHex(0, 1, 1), '#FF0000');
  assert.equal(hsvToHex(120, 1, 1), '#00FF00');
  assert.equal(hsvToHex(240, 1, 1), '#0000FF');
  for (const hex of [...SWATCHES, '#1264A8', '#AF2789', '#010101']) assert.equal(hsvToHex(...hexToHSV(hex)), hex);
});

test('continuous plane and hue dragging synchronize exact HEX, RGB and preview before applying', () => {
  const { picker: p, form: f, applied } = mock();
  const pointer = (surface, phase, x, y, button = 1) => p.dispatch('pointer', { surface, phase, button, x, y });
  pointer('hue', 'down', 280, 12); pointer('hue', 'up', 280, 12, 0);
  pointer('plane', 'down', 560, 0); pointer('plane', 'up', 560, 0, 0);
  assert.equal(f.Hex.Text, '#00FFFF');
  assert.equal(f.R.Text, '0'); assert.equal(f.G.Text, '255'); assert.equal(f.B.Text, '255');
  assert.equal(f.Current.BackColor, 0x00FFFF); assert.equal(f.Original.BackColor, 0xC00000);
  pointer('plane', 'down', 100.25, 46.75); pointer('plane', 'move', 237.125, 121.25);
  assert.equal(p.hsv[1], 237.125 / 560); assert.equal(p.hsv[2], 1 - 121.25 / 280);
  pointer('plane', 'up', 237.125, 121.25, 0);
  const color = p.color;
  pointer('plane', 'move', 560, 280, 0); assert.equal(p.color, color, 'hover must not change selection');
  assert.deepEqual(applied, []); p.dispatch('apply'); assert.deepEqual(applied, [color]);
  assert.equal(f.hides, 1);
});

test('dragging clamps coordinates and retains hue while white or black is selected', () => {
  const { picker: p, form: f } = mock();
  p.dispatch('pointer', { surface: 'plane', phase: 'down', button: 1, x: -20, y: 400 });
  assert.equal(p.color, '#000000');
  p.dispatch('pointer', { surface: 'hue', phase: 'down', button: 1, x: 560 / 3, y: 12 });
  assert.equal(p.hsv[0], 240.00000000000003);
  p.dispatch('pointer', { surface: 'plane', phase: 'down', button: 1, x: 560, y: 0 });
  assert.equal(f.Hex.Text, '#0000FF');
});

test('partial HEX and invalid RGB cannot apply; exact RGB applies and cancellation does not call back', () => {
  const { picker: p, form: f, applied } = mock();
  f.Hex.Text = '#12'; p.dispatch('hex'); p.dispatch('apply');
  assert.match(f.ErrorText.Caption, /六位/); assert.deepEqual(applied, []);
  f.Hex.Text = '#1264a8'; p.dispatch('hex'); assert.equal(f.Current.BackColor, 0x1264A8);
  f.R.Text = '256'; p.dispatch('rgb'); p.dispatch('apply');
  assert.match(f.ErrorText.Caption, /0–255/); assert.deepEqual(applied, []);
  f.R.Text = '18'; f.G.Text = '52'; f.B.Text = '86'; p.dispatch('rgb'); p.dispatch('apply');
  assert.deepEqual(applied, ['#123456']);
  p.open('#C00000', color => applied.push(color)); p.dispatch('swatch', 4); p.dispatch('cancel');
  assert.deepEqual(applied, ['#123456']);
});
