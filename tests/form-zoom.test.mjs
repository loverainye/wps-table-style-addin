import test from 'node:test';
import assert from 'node:assert/strict';
import { setupFormZoom, FORM_ZOOM } from '../src/form-zoom.js';

function controls(saved, unavailable = false) {
  const values = new Map(saved === undefined ? [] : [[FORM_ZOOM, saved]]);
  const element = () => ({
    value: '', style: {}, options: [], attributes: {}, listeners: {},
    ownerDocument: { createElement: () => ({}) },
    append(option) { this.options.push(option); },
    setAttribute(key, value) { this.attributes[key] = value; },
    addEventListener(type, work) { this.listeners[type] = work; },
    fire(type, value) { if (value !== undefined) this.value = value; this.listeners[type](); },
  });
  const ui = { form: element(), level: element(), slider: element(), out: element(), plus: element() };
  setupFormZoom({ ...ui, storage: () => {
    if (unavailable) throw new Error('no storage');
    return { getItem: key => values.get(key), setItem: (key, value) => values.set(key, value) };
  } });
  return { ...ui, values };
}

test('zoom slider, selector and step buttons stay synchronized and restore 100%', () => {
  const ui = controls('80');
  assert.equal(ui.level.options.length, 16);
  assert.equal(ui.slider.value, '80');
  ui.plus.fire('click');
  assert.equal(ui.level.value, '90');
  assert.equal(ui.form.style.zoom, '0.9');
  ui.slider.fire('input', '140');
  assert.equal(ui.level.value, '140');
  assert.equal(ui.slider.attributes['aria-valuetext'], '140%');
  assert.equal(ui.values.get(FORM_ZOOM), '140');
  ui.level.fire('change', '100');
  assert.equal(ui.form.style.zoom, '1');
  assert.equal(ui.slider.value, '100');
  assert.equal(ui.form.style.maxWidth, '1120px');
});

test('zoom keeps rendered width bounded and disables the matching boundary button', () => {
  const ui = controls('50');
  assert.equal(ui.out.disabled, true);
  assert.equal(ui.plus.disabled, false);
  assert.equal(ui.form.style.maxWidth, '2240px');
  ui.slider.fire('input', '200');
  assert.equal(ui.out.disabled, false);
  assert.equal(ui.plus.disabled, true);
  assert.equal(ui.form.style.maxWidth, '560px');
  ui.plus.fire('click');
  assert.equal(ui.level.value, '200');
  ui.out.fire('click');
  assert.equal(ui.level.value, '190');
});

test('empty, invalid and unavailable storage starts at 100% and does not block zoom', () => {
  for (const value of ['', 'invalid', '0', '-20', undefined]) assert.equal(controls(value).level.value, '100');
  const ui = controls(undefined, true);
  ui.slider.fire('input', '70');
  assert.equal(ui.level.value, '70');
  assert.equal(ui.form.style.zoom, '0.7');
});
