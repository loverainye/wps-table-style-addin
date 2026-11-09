import test from 'node:test';
import assert from 'node:assert/strict';
import { screenBounds, measureWindow, floatingSize } from '../src/window-fit.js';

test('uses available logical pixels without multiplying by HiDPI scale', () => {
  const report = measureWindow({ screen: { availWidth: 1440, availHeight: 860, width: 1440, height: 900 }, devicePixelRatio: 2, innerWidth: 800, innerHeight: 600, outerWidth: 816, outerHeight: 648 });
  assert.deepEqual(screenBounds(report), { availWidth: 1440, availHeight: 860 });
  assert.deepEqual(floatingSize(screenBounds(report)), { width: 1408, height: 796 });
});

test('fits decoration-inclusive window bounds across native pixel and point units', () => {
  for (const units of [0.75, 1, 2]) {
    let native = { width: 900 * units, height: 600 * units };
    for (let round = 0; round < 3; round++) {
      native = floatingSize({ availWidth: 1600, availHeight: 900 }, native, { innerWidth: native.width / units - 16, innerHeight: native.height / units - 48, outerWidth: native.width / units, outerHeight: native.height / units });
    }
    assert.ok(Math.abs(native.width / units - 1584) <= 4, 'outer width fits the available screen');
    assert.ok(Math.abs(native.height / units - 884) <= 4, 'outer height leaves taskbar and borders visible');
  }
});

test('falls back for unavailable metrics and ignores implausible embedded outer geometry', () => {
  const bounds = screenBounds({ availWidth: NaN, availHeight: -1 }, { availWidth: 1024, availHeight: 768 });
  assert.deepEqual(bounds, { availWidth: 1024, availHeight: 768 });
  assert.deepEqual(floatingSize(bounds, {}, { innerWidth: 0, innerHeight: 0 }), { width: 992, height: 704 });
  assert.deepEqual(floatingSize(bounds, {}, { innerWidth: 300, outerWidth: 8000 }), { width: 992, height: 704 });
});

test('does not resize a pane already within the fitting tolerance', () => {
  const result = floatingSize({ availWidth: 1440, availHeight: 900 }, { width: 1424, height: 884 }, { innerWidth: 1408, innerHeight: 836, outerWidth: 1424, outerHeight: 884 });
  assert.deepEqual(result, { width: 1424, height: 884 });
});
