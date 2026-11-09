import test from 'node:test';
import assert from 'node:assert/strict';
import { maximizePaneDimension } from '../src/pane-size.js';

for (const mode of ['ignore', 'clamp', 'throw']) test('finds maximum accepted top height when excessive requests ' + mode, async () => {
  let actual = 68, queued;
  const pane = {
    get Height() { return actual; },
    set Height(value) {
      if (value > 617 && mode === 'throw') throw new Error('native limit');
      queued = value <= 617 ? value : mode === 'clamp' ? 617 : actual;
    },
  };
  const result = await maximizePaneDimension(pane, 'Height', 900, { wait: async () => { if (queued !== undefined) actual = queued; } });
  assert.equal(result.value, 617);
  assert.equal(pane.Height, 617);
});

test('cancels an obsolete search without restoring dimensions over a new layout', async () => {
  let cancelled = false, width = 320, writes = 0;
  const pane = { get Width() { return width; }, set Width(value) { writes++; width = Math.min(value, 800); } };
  const result = await maximizePaneDimension(pane, 'Width', 1440, { cancelled: () => cancelled, wait: async () => { cancelled = true; width = 460; } });
  assert.equal(result.cancelled, true);
  assert.equal(writes, 1);
  assert.equal(width, 460);
});
