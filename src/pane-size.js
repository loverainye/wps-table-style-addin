export const settlePane = ms => new Promise(resolve => setTimeout(resolve, ms));

// Some WPS versions reject an oversized docked dimension and leave its old
// value unchanged. Probe and read back the accepted value instead of assuming
// that a request equal to the screen size will be clamped by the host.
export async function maximizePaneDimension(pane, property, upper, { wait = settlePane, cancelled = () => false } = {}) {
  if (!['Width', 'Height'].includes(property)) throw new Error('未知的窗格尺寸。');
  if (!Number.isFinite(upper) || upper < 1) throw new Error('无效的窗格尺寸上限。');
  let low = 1, high = Math.ceil(upper), best = Math.max(1, Number(pane[property]) || 1), candidate = high;
  for (let step = 0; step < 16 && low <= high; step++) {
    if (cancelled()) return { cancelled: true };
    try { pane[property] = candidate; } catch { /* Oversized native dimensions can throw. */ }
    await wait(45);
    if (cancelled()) return { cancelled: true };
    const actual = Number(pane[property]);
    if (!Number.isFinite(actual) || actual < 1) throw new Error('WPS 未返回有效的窗格尺寸。');
    best = Math.max(best, actual);
    if (actual >= candidate - 1) low = candidate + 1;
    else high = candidate - 1;
    low = Math.max(low, Math.ceil(best) + 1);
    candidate = Math.floor((low + high) / 2);
  }
  if (cancelled()) return { cancelled: true };
  if (Number(pane[property]) < best) { pane[property] = best; await wait(45); }
  return { value: Number(pane[property]), cancelled: false };
}
