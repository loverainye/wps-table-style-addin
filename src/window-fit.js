const positive = value => typeof value === 'number' && Number.isFinite(value) && value > 0;
export function screenBounds(value = {}, fallback = {}) {
  const dimension = (key, defaultValue) => Math.floor(positive(value[key]) ? value[key] : positive(fallback[key]) ? fallback[key] : defaultValue);
  return { availWidth: dimension('availWidth', 1280), availHeight: dimension('availHeight', 900) };
}
export function measureWindow(view) {
  return {
    ...screenBounds(view.screen),
    innerWidth: view.innerWidth, innerHeight: view.innerHeight,
    outerWidth: view.outerWidth, outerHeight: view.outerHeight,
  };
}
// Screen/viewport dimensions are CSS pixels. Calibrate against the visible pane
// instead of multiplying by devicePixelRatio (which would oversize HiDPI panes).
export function floatingSize(bounds, native = {}, viewport = {}) {
  function axis(available, current, inner, outer, fallbackFrame) {
    const measuredFrame = positive(inner) && positive(outer) ? outer - inner : 0;
    const frame = measuredFrame > 0 && measuredFrame <= 160 ? measuredFrame : fallbackFrame;
    const goal = Math.max(1, available - frame - 16);
    if (!positive(current) || !positive(inner) || inner < 100) return Math.floor(goal);
    if (Math.abs(inner - goal) <= 4) return current;
    const ratio = current / inner;
    if (ratio < 0.4 || ratio > 3) return Math.floor(goal);
    return Math.max(1, Math.min(Math.floor(available * 3), Math.round(current * goal / inner)));
  }
  return {
    width: axis(bounds.availWidth, native.width, viewport.innerWidth, viewport.outerWidth, 16),
    height: axis(bounds.availHeight, native.height, viewport.innerHeight, viewport.outerHeight, 48),
  };
}
