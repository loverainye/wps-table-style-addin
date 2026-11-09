export const FORM_ZOOM = 'table-style-editor.form-zoom.v1';

// Keep the native toolbar at its normal size while scaling only the form.
export function setupFormZoom({ form, level, slider, out, plus, storage }) {
  let current = 100;
  function set(value, persist = true) {
    const parsed = Number(value);
    current = Number.isFinite(parsed) && parsed > 0
      ? Math.max(50, Math.min(200, Math.round(parsed / 10) * 10)) : 100;
    const scale = current / 100;
    form.style.zoom = String(scale);
    // Limit the rendered width at any scale, without making low zoom narrow.
    form.style.maxWidth = (1120 / scale) + 'px';
    level.value = slider.value = String(current);
    slider.setAttribute('aria-valuetext', current + '%');
    level.title = `缩放比例：${current}%；选择 100% 恢复原大小`;
    out.disabled = current <= 50;
    plus.disabled = current >= 200;
    out.title = `缩小表单（当前 ${current}%）`;
    plus.title = `放大表单（当前 ${current}%）`;
    if (persist) {
      try { storage()?.setItem(FORM_ZOOM, String(current)); } catch { /* Session-only if storage is unavailable. */ }
    }
  }
  for (let value = 50; value <= 200; value += 10) {
    const option = level.ownerDocument.createElement('option');
    option.value = String(value); option.textContent = value + '%'; level.append(option);
  }
  out.addEventListener('click', () => set(current - 10));
  plus.addEventListener('click', () => set(current + 10));
  level.addEventListener('change', () => set(level.value));
  slider.addEventListener('input', () => set(slider.value));
  let saved;
  try { saved = storage()?.getItem(FORM_ZOOM); } catch { /* Default to 100%. */ }
  set(saved ?? 100, false);
}
