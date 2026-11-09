import { WpsHost, REQUEST, RESPONSE, LAYOUT } from './host.js';
import { screenBounds, floatingSize } from './window-fit.js';
import { maximizePaneDimension, settlePane } from './pane-size.js';

let host, running = false, lastRequest, pane, resizeTimer, floatingTarget, layoutGeneration = 0;
let maximized;
const baseURL = () => new URL('.', window.location.href).href;
function app() { return window.Application || window.wps; }

function dockSpaceKey(currentPane, layout, bounds) {
  let width = 0, height = 0;
  try { width = Number(app().ActiveWindow.Width) || 0; height = Number(app().ActiveWindow.Height) || 0; } catch { /* Optional native window geometry. */ }
  return JSON.stringify([bounds.availWidth, bounds.availHeight, Number(currentPane[layout === 'workspace' ? 'Width' : 'Height']), width, height]);
}

function resizePane(layout = 'workspace', metrics = {}) {
  // Migrate the previous persisted name to the sole right-docked layout.
  if (layout === 'sidebar-max') layout = 'sidebar';
  if (!['workspace', 'floating', 'sidebar'].includes(layout)) throw new Error('未知的编辑器布局。');
  try { if (pane) void pane.Visible; }
  catch { pane = null; maximized = null; }
  const bounds = screenBounds(metrics, window.screen);
  const dock = layout === 'workspace' ? 1 : layout === 'floating' ? 4 : 2;
  const alreadyDocked = !!pane && !!pane.Visible && Number(pane.DockPosition) === dock;
  if (layout !== 'floating' && alreadyDocked && maximized?.pane === pane && maximized.layout === layout
    && maximized.space === dockSpaceKey(pane, layout, bounds)
    && Math.abs(Number(pane[layout === 'workspace' ? 'Height' : 'Width']) - maximized.value) <= 1) {
    // A repeated expand click must not hide, redock, shrink or reload the editor.
    return { layout, width: pane.Width, height: pane.Height, changed: false };
  }
  const generation = ++layoutGeneration;
  clearTimeout(resizeTimer);
  maximized = null;
  // Set the mode before making the pane visible: its page may load immediately.
  app().PluginStorage?.setItem(LAYOUT, layout);
  if (!pane) {
    pane = app().CreateTaskPane(baseURL() + 'editor.html', '表格样式编辑器');
  }
  // Reuse one pane in every layout, retaining its DOM and unsaved edits.
  floatingTarget = floatingSize(bounds);
  // Show first because WPS initializes native geometry when Visible changes.
  if (!alreadyDocked) {
    pane.Visible = true;
    pane.DockPosition = dock;
    // Docking can hide/recreate the native panel. Show it again before sizing;
    // never treat that transient hidden state as a request to cancel opening.
    pane.Visible = true;
  }
  const currentPane = pane;
  const cancelled = () => {
    if (generation !== layoutGeneration || pane !== currentPane) return true;
    try { return !currentPane.Visible; } catch { return true; }
  };
  if (layout === 'workspace' || layout === 'sidebar') {
    const property = layout === 'workspace' ? 'Height' : 'Width';
    // Leave a usable editor synchronously, even if WPS delays background
    // timers for the invisible ribbon page. Expansion is a separate phase.
    if (!alreadyDocked) {
      for (const initial of layout === 'workspace' ? [360, 240, 120] : [460, 320, 240]) {
        try { currentPane[property] = initial; } catch { continue; }
        if (Number(currentPane[property]) >= initial - 1) break;
      }
      currentPane.Visible = true;
    }
    // Width is owned by WPS for top docking, and Height for right docking.
    // Wait for native docking before searching its adjustable dimension.
    return (async () => {
      await settlePane(220);
      if (cancelled()) return { layout, cancelled: true };
      const result = await maximizePaneDimension(currentPane, property, layout === 'workspace' ? bounds.availHeight : bounds.availWidth, { cancelled });
      if (!result.cancelled && !cancelled()) maximized = { pane: currentPane, layout, value: result.value, space: dockSpaceKey(currentPane, layout, bounds) };
      return { layout, width: currentPane.Width, height: currentPane.Height, cancelled: result.cancelled };
    })();
  }
  const size = () => {
    currentPane.Width = floatingTarget.width;
    if (layout === 'floating') currentPane.Height = floatingTarget.height;
  };
  size();
  // Apply once more after WPS completes its asynchronous show/dock layout.
  // Cancel on a new mode request so an old expansion cannot resize the sidebar.
  resizeTimer = setTimeout(() => {
    resizeTimer = undefined;
    if (cancelled()) return;
    try { if (currentPane.Visible) size(); } catch { /* Closed native pane. */ }
  }, 250);
  return { layout, width: pane.Width, height: pane.Height };
}

function fitFloating(metrics = {}) {
  if (!pane || app().PluginStorage?.getItem(LAYOUT) !== 'floating' || !pane.Visible || Number(pane.DockPosition) !== 4) return { changed: false };
  const bounds = screenBounds(metrics, window.screen);
  floatingTarget = floatingSize(bounds, { width: Number(pane.Width), height: Number(pane.Height) }, metrics);
  const changed = Math.abs(Number(pane.Width) - floatingTarget.width) > 2 || Math.abs(Number(pane.Height) - floatingTarget.height) > 2;
  if (changed) { pane.Width = floatingTarget.width; pane.Height = floatingTarget.height; }
  return { changed, width: pane.Width, height: pane.Height };
}

window.OnTableStylesLoad = function () {
  host = new WpsHost(app());
  // Dispatch in the application-level page, not the document-level task pane.
  setInterval(async () => {
    if (running) return;
    const store = app().PluginStorage;
    const raw = store.getItem(REQUEST);
    if (!raw) return;
    let request;
    try { request = JSON.parse(raw); } catch { return; }
    if (request.id === lastRequest) return;
    lastRequest = request.id; running = true;
    store.setItem(REQUEST, '');
    try {
      if (!['inspect', 'apply', 'resize', 'fit'].includes(request.method)) throw new Error('未知的编辑器操作。');
      const value = request.method === 'resize'
        ? await resizePane(request.params.layout, request.params.metrics)
        : request.method === 'fit' ? fitFloating(request.params.metrics) : await host[request.method](request.params);
      store.setItem(RESPONSE, JSON.stringify({ id: request.id, value }));
    } catch (e) { store.setItem(RESPONSE, JSON.stringify({ id: request.id, error: e.message || String(e) })); }
    finally { running = false; }
  }, 250);
  return true;
};
window.OpenTableStyles = function () {
  const report = e => alert('无法打开表格样式编辑器：' + (e.message || String(e)));
  // WPS ribbon callbacks are synchronous. Catch native opening errors here;
  // only the optional expansion continues after the callback has returned.
  try { Promise.resolve(resizePane('workspace')).catch(report); }
  catch (e) { report(e); }
  return true;
};
// Ribbon images use resource-relative paths for WPS offline loading.
window.GetTableStylesImage = () => 'table-style.svg';
