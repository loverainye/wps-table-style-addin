import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, access, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { buildAddon } from '../scripts/build.mjs';

for (const hideOnDock of [false, true]) test(`opens synchronously and maximizes docked dimensions (dock hides pane: ${hideOnDock})`, async () => {
  const out = await mkdtemp(path.join(tmpdir(), 'wps-offline-bootstrap-'));
  try {
    await buildAddon(out);
    const entry = await readFile(new URL('fixtures/wps-linux-index.html', import.meta.url), 'utf8');
    const code = await readFile(path.join(out, entry.match(/src='([^']+)'/)[1]), 'utf8');
    await assert.rejects(access(path.join(out, 'index.html')));
    const ribbon = await readFile(path.join(out, 'ribbon.xml'), 'utf8');
    const action = ribbon.match(/onAction="([^"]+)"/)[1];
    const onLoad = ribbon.match(/onLoad="([^"]+)"/)[1];
    const created = [], alerts = [], initialLayouts = [], writes = [];
    const nativeWindow = { Width: 1440, Height: 900 };
    // Native geometry initializes on showing/docking. Oversized requests are
    // ignored; WPS owns top width and right height rather than accepting writes.
    let visible = false, width = 320, height = 68, dock = 2, maxWidth = 934;
    const pane = {
      get Visible() { return visible; },
      set Visible(value) { writes.push(['Visible', value]); if (value && !visible) { width = 320; height = 68; } visible = value; },
      get Width() { return dock === 1 ? 1440 : width; },
      set Width(value) {
        writes.push(['Width', value]);
        assert.notEqual(dock, 1, 'top width must remain controlled by WPS');
        if (visible && (dock === 4 || value <= maxWidth)) width = value;
      },
      get Height() { return dock === 2 ? 720 : height; },
      set Height(value) {
        writes.push(['Height', value]);
        assert.notEqual(dock, 2, 'right height must remain controlled by WPS');
        if (visible && (dock === 4 || value <= 618)) height = value;
      },
      get DockPosition() { return dock; },
      set DockPosition(value) {
        writes.push(['DockPosition', value]);
        dock = value;
        if (visible) { width = 320; height = 68; }
        if (hideOnDock) visible = false;
      },
    };
    const storage = new Map(), timers = new Map();
    let nextTimer = 0, dispatch;
    const flushTimers = async () => {
      for (let step = 0; step < 80; step++) {
        const pending = [...timers.values()]; timers.clear();
        for (const callback of pending) callback();
        await Promise.resolve(); await Promise.resolve();
      }
    };
    const context = {
      URL, setInterval: callback => { dispatch = callback; return 1; },
      setTimeout: callback => { const id = ++nextTimer; timers.set(id, callback); return id; },
      clearTimeout: id => timers.delete(id),
      screen: { availWidth: 1440, availHeight: 900 },
      location: { href: 'file:///home/test/.local/share/Kingsoft/wps/jsaddons/WpsPresentationTableStyles_1.8.0/index.html' },
      Application: {
        ActiveWindow: nativeWindow,
        CreateTaskPane: url => { created.push(url); initialLayouts.push(storage.get('table-style-editor.layout.v1')); return pane; },
        PluginStorage: { getItem: key => storage.get(key), setItem: (key, value) => storage.set(key, value) },
      },
      alert: message => alerts.push(message),
    };
    context.window = context;
    vm.runInNewContext(code, context);
    assert.equal(context[onLoad](), true);
    assert.equal(context[action](), true);
    assert.equal(pane.Visible, true, 'the ribbon action must show the pane after docking');
    assert.ok(pane.Height >= 220, 'the editor must be usable before asynchronous maximization');
    await flushTimers();
    assert.deepEqual(created, ['file:///home/test/.local/share/Kingsoft/wps/jsaddons/WpsPresentationTableStyles_1.8.0/editor.html']);
    assert.deepEqual(initialLayouts, ['workspace']);
    assert.equal(pane.Visible, true);
    assert.equal(pane.DockPosition, 1);
    assert.equal(pane.Width, 1440);
    assert.equal(pane.Height, 618);
    await access(path.join(out, context.GetTableStylesImage()));
    assert.equal(context.OpenTableStylesSidebar, undefined);
    assert.deepEqual(alerts, []);
    const request = async (id, method, params) => {
      storage.set('table-style-editor.request.v1', JSON.stringify({ id, method, params }));
      const work = dispatch(); await flushTimers(); await work;
      return JSON.parse(storage.get('table-style-editor.response.v1'));
    };
    writes.length = 0;
    const topRepeat = await request('top-repeat', 'resize', { layout: 'workspace' });
    assert.equal(topRepeat.value.changed, false);
    assert.deepEqual(writes, [], 'already-maximized top must not touch native geometry');
    await request('sidebar', 'resize', { layout: 'sidebar' });
    assert.equal(pane.DockPosition, 2); assert.equal(pane.Width, 934); assert.equal(pane.Height, 720);
    writes.length = 0;
    const sideRepeat = await request('side-repeat', 'resize', { layout: 'sidebar' });
    assert.equal(sideRepeat.value.changed, false);
    assert.deepEqual(writes, [], 'already-maximized sidebar must not touch native geometry');
    // A manual drag to a smaller width should allow expansion without redocking.
    width = 600;
    await request('side-manual-restore', 'resize', { layout: 'sidebar' });
    assert.equal(pane.Width, 934);
    assert.ok(writes.length > 0);
    assert.ok(writes.every(([property]) => property === 'Width'));
    // Recalculate after the main WPS window gains more available space.
    maxWidth = 1040;
    nativeWindow.Width = 1600;
    writes.length = 0;
    await request('sidebar-refit', 'resize', { layout: 'sidebar' });
    assert.equal(pane.Width, 1040);
    assert.ok(writes.every(([property]) => property === 'Width'), 'window resize must not redock');
    await request('floating', 'resize', { layout: 'floating' });
    assert.equal(pane.DockPosition, 4); assert.equal(pane.Width, 1408); assert.equal(pane.Height, 836);
    await request('fit-monitor', 'fit', { metrics: { availWidth: 1920, availHeight: 1080, innerWidth: pane.Width - 16, innerHeight: pane.Height - 48, outerWidth: pane.Width, outerHeight: pane.Height } });
    assert.ok(pane.Width > 1800); assert.ok(pane.Height > 1000);
    await request('top-again', 'resize', { layout: 'workspace' });
    assert.equal(pane.DockPosition, 1); assert.equal(pane.Width, 1440); assert.equal(pane.Height, 618);
    const invalid = await request('invalid', 'resize', { layout: 'invalid' });
    assert.match(invalid.error, /布局/);
    assert.equal(pane.DockPosition, 1);
    assert.equal(storage.get('table-style-editor.layout.v1'), 'workspace');
    const ignored = await request('stale-fit', 'fit', { metrics: { availWidth: 1920, availHeight: 1080 } });
    assert.equal(ignored.value.changed, false); assert.equal(pane.Height, 618);
    pane.Visible = false;
    context[action](); await flushTimers();
    assert.equal(pane.DockPosition, 1); assert.equal(pane.Height, 618);
    assert.equal(created.length, 1);
  } finally { await rm(out, { recursive: true, force: true }); }
});
