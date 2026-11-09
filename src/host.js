import { openDeck, applyStyle, guid } from './styles.js';

const bytes = value => {
  if (typeof value !== 'string') throw new Error('WPS 返回了无法识别的文件数据。');
  return Uint8Array.from(value, c => c.charCodeAt(0) & 255);
};
export function binaryString(data) {
  const parts = [];
  for (let i = 0; i < data.length; i += 16384) parts.push(String.fromCharCode(...data.subarray(i, i + 16384)));
  return parts.join('');
}
function identical(a, b) {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}
function contentSignature(deck) {
  return Object.entries(deck.zip.files)
    .filter(([name, entry]) => !entry.dir && !['docProps/core.xml', 'docProps/app.xml'].includes(name))
    .map(([name, entry]) => `${name}:${entry._data?.crc32}:${entry._data?.uncompressedSize}`)
    .sort().join('\n');
}
export class WpsHost {
  constructor(app) { this.app = app; this.busy = false; this.pending = new Map(); }
  current() {
    const p = this.app.ActivePresentation;
    if (!p) throw new Error('请先打开需要修改的演示稿。');
    const fullName = String(p.FullName || '');
    // FullName carries the saved local path. Do not also require the redundant
    // Path property when FullName is already a valid local absolute filename.
    const local = /^(?:\/|[A-Za-z]:[\\/]|\\\\)/.test(fullName) && !/[\u0000\r\n]/.test(fullName);
    if (!local || !/\.pptx$/i.test(fullName)) {
      const name = String(p.Name || fullName || '未命名');
      throw new Error('当前文件“' + name + '”未被识别为本地 .pptx。请保存为本地 .pptx 后再打开编辑器。'
        + '（WPS 返回 FullName=' + (fullName || '空') + '；Path=' + String(p.Path || '空') + '）');
    }
    if (p.ReadOnly) throw new Error('当前演示稿为只读，请先取得编辑权限。');
    const fs = this.app.FileSystem;
    if (!fs || typeof fs.readAsBinaryString !== 'function' || typeof fs.writeAsBinaryString !== 'function') throw new Error('此 WPS 版本缺少加载项二进制文件接口，请升级带 JS 加载项支持的桌面版。');
    return p;
  }
  selection() {
    try {
      const win = this.app.ActiveWindow, slideID = String(win.View.Slide.SlideID), range = win.Selection.ShapeRange;
      const targets = [];
      const visit = shape => {
        if (shape.HasTable) targets.push({ slideID, shapeID: String(shape.Id ?? shape.ID) });
        else if (Number(shape.Type) === 6) for (let i = 1; i <= shape.GroupItems.Count; i++) visit(shape.GroupItems.Item(i));
      };
      for (let i = 1; i <= range.Count; i++) visit(range.Item(i));
      return targets;
    } catch { return []; }
  }
  read(path) { return bytes(this.app.FileSystem.readAsBinaryString(path)); }
  write(path, data) {
    this.app.FileSystem.writeAsBinaryString(path, binaryString(data));
    if (!identical(this.read(path), data)) throw new Error('文件写入校验失败：' + path);
  }
  remove(path) { try { if (this.app.FileSystem.Exists(path)) this.app.FileSystem.Remove(path); } catch { /* The recovery file may remain. */ } }
  async snapshot() {
    const p = this.current();
    const path = String(p.FullName), name = String(p.Name), targets = this.selection();
    const tmp = String(this.app.Env.GetTempPath()).replace(/[\\/]?$/, '/') + 'wps-table-style-' + guid().slice(1, -1) + '.pptx';
    let data;
    try {
      // SaveCopyAs includes unsaved edits without changing the open file or its name.
      p.SaveCopyAs(tmp, 24);
      data = this.read(tmp);
    } finally { this.remove(tmp); }
    const deck = await openDeck(data);
    let slideIndex = 1;
    try { slideIndex = Number(this.app.ActiveWindow.View.Slide.SlideIndex); } catch { /* Slide sorter or no active slide. */ }
    return { deck, path, name, targets, slideIndex };
  }
  describe(deck, path, name, targets = this.selection()) {
    return { path, name, targets, defaultID: deck.defaultID, styles: deck.styles, tables: deck.tables, tableCount: deck.tables.length, fontDefaults: deck.fontDefaults };
  }
  async inspect() {
    if (this.busy) throw new Error('正在保存样式，请稍候。');
    const p = this.current(), path = String(p.FullName);
    // The open document intentionally keeps its old in-memory styles. Inspect
    // the written file so the editor retains the latest definitions and GUIDs.
    if (this.pending.has(path)) return this.describe(await openDeck(this.read(path)), path, String(p.Name));
    const { deck, path: namePath, name, targets } = await this.snapshot();
    return this.describe(deck, namePath, name, targets);
  }
  async apply(plan) {
    if (this.busy) throw new Error('正在保存样式，请稍候。');
    this.busy = true;
    let path, backup, staging, original, writeStarted = false;
    try {
      const p = this.current();
      path = String(p.FullName);
      if (path !== plan.path) throw new Error('当前文稿已经切换，请重新选择需要修改的文件。');
      const snap = await this.snapshot();
      const memorySignature = contentSignature(snap.deck);
      const pending = this.pending.get(path);
      let deck = snap.deck, reuseDisk = false;
      if (pending) {
        const disk = await openDeck(this.read(path));
        const diskSignature = contentSignature(disk);
        if (memorySignature === diskSignature) {
          // A manual reopen loaded the written styles; normal saves are safe.
          this.pending.delete(path);
        } else if (p.Saved && memorySignature === pending.memorySignature && diskSignature === pending.diskSignature) {
          // Repeated editor saves use the latest XML without saving stale native
          // content over it. No document close, open or slide navigation occurs.
          deck = disk; reuseDisk = true;
        } else {
          throw new Error('文件已写入新样式，当前 WPS 文稿仍保留旧内容且已发生变化。请先将需要保留的正文修改另存，再手动重新打开目标文件。');
        }
      }
      const existing = deck.styles.find(s => s.id.toUpperCase() === String(plan.id).toUpperCase());
      if (!plan.create && (existing?.xml || null) !== (plan.sourceXML || null)) throw new Error('文稿中的样式已被其他操作更改，请重新读取文件后重试。');
      const result = await applyStyle(deck, { ...plan, targets: snap.targets });
      const verified = await openDeck(result.bytes);
      if (!verified.styles.some(s => s.id === result.id)) throw new Error('生成的表格样式校验失败。');
      if (plan.defaultFontSize !== undefined && (verified.fontDefaults.mixed || verified.fontDefaults.points !== plan.defaultFontSize)) throw new Error('生成的母版默认字号校验失败。');
      if (String(this.current().FullName) !== path) throw new Error('当前文稿已经切换，本次操作已停止。');
      if (!reuseDisk) p.Save();
      if (!p.Saved) throw new Error('当前文稿未能保存，尚未修改样式。');
      original = this.read(path);
      if (contentSignature(await openDeck(original)) !== contentSignature(deck)) throw new Error('保存期间文稿内容发生变化，已保留你的最新修改，请重新读取文件后重试。');
      if (String(this.current().FullName) !== path || !p.Saved) throw new Error('文稿在保存期间发生变化，本次操作已停止。');
      const token = new Date().toISOString().replace(/[:.]/g, '-') + '-' + guid().slice(1, 9);
      backup = path.replace(/\.pptx$/i, '') + '.表格样式备份-' + token + '.pptx';
      staging = path.replace(/\.pptx$/i, '') + '.table-style-pending-' + token + '.pptx';
      this.write(backup, original);
      this.write(staging, result.bytes);
      writeStarted = true;
      this.write(path, result.bytes);
      this.pending.set(path, { memorySignature, diskSignature: contentSignature(verified) });
      return { id: result.id, name: result.name, changedTables: result.changedTables, changedMasters: result.changedMasters, backup, path };
    } catch (error) {
      if (writeStarted && original) {
        try {
          this.write(path, original);
        } catch (recovery) {
          throw new Error(error.message + '；自动恢复未完成，请打开备份：' + backup + '。' + recovery.message);
        }
        throw new Error(error.message + '；已恢复修改前的文件。');
      }
      throw error;
    } finally {
      if (staging) this.remove(staging);
      this.busy = false;
    }
  }
}

export const REQUEST = 'table-style-editor.request.v1';
export const RESPONSE = 'table-style-editor.response.v1';
export const LAYOUT = 'table-style-editor.layout.v1';
export function createClient(app) {
  if (!app?.PluginStorage) return null;
  let pending = false;
  return {
    async call(method, params = {}) {
      if (pending) throw new Error('前一个操作尚未完成。');
      pending = true;
      const id = guid(), store = app.PluginStorage;
      store.setItem(REQUEST, JSON.stringify({ id, method, params }));
      try {
        return await new Promise((resolve, reject) => {
          let ticks = 0;
          const timer = setInterval(() => {
            try {
              const text = store.getItem(RESPONSE);
              const result = text ? JSON.parse(text) : null;
              if (result?.id === id) { clearInterval(timer); result.error ? reject(new Error(result.error)) : resolve(result.value); }
              else if (++ticks > 2400) { clearInterval(timer); reject(new Error('操作等待超时，请查看当前文件及备份后重新打开编辑器。')); }
            } catch (e) { clearInterval(timer); reject(e); }
          }, 250);
        });
      } finally { pending = false; }
    },
  };
}
