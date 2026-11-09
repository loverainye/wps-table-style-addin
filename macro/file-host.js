import { binaryString } from '../src/host.js';
import { openDeck, applyStyle, guid } from '../src/styles.js';

// UserForm edits the saved file directly. No SaveCopyAs, snapshots, temp
// directory, presentation opening/closing, or slideshow/image APIs are used.
export class MacroFileHost {
  constructor(app) { this.app = app; this.path = null; this.busy = false; }
  opened(path) {
    const docs = this.app.Presentations;
    for (let i = 1; i <= Number(docs.Count); i++) {
      const doc = docs.Item(i); if (String(doc.FullName) === path) return doc;
    }
    return null;
  }
  ready(path) {
    const doc = this.opened(path);
    if (!doc) return null;
    if (String(this.app.ActivePresentation?.FullName) !== path) throw new Error('目标文稿已在 WPS 打开，请切换到该文稿后使用“当前文稿”。');
    if (doc.ReadOnly) throw new Error('当前文稿为只读。');
    if (![true, 1, -1].includes(doc.Saved)) throw new Error('当前文稿有未保存修改。请先保存正文；若已写入新样式，请先手动重新打开文稿，避免旧内容覆盖新样式。');
    return doc;
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
  read(path) {
    const data = this.app.FileSystem.readAsBinaryString(path);
    if (typeof data !== 'string') throw new Error('WPS 返回了无法识别的文件数据。');
    return Uint8Array.from(data, c => c.charCodeAt(0) & 255);
  }
  write(path, bytes) {
    this.app.FileSystem.writeAsBinaryString(path, binaryString(bytes));
    if (binaryString(this.read(path)) !== binaryString(bytes)) throw new Error('文件写入校验失败：' + path);
  }
  async openFile() {
    const picker = this.app.FileDialog(3);
    picker.Title = '打开需要编辑表格样式的 PPTX'; picker.AllowMultiSelect = false;
    picker.Filters.Clear(); picker.Filters.Add('PowerPoint 演示文稿', '*.pptx');
    if (Number(picker.Show()) !== -1) return false;
    const path = String(picker.SelectedItems.Item(1));
    if (!/^(?:\/|[A-Za-z]:[\\/]|\\\\)/.test(path) || !/\.pptx$/i.test(path)) throw new Error('请选择本地 .pptx 文件，宏工具保留为 .pptm。');
    this.ready(path); await openDeck(this.read(path)); this.path = path;
    return true;
  }
  async inspect() {
    if (!this.path) {
      const active = String(this.app.ActivePresentation?.FullName || '');
      if (/^(?:\/|[A-Za-z]:[\\/]|\\\\)/.test(active) && /\.pptx$/i.test(active)) this.path = active;
    }
    if (!this.path) return null;
    const doc = this.ready(this.path), deck = await openDeck(this.read(this.path));
    return { path: this.path, name: this.path.split(/[\\/]/).pop(), targets: doc ? this.selection() : [],
      defaultID: deck.defaultID, styles: deck.styles, tables: deck.tables, tableCount: deck.tables.length, fontDefaults: deck.fontDefaults };
  }
  async apply(plan) {
    if (this.busy) throw new Error('正在保存样式，请稍候。');
    if (plan.path !== this.path) throw new Error('编辑目标已切换，请重新读取目标文件。');
    const path = this.path;
    this.busy = true;
    let original, backup, started = false;
    try {
      const doc = this.ready(path), targets = doc ? this.selection() : [];
      if (!doc && plan.scope === 'selected') throw new Error('请先切换到目标文稿并选中表格，或选择其它应用范围。');
      original = this.read(path);
      const deck = await openDeck(original);
      const existing = deck.styles.find(s => s.id.toUpperCase() === String(plan.id).toUpperCase());
      if (!plan.create && (existing?.xml || null) !== (plan.sourceXML || null)) throw new Error('文件中的样式已更改，请重新打开文件后重试。');
      const result = await applyStyle(deck, { ...plan, targets });
      const verified = await openDeck(result.bytes);
      if (!verified.styles.some(s => s.id === result.id)) throw new Error('生成的样式校验失败。');
      if (plan.defaultFontSize !== undefined && (verified.fontDefaults.mixed || verified.fontDefaults.points !== plan.defaultFontSize)) throw new Error('生成的默认字号校验失败。');
      if (this.path !== path || !!this.ready(path) !== !!doc || binaryString(this.read(path)) !== binaryString(original)) throw new Error('目标文件在编辑期间被打开或修改，本次保存已停止。');
      const token = new Date().toISOString().replace(/[:.]/g, '-') + '-' + guid().slice(1, 9);
      backup = path.replace(/\.pptx$/i, '') + '.表格样式备份-' + token + '.pptx';
      this.write(backup, original);
      started = true; this.write(path, result.bytes);
      return { id: result.id, name: result.name, path, backup, changedTables: result.changedTables, changedMasters: result.changedMasters };
    } catch (error) {
      if (started && original) try { this.write(path, original); }
      catch (recovery) { throw new Error(error.message + '；恢复失败，请打开备份：' + backup + '。' + recovery.message); }
      throw error;
    } finally { this.busy = false; }
  }
}
