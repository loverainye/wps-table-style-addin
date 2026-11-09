import { WpsHost, binaryString } from './host.js';
import { openDeck, applyStyle, guid } from './styles.js';

// Standalone macro windows call WPS directly; no registered add-in or polling
// ribbon page is required. Files outside the active presentation stay closed.
export class FileHost extends WpsHost {
  constructor(app) { super(app); this.path = null; }
  opened(path) {
    const docs = this.app.Presentations;
    for (let i = 1; i <= Number(docs.Count); i++) {
      const doc = docs.Item(i);
      if (String(doc.FullName) === path) return doc;
    }
    return null;
  }
  async openFile() {
    const picker = this.app.FileDialog(3);
    picker.Title = '打开需要编辑表格样式的 PPTX';
    picker.AllowMultiSelect = false;
    picker.Filters.Clear(); picker.Filters.Add('PowerPoint 演示文稿', '*.pptx');
    if (Number(picker.Show()) !== -1) return false;
    const path = String(picker.SelectedItems.Item(1));
    if (!/\.pptx$/i.test(path)) throw new Error('请选择 .pptx 文件，宏工具本身保留为 .pptm。');
    const doc = this.opened(path);
    if (doc && String(this.app.ActivePresentation?.FullName) !== path) throw new Error('该文件已在 WPS 中打开，请先切换到该文稿，再使用“当前文稿”。');
    // Parse before changing the target so a cancelled/invalid open keeps edits.
    await openDeck(this.read(path));
    this.path = path;
    return true;
  }
  async inspect() {
    const activePath = String(this.app.ActivePresentation?.FullName || '');
    if (!this.path && /\.pptx$/i.test(activePath)) this.path = activePath;
    if (!this.path) return null;
    const doc = this.opened(this.path);
    if (doc) {
      if (activePath !== this.path) throw new Error('目标文稿已打开，请切换到该文稿后再编辑样式。');
      return super.inspect();
    }
    const name = this.path.split(/[\\/]/).pop();
    return this.describe(await openDeck(this.read(this.path)), this.path, name, []);
  }
  async apply(plan) {
    if (plan.path !== this.path) throw new Error('编辑目标已切换，请重新读取目标文件。');
    if (this.opened(this.path)) {
      if (String(this.app.ActivePresentation?.FullName) !== this.path) throw new Error('目标文稿已打开，请切换到该文稿后再保存。');
      return super.apply(plan);
    }
    if (this.busy) throw new Error('正在保存样式，请稍候。');
    this.busy = true;
    let original, backup, staging, writeStarted = false;
    try {
      original = this.read(this.path);
      const deck = await openDeck(original);
      const existing = deck.styles.find(s => s.id.toUpperCase() === String(plan.id).toUpperCase());
      if (!plan.create && (existing?.xml || null) !== (plan.sourceXML || null)) throw new Error('文件中的样式已更改，请重新打开文件后重试。');
      const result = await applyStyle(deck, { ...plan, targets: [] });
      const verified = await openDeck(result.bytes);
      if (!verified.styles.some(s => s.id === result.id)) throw new Error('生成的样式校验失败。');
      if (plan.defaultFontSize !== undefined && (verified.fontDefaults.mixed || verified.fontDefaults.points !== plan.defaultFontSize)) throw new Error('生成的默认字号校验失败。');
      if (this.opened(this.path) || binaryString(this.read(this.path)) !== binaryString(original)) throw new Error('目标文件在编辑期间被打开或修改，本次保存已停止。');
      const token = new Date().toISOString().replace(/[:.]/g, '-') + '-' + guid().slice(1, 9);
      backup = this.path.replace(/\.pptx$/i, '') + '.表格样式备份-' + token + '.pptx';
      staging = this.path.replace(/\.pptx$/i, '') + '.table-style-pending-' + token + '.pptx';
      this.write(backup, original); this.write(staging, result.bytes);
      writeStarted = true; this.write(this.path, result.bytes);
      return { id: result.id, name: result.name, path: this.path, backup, changedTables: result.changedTables, changedMasters: result.changedMasters };
    } catch (error) {
      if (writeStarted && original) {
        try { this.write(this.path, original); }
        catch (recovery) { throw new Error(error.message + '；恢复失败，请打开备份：' + backup + '。' + recovery.message); }
      }
      throw error;
    } finally {
      if (staging) this.remove(staging);
      this.busy = false;
    }
  }
  async call(method, params) {
    if (method === 'openFile') return this.openFile();
    if (method === 'useCurrent') { this.path = null; return this.inspect(); }
    if (method === 'inspect') return this.inspect();
    if (method === 'apply') return this.apply(params);
    if (method === 'fit') return { changed: false };
    throw new Error('未知的宏编辑器操作。');
  }
}
