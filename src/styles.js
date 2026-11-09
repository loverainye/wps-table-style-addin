import JSZip from 'jszip';
import { guid } from './uuid.js';
export { guid };

export const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const PARTS = [
  ['wholeTbl', '整个表格'], ['band1H', '第一行条纹'], ['band2H', '第二行条纹'],
  ['band1V', '第一列条纹'], ['band2V', '第二列条纹'], ['lastCol', '最后一列'],
  ['firstCol', '第一列'], ['lastRow', '汇总行'], ['seCell', '最后一个汇总单元格'],
  ['swCell', '第一个汇总单元格'], ['firstRow', '标题行'], ['neCell', '最后一个标题单元格'],
  ['nwCell', '第一个标题单元格'],
];
export const EDGES = [['left', '左'], ['right', '右'], ['top', '上'], ['bottom', '下'], ['insideH', '内部横线'], ['insideV', '内部竖线']];
export const BORDER_MIN = 0.25;
export const BORDER_MAX = 20;
export const BORDER_STEP = 0.25;
const COLORS = ['srgbClr', 'schemeClr', 'prstClr', 'sysClr', 'scrgbClr', 'hslClr'];
const FILLS = ['solidFill', 'noFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
export const child = (node, local, ns = A) => Array.from(node?.childNodes || []).find(n => n.nodeType === 1 && n.localName === local && n.namespaceURI === ns);
export const children = (node, local, ns = A) => Array.from(node?.childNodes || []).filter(n => n.nodeType === 1 && n.namespaceURI === ns && (!local || n.localName === local));
export function parseXML(text) {
  if (/<!DOCTYPE/i.test(text)) throw new Error('不支持包含 DTD 的 XML。');
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (!doc.documentElement || doc.getElementsByTagName('parsererror').length) throw new Error('文稿中的 XML 无法解析。');
  return doc;
}
export const xml = node => new XMLSerializer().serializeToString(node);
const create = (doc, tag) => doc.createElementNS(A, 'a:' + tag);
function ensure(node, tag) {
  let item = child(node, tag);
  if (!item) { item = create(node.ownerDocument, tag); node.appendChild(item); }
  return item;
}
function remove(node, names) {
  for (const n of children(node)) if (names.includes(n.localName)) node.removeChild(n);
}
function readColor(node) {
  const c = children(node).find(n => COLORS.includes(n.localName));
  if (!c) return 'inherit';
  if (c.localName === 'srgbClr') return '#' + c.getAttribute('val');
  if (c.localName === 'schemeClr') return 'theme:' + c.getAttribute('val');
  if (c.localName === 'sysClr' && c.getAttribute('lastClr')) return '#' + c.getAttribute('lastClr');
  return 'unsupported';
}
function writeColor(node, value) {
  remove(node, COLORS);
  if (value === 'inherit') return;
  let c;
  if (/^#[\da-f]{6}$/i.test(value)) { c = create(node.ownerDocument, 'srgbClr'); c.setAttribute('val', value.slice(1).toUpperCase()); }
  else if (/^theme:(accent[1-6]|dk[12]|lt[12]|tx[12]|bg[12]|hlink|folHlink)$/.test(value)) { c = create(node.ownerDocument, 'schemeClr'); c.setAttribute('val', value.slice(6)); }
  else throw new Error('无效的颜色值：' + value);
  node.appendChild(c);
}
function readFill(node) {
  if (child(node, 'noFill')) return 'none';
  const solid = child(node, 'solidFill');
  return solid ? readColor(solid) : children(node).some(n => FILLS.includes(n.localName)) ? 'unsupported' : 'inherit';
}
export function readPart(styleXML, partName) {
  const root = parseXML(styleXML).documentElement;
  const part = child(root, partName), tx = child(part, 'tcTxStyle'), cell = child(part, 'tcStyle');
  const borders = {};
  for (const [edge] of EDGES) {
    const e = child(child(cell, 'tcBdr'), edge), ln = child(e, 'ln');
    borders[edge] = !e ? { mode: 'inherit' } : child(ln, 'noFill') ? { mode: 'none' } : ln ? {
      mode: 'solid', color: readColor(child(ln, 'solidFill')),
      width: Number(ln.getAttribute('w') || 12700) / 12700,
      dash: child(ln, 'prstDash')?.getAttribute('val') || 'solid',
    } : { mode: 'unsupported' };
  }
  return {
    fill: readFill(child(cell, 'fill')),
    color: readColor(tx),
    bold: onOff(tx?.getAttribute('b')),
    italic: onOff(tx?.getAttribute('i')),
    borders,
  };
}
function onOff(value) { return ['on', '1', 'true'].includes(value) ? 'on' : ['off', '0', 'false'].includes(value) ? 'off' : 'inherit'; }
export function editStyle(styleXML, name, id, patches) {
  const doc = parseXML(styleXML), root = doc.documentElement;
  root.setAttribute('styleId', id); root.setAttribute('styleName', name.trim());
  for (const [partName, patch] of Object.entries(patches || {})) {
    if (!PARTS.some(([p]) => p === partName)) throw new Error('未知的表格区域。');
    const part = ensure(root, partName);
    for (const key of ['bold', 'italic']) {
      if (!(key in patch)) continue;
      const tx = ensure(part, 'tcTxStyle'), attr = key === 'bold' ? 'b' : 'i';
      if (!['inherit', 'on', 'off'].includes(patch[key])) throw new Error('无效的字体设置。');
      if (patch[key] === 'inherit') tx.removeAttribute(attr); else tx.setAttribute(attr, patch[key]);
    }
    if ('color' in patch) writeColor(ensure(part, 'tcTxStyle'), patch.color);
    if ('fill' in patch) {
      const cell = ensure(part, 'tcStyle');
      remove(cell, ['fill', 'fillRef']);
      if (patch.fill !== 'inherit') {
        const f = ensure(cell, 'fill');
        if (patch.fill === 'none') f.appendChild(create(doc, 'noFill'));
        else writeColor(ensure(f, 'solidFill'), patch.fill);
      }
    }
    for (const [edge, rule] of Object.entries(patch.borders || {})) {
      if (!EDGES.some(([e]) => e === edge)) throw new Error('未知的边框方向。');
      const bdr = ensure(ensure(part, 'tcStyle'), 'tcBdr');
      const prev = child(bdr, edge); if (prev) bdr.removeChild(prev);
      if (rule.mode === 'inherit') continue;
      const ln = ensure(ensure(bdr, edge), 'ln');
      if (rule.mode === 'none') { ln.appendChild(create(doc, 'noFill')); continue; }
      if (rule.mode !== 'solid' || !Number.isFinite(Number(rule.width)) || Number(rule.width) < BORDER_MIN || Number(rule.width) > BORDER_MAX) throw new Error('边框线宽需为 0.25–20 磅。');
      ln.setAttribute('w', Math.round(Number(rule.width) * 12700));
      writeColor(ensure(ln, 'solidFill'), rule.color);
      const dash = ensure(ln, 'prstDash');
      if (!['solid', 'dash', 'dot', 'dashDot', 'lgDash'].includes(rule.dash || 'solid')) throw new Error('无效线型。');
      dash.setAttribute('val', rule.dash || 'solid');
    }
    // DrawingML child order matters to Office, particularly text before cell style.
    const tx = child(part, 'tcTxStyle');
    if (tx && !child(tx, 'fontRef') && !child(tx, 'font')) {
      const ref = create(doc, 'fontRef'); ref.setAttribute('idx', 'minor');
      const c = create(doc, 'schemeClr'); c.setAttribute('val', 'phClr'); ref.appendChild(c);
      tx.insertBefore(ref, tx.firstChild);
    }
    for (const tag of ['tcTxStyle', 'tcStyle']) { const n = child(part, tag); if (n) part.appendChild(n); }
    const cell = child(part, 'tcStyle');
    for (const tag of ['tcBdr', 'fill', 'fillRef', 'cell3D']) { const n = child(cell, tag); if (n) cell.appendChild(n); }
    const bdr = child(cell, 'tcBdr');
    for (const [tag] of EDGES) { const n = child(bdr, tag); if (n) bdr.appendChild(n); }
  }
  for (const [tag] of PARTS) { const n = child(root, tag); if (n) root.appendChild(n); }
  const ext = child(root, 'extLst'); if (ext) root.appendChild(ext);
  return xml(root);
}
export function preset(name = 'green') {
  const [header, odd, even] = {
    green: ['385723', 'D5E3CF', 'EBF1E9'], red: ['C00000', 'FCE4E4', 'FFF2F2'],
    blue: ['2455A4', 'DFEAF9', 'F3F7FE'], blank: ['FFFFFF', 'FFFFFF', 'FFFFFF'],
  }[name] || ['385723', 'D5E3CF', 'EBF1E9'];
  const borders = Object.fromEntries(EDGES.map(([e]) => [e, { mode: 'solid', color: '#FFFFFF', width: 1, dash: 'solid' }]));
  return editStyle(`<a:tblStyle xmlns:a="${A}"/>`, '新建表格样式', guid(), {
    wholeTbl: { fill: '#' + even, color: '#20242A', bold: 'off', italic: 'off', borders },
    band1H: { fill: '#' + odd }, band2H: { fill: '#' + even },
    firstRow: { fill: '#' + header, color: name === 'blank' ? '#20242A' : '#FFFFFF', bold: 'on' },
  });
}
const upper = id => String(id || '').toUpperCase();
const masterPaths = zip => Object.keys(zip.files).filter(n => /^ppt\/slideMasters\/[^/]+\.xml$/.test(n));
async function readFontDefaults(zip) {
  const paths = masterPaths(zip), sizes = [];
  for (const path of paths) {
    const master = parseXML(await zip.file(path).async('string')).documentElement;
    const other = child(child(master, 'txStyles', P), 'otherStyle', P);
    const fallback = child(child(other, 'defPPr'), 'defRPr')?.getAttribute('sz');
    for (let level = 1; level <= 9; level++) {
      const value = child(child(other, `lvl${level}pPr`), 'defRPr')?.getAttribute('sz') || fallback;
      sizes.push(value && Number.isFinite(Number(value)) ? Number(value) / 100 : null);
    }
  }
  const values = [...new Set(sizes)];
  return { masterCount: paths.length, points: values.length === 1 ? values[0] : null, mixed: values.length > 1 };
}
async function setFontDefaults(zip, points) {
  if (typeof points !== 'number' || !Number.isFinite(points) || points < 1 || points > 4000 || Math.abs(points * 100 - Math.round(points * 100)) > 0.000001) throw new Error('默认字号需为 1–4000 磅，最多两位小数。');
  const paths = masterPaths(zip);
  if (!paths.length) throw new Error('文稿中没有幻灯片母版，无法设置默认字号。');
  for (const path of paths) {
    const doc = parseXML(await zip.file(path).async('string')), root = doc.documentElement;
    let tx = child(root, 'txStyles', P);
    if (!tx) {
      tx = doc.createElementNS(P, 'p:txStyles');
      root.insertBefore(tx, child(root, 'extLst', P) || null);
    }
    let other = child(tx, 'otherStyle', P);
    if (!other) {
      other = doc.createElementNS(P, 'p:otherStyle');
      tx.insertBefore(other, child(tx, 'extLst', P) || null);
    }
    const tags = ['defPPr', ...Array.from({ length: 9 }, (_, i) => `lvl${i + 1}pPr`)];
    for (const [index, tag] of tags.entries()) {
      let para = child(other, tag);
      if (!para) {
        para = create(doc, tag);
        const next = children(other).find(n => tags.indexOf(n.localName) > index || n.localName === 'extLst');
        other.insertBefore(para, next || null);
      }
      let run = child(para, 'defRPr');
      if (!run) {
        run = create(doc, 'defRPr');
        para.insertBefore(run, child(para, 'extLst') || null);
      }
      run.setAttribute('sz', String(Math.round(points * 100)));
    }
    zip.file(path, xml(doc));
  }
  return paths.length;
}
export async function openDeck(binary) {
  if (binary.length > 100 * 1024 * 1024) throw new Error('当前版本支持 100 MB 以内的演示稿。');
  const zip = await JSZip.loadAsync(binary, { checkCRC32: true });
  const size = Object.values(zip.files).reduce((n, e) => n + (e._data?.uncompressedSize || 0), 0);
  if (size > 500 * 1024 * 1024) throw new Error('文稿解压后超过 500 MB，无法在加载项内处理。');
  if (!zip.file('ppt/presentation.xml')) throw new Error('当前文件不是有效的 PowerPoint OOXML 演示稿。');
  if (Object.keys(zip.files).some(n => n.startsWith('_xmlsignatures/'))) throw new Error('当前文稿带有数字签名，修改样式会使签名失效。请先另存为无签名副本。');
  const stylesXML = await zip.file('ppt/tableStyles.xml')?.async('string') || `<a:tblStyleLst xmlns:a="${A}" def=""/>`;
  const root = parseXML(stylesXML).documentElement;
  if (root.namespaceURI !== A || root.localName !== 'tblStyleLst') throw new Error('表格样式资源格式不正确。');
  const styles = children(root, 'tblStyle').map(n => ({ id: n.getAttribute('styleId'), name: n.getAttribute('styleName'), xml: xml(n), builtin: false }));
  const tables = [];
  const slideIDs = {};
  const presentation = parseXML(await zip.file('ppt/presentation.xml').async('string'));
  const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
  const rels = relsFile ? parseXML(await relsFile.async('string')) : null;
  for (const s of Array.from(presentation.getElementsByTagNameNS(P, 'sldId'))) {
    const rel = Array.from(rels?.getElementsByTagNameNS(PKG, 'Relationship') || []).find(r => r.getAttribute('Id') === s.getAttributeNS(R, 'id'));
    if (rel) {
      const target = rel.getAttribute('Target');
      const path = target.startsWith('/') ? target.slice(1) : 'ppt/' + target.replace(/^\.\//, '');
      slideIDs[path] = s.getAttribute('id');
    }
  }
  for (const path of Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort()) {
    const doc = parseXML(await zip.file(path).async('string'));
    for (const frame of Array.from(doc.getElementsByTagNameNS(P, 'graphicFrame'))) {
      const tbl = frame.getElementsByTagNameNS(A, 'tbl')[0];
      if (!tbl) continue;
      const nv = frame.getElementsByTagNameNS(P, 'cNvPr')[0];
      const id = child(child(tbl, 'tblPr'), 'tableStyleId')?.textContent || root.getAttribute('def');
      const props = child(tbl, 'tblPr');
      const flags = Object.fromEntries(['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'].map(key => [key, ['1', 'true', 'on'].includes(props?.getAttribute(key))]));
      tables.push({ path, slideID: slideIDs[path], shapeID: nv?.getAttribute('id'), name: nv?.getAttribute('name') || '表格', styleID: id, flags });
    }
  }
  for (const id of [root.getAttribute('def'), ...tables.map(t => t.styleID)]) {
    if (id && !styles.some(s => upper(s.id) === upper(id))) styles.push({ id, name: '内置样式 ' + id.slice(1, 9), builtin: true, xml: null });
  }
  return { zip, stylesXML, defaultID: root.getAttribute('def'), styles, tables, fontDefaults: await readFontDefaults(zip) };
}
function removeOverrides(tbl) {
  for (const tc of Array.from(tbl.getElementsByTagNameNS(A, 'tcPr'))) remove(tc, [...FILLS, 'lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr']);
  for (const tag of ['rPr', 'defRPr', 'endParaRPr']) {
    for (const r of Array.from(tbl.getElementsByTagNameNS(A, tag))) { remove(r, FILLS); r.removeAttribute('b'); r.removeAttribute('i'); }
  }
}
async function ensureStyleRelationship(zip) {
  const path = 'ppt/_rels/presentation.xml.rels';
  const d = parseXML(await zip.file(path)?.async('string') || `<Relationships xmlns="${PKG}"/>`);
  if (!Array.from(d.documentElement.childNodes).some(n => n.getAttribute?.('Type') === R + '/tableStyles')) {
    const ids = new Set(Array.from(d.documentElement.childNodes).map(n => n.getAttribute?.('Id')));
    let i = 1; while (ids.has('rIdTableStyles' + i)) i++;
    const n = d.createElementNS(PKG, 'Relationship');
    n.setAttribute('Id', 'rIdTableStyles' + i); n.setAttribute('Type', R + '/tableStyles'); n.setAttribute('Target', 'tableStyles.xml');
    d.documentElement.appendChild(n); zip.file(path, xml(d));
  }
  const ct = parseXML(await zip.file('[Content_Types].xml').async('string'));
  const NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  if (!Array.from(ct.getElementsByTagNameNS(NS, 'Override')).some(n => n.getAttribute('PartName') === '/ppt/tableStyles.xml')) {
    const n = ct.createElementNS(NS, 'Override');
    n.setAttribute('PartName', '/ppt/tableStyles.xml'); n.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml');
    ct.documentElement.appendChild(n); zip.file('[Content_Types].xml', xml(ct));
  }
}
export async function applyStyle(deck, plan) {
  if (!plan.name?.trim() || plan.name.trim().length > 100) throw new Error('样式名称需为 1–100 个字符。');
  if (!['linked', 'selected', 'all', 'none'].includes(plan.scope)) throw new Error('请选择应用范围。');
  const original = deck.styles.find(s => upper(s.id) === upper(plan.id));
  if (!original && !plan.create) throw new Error('原样式已改变，请刷新后重试。');
  const modified = plan.create || plan.name !== original?.name || Object.keys(plan.patches || {}).length > 0;
  if (modified && original?.builtin && !plan.baseXML) throw new Error('内置样式没有可编辑定义，请先新建一个自定义样式。');
  const id = modified ? guid() : original.id;
  const customXML = modified ? editStyle(plan.baseXML || original.xml, plan.name, id, plan.patches) : null;
  const d = parseXML(deck.stylesXML), root = d.documentElement;
  if (customXML) root.appendChild(d.importNode(parseXML(customXML).documentElement, true));
  if (plan.setDefault || (modified && upper(deck.defaultID) === upper(plan.id))) root.setAttribute('def', id);
  const targetSet = new Set((plan.targets || []).map(t => `${t.slideID}:${t.shapeID}`));
  if (plan.scope === 'selected' && !targetSet.size) throw new Error('请先在幻灯片上选中表格，再点击“保存并应用”。');
  const affected = deck.tables.filter(t => plan.scope === 'all' || (plan.scope === 'linked' && upper(t.styleID) === upper(plan.id)) || (plan.scope === 'selected' && targetSet.has(`${t.slideID}:${t.shapeID}`)));
  if (plan.scope === 'selected' && !affected.length) throw new Error('未找到选中的表格，请重新选择后再保存并应用。');
  // Clone the ZIP container so a failed operation never mutates the loaded source.
  const zip = deck.zip.clone();
  zip.files = Object.assign(Object.create(null), deck.zip.files);
  let changedTables = 0;
  for (const path of [...new Set(affected.map(t => t.path))]) {
    const slide = parseXML(await zip.file(path).async('string'));
    const ids = new Set(affected.filter(t => t.path === path).map(t => t.shapeID));
    for (const frame of Array.from(slide.getElementsByTagNameNS(P, 'graphicFrame'))) {
      if (!ids.has(frame.getElementsByTagNameNS(P, 'cNvPr')[0]?.getAttribute('id'))) continue;
      const tbl = frame.getElementsByTagNameNS(A, 'tbl')[0]; if (!tbl) continue;
      let props = child(tbl, 'tblPr');
      if (!props) { props = create(slide, 'tblPr'); tbl.insertBefore(props, tbl.firstChild); }
      ensure(props, 'tableStyleId').textContent = id;
      for (const [key, val] of Object.entries(plan.flags || {})) {
        if (['firstRow', 'lastRow', 'firstCol', 'lastCol', 'bandRow', 'bandCol'].includes(key)) props.setAttribute(key, val ? '1' : '0');
      }
      if (plan.clearOverrides) removeOverrides(tbl);
      changedTables++;
    }
    zip.file(path, xml(slide));
  }
  // Drop the obsolete custom definition only when no table/default still uses it.
  if (modified && original && !original.builtin && upper(root.getAttribute('def')) !== upper(original.id)) {
    const unaffectedUses = deck.tables.some(t => upper(t.styleID) === upper(original.id) && !affected.includes(t));
    if (!unaffectedUses) { const old = children(root, 'tblStyle').find(n => upper(n.getAttribute('styleId')) === upper(original.id)); if (old) root.removeChild(old); }
  }
  const changedMasters = plan.defaultFontSize === undefined ? 0 : await setFontDefaults(zip, plan.defaultFontSize);
  if (customXML || root.getAttribute('def') !== deck.defaultID || !zip.file('ppt/tableStyles.xml')) zip.file('ppt/tableStyles.xml', xml(d));
  await ensureStyleRelationship(zip);
  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { bytes, id, changedTables, changedMasters, modified, name: plan.name.trim() };
}

export function colorForPreview(value, theme = {}, fallback = '#FFFFFF') {
  if (value === 'none') return 'transparent';
  if (/^#[\da-f]{6}$/i.test(value)) return value;
  if (value?.startsWith('theme:')) return theme[value.slice(6)] || fallback;
  return fallback;
}
