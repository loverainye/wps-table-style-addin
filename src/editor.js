import { PARTS, EDGES, BORDER_MIN, BORDER_MAX, BORDER_STEP, readPart, editStyle, preset, guid, colorForPreview, parseXML } from './styles.js';
import { createClient, LAYOUT } from './host.js';
import { measureWindow } from './window-fit.js';
import { setupFormZoom } from './form-zoom.js';
import { FileHost } from './file-host.js';

const $ = id => document.getElementById(id);
let context, selected, patches = {}, drafts = new Map(), newStyles = [], busy = false;
let client = createClient(window.Application || window.wps);
const isPreview = new URLSearchParams(location.search).get('preview') === '1';
const isMacro = new URLSearchParams(location.search).get('macro') === '1';
if (isMacro) client = new FileHost(window.Application || window.wps);
let layout = 'workspace', fitTimer, fitGeneration = 0;
function cancelFloatingFit() { clearTimeout(fitTimer); fitGeneration++; }
function fitToScreen() {
  cancelFloatingFit();
  if (layout !== 'floating' || !client) return;
  const generation = fitGeneration;
  const attempt = (round = 0, waits = 0) => {
    fitTimer = setTimeout(async () => {
      if (generation !== fitGeneration || layout !== 'floating') return;
      if (busy) { if (waits < 10) attempt(round, waits + 1); return; }
      let changed = false;
      await run(async () => { changed = (await client.call('fit', { metrics: measureWindow(window) })).changed; });
      if (generation === fitGeneration && changed && round < 2) attempt(round + 1);
    }, 300);
  };
  attempt();
}
function showWindowMode() {
  if (isMacro) {
    document.querySelector('.layout-control').hidden = true;
    $('fit-window').hidden = true;
    return;
  }
  if (!isPreview) {
    const mode = (window.Application || window.wps)?.PluginStorage?.getItem(LAYOUT);
    if (mode === 'sidebar-max') layout = 'sidebar';
    else if (['workspace', 'floating', 'sidebar'].includes(mode)) layout = mode;
  }
  $('layout').value = layout;
  document.body.dataset.layout = layout;
  $('fit-window').hidden = false;
  $('fit-window').title = layout === 'floating' ? '按当前屏幕重新调整浮动窗口' : '展开至最大尺寸；已经最大时忽略点击';
}
const THEMES = { dk1: '#202020', lt1: '#FFFFFF', dk2: '#203864', lt2: '#E7E6E6', accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47', tx1: '#202020', tx2: '#203864', bg1: '#FFFFFF', bg2: '#E7E6E6' };
function message(text, kind = 'success') { $('status').hidden = false; $('status').textContent = text; $('status').dataset.kind = kind; }
async function run(work, dim = true) {
  if (busy) return;
  busy = true; if (dim) document.body.classList.add('busy'); $('apply').disabled = true; $('layout').disabled = true; $('fit-window').disabled = true;
  try { await work(); } catch (e) { message(e.message, 'error'); }
  finally { busy = false; document.body.classList.remove('busy'); $('apply').disabled = !selected; $('layout').disabled = false; $('fit-window').disabled = false; }
}
function flags() { return Object.fromEntries(Array.from(document.querySelectorAll('[data-flag]')).map(e => [e.dataset.flag, e.checked])); }
function editedXML() { return selected?.xml ? editStyle(selected.xml, $('style-name').value || selected.name, selected.id, patches) : null; }
function saveDraft() { if (selected) drafts.set(selected.id, { patches, name: $('style-name').value }); }
function populateStyles(id) {
  $('style-list').replaceChildren();
  for (const s of [...context.styles, ...newStyles]) {
    const o = document.createElement('option'); o.value = s.id;
    o.textContent = s.name + (s.id.toUpperCase() === context.defaultID.toUpperCase() ? ' · 默认' : '') + (s.create ? ' · 新建' : '');
    $('style-list').append(o);
  }
  if (id) {
    const match = [...context.styles, ...newStyles].find(s => s.id.toUpperCase() === id.toUpperCase());
    if (match) $('style-list').value = match.id;
  }
  selectStyle($('style-list').value);
}
function selectStyle(id) {
  selected = [...context.styles, ...newStyles].find(s => s.id === id);
  if (!selected) { message('没有可选的样式。请点击“新建”。'); return; }
  const draft = drafts.get(id);
  patches = draft?.patches || {};
  $('style-name').value = draft?.name || selected.name;
  $('style-name').disabled = !!selected.builtin;
  $('duplicate').disabled = !!selected.builtin;
  $('builtin-note').hidden = !selected.builtin;
  $('format-card').hidden = !!selected.builtin;
  $('set-default').checked = selected.id.toUpperCase() === context.defaultID.toUpperCase();
  $('style-id').textContent = selected.id;
  const example = context.tables.find(t => t.styleID.toUpperCase() === selected.id.toUpperCase());
  const f = example?.flags || { firstRow: true, bandRow: true };
  for (const e of document.querySelectorAll('[data-flag]')) e.checked = !!f[e.dataset.flag];
  renderFields(); renderPreview();
}
function addStyle(base, name) {
  saveDraft();
  const id = guid();
  const s = { id, name, xml: editStyle(base, name, id, {}), create: true, builtin: false };
  newStyles.push(s); populateStyles(id);
  $('new-form').hidden = true;
  $('scope').value = 'all'; $('set-default').checked = true;
}
function addColorOptions(select, allowNone) {
  for (const [value, label] of [['inherit', '继承'], ...(allowNone ? [['none', '无填充']] : []), ['solid', '自定义颜色'], ...Object.keys(THEMES).filter(k => !/^[tb][xg]/.test(k)).map((k, i) => ['theme:' + k, k.startsWith('accent') ? '主题强调色 ' + k.slice(-1) : '主题 ' + k]), ['unsupported', '保留现有设置']]) {
    const option = document.createElement('option'); option.value = value; option.textContent = label; select.append(option);
  }
}
function renderColor(prefix, value) {
  const mode = $(prefix + '-mode'), input = $(prefix === 'fill' ? 'fill-color' : 'text-color');
  mode.value = value.startsWith('#') ? 'solid' : value;
  input.value = colorForPreview(value, THEMES, '#FFFFFF');
  input.disabled = mode.value !== 'solid';
}
function renderFields() {
  if (!selected?.xml) return;
  const p = readPart(editedXML(), $('part').value);
  renderColor('fill', p.fill); renderColor('color', p.color);
  $('bold').value = p.bold; $('italic').value = p.italic;
  for (const [edge] of EDGES) {
    const b = p.borders[edge];
    $('border-mode-' + edge).value = b.mode;
    $('border-color-' + edge).value = colorForPreview(b.color, THEMES, '#FFFFFF');
    $('border-width-' + edge).value = b.width || 1;
    $('border-dash-' + edge).value = b.dash || 'solid';
  }
}
function patch(key, value) {
  const p = $('part').value;
  patches[p] ||= {};
  patches[p][key] = value;
  renderPreview();
}
function borderValue(edge) {
  return { mode: $('border-mode-' + edge).value, color: $('border-color-' + edge).value, width: Number($('border-width-' + edge).value), dash: $('border-dash-' + edge).value };
}
function stepBorderInput(edge, direction) {
  const input = $('border-width-' + edge);
  const value = Number(input.value);
  const base = Number.isFinite(value) && input.value.trim() !== '' ? value : 1;
  const next = Math.max(BORDER_MIN, Math.min(BORDER_MAX, Math.round((base + direction * BORDER_STEP) * 100) / 100));
  input.value = String(next);
  $('border-mode-' + edge).value = 'solid';
  patchBorder(edge);
}
function patchBorder(edge) {
  const p = $('part').value;
  patches[p] ||= {}; patches[p].borders ||= {}; patches[p].borders[edge] = borderValue(edge);
  renderPreview();
}
function renderPreview() {
  const table = $('preview'); table.replaceChildren();
  const f = flags();
  const style = editedXML();
  const parts = style ? Object.fromEntries(PARTS.map(([key]) => [key, readPart(style, key)])) : {};
  const data = [['项目', '目标', '完成', '达成率'], ['交付质量', '98%', '99%', '101%'], ['客户响应', '24h', '12h', '200%'], ['项目交付', '12项', '12项', '100%'], ['季度合计', '—', '已完成', '100%']];
  data.forEach((row, ri) => {
    const tr = document.createElement('tr');
    row.forEach((text, ci) => {
      const td = document.createElement(ri === 0 && f.firstRow ? 'th' : 'td'); td.textContent = text;
      const rules = ['wholeTbl'];
      if (f.bandRow) rules.push((ri - (f.firstRow ? 1 : 0)) % 2 === 0 ? 'band1H' : 'band2H');
      if (f.bandCol) rules.push((ci - (f.firstCol ? 1 : 0)) % 2 === 0 ? 'band1V' : 'band2V');
      if (f.lastCol && ci === row.length - 1) rules.push('lastCol');
      if (f.firstCol && ci === 0) rules.push('firstCol');
      if (f.lastRow && ri === data.length - 1) { rules.push('lastRow'); if (f.lastCol && ci === row.length - 1) rules.push('seCell'); if (f.firstCol && ci === 0) rules.push('swCell'); }
      if (f.firstRow && ri === 0) { rules.push('firstRow'); if (f.lastCol && ci === row.length - 1) rules.push('neCell'); if (f.firstCol && ci === 0) rules.push('nwCell'); }
      const result = { fill: '#FFFFFF', color: '#24332D', bold: 'off', italic: 'off', borders: {} };
      for (const key of rules) {
        const rule = parts[key]; if (!rule) continue;
        for (const k of ['fill', 'color', 'bold', 'italic']) if (!['inherit', 'unsupported', ''].includes(rule[k])) result[k] = rule[k];
        for (const [e, b] of Object.entries(rule.borders)) if (!['inherit', 'unsupported'].includes(b.mode)) result.borders[e] = b;
      }
      td.style.background = colorForPreview(result.fill, THEMES); td.style.color = colorForPreview(result.color, THEMES, '#202020');
      td.style.fontWeight = result.bold === 'on' ? '700' : '400'; td.style.fontStyle = result.italic === 'on' ? 'italic' : 'normal';
      for (const [edge, pos] of [['left', 'Left'], ['right', 'Right'], ['top', 'Top'], ['bottom', 'Bottom']]) {
        const inside = ['left', 'right'].includes(edge) ? 'insideV' : 'insideH';
        const isOutside = edge === 'left' ? ci === 0 : edge === 'right' ? ci === row.length - 1 : edge === 'top' ? ri === 0 : ri === data.length - 1;
        const b = result.borders[isOutside ? edge : inside] || result.borders[edge];
        td.style['border' + pos] = b?.mode === 'solid' ? `${b.width}pt ${b.dash === 'dot' ? 'dotted' : b.dash && b.dash !== 'solid' ? 'dashed' : 'solid'} ${colorForPreview(b.color, THEMES, '#FFFFFF')}` : '0 solid transparent';
      }
      tr.append(td);
    });
    table.append(tr);
  });
  table.closest('.preview-card').hidden = !!selected?.builtin;
}
async function refresh(preferredStyleID) {
  if (!client) throw new Error('请在 WPS 演示中，通过“表格样式编辑器 → 编辑样式”打开此面板。');
  context = await client.call('inspect');
  drafts = new Map(); newStyles = [];
  if (!context) { selected = undefined; $('workspace').hidden = true; $('macro-empty').hidden = false; return; }
  $('macro-empty').hidden = true;
  const font = context.fontDefaults;
  $('set-font-default').checked = false;
  $('set-font-default').disabled = !font?.masterCount;
  $('font-default-fields').hidden = true;
  $('font-default-size').value = font?.points ?? 18;
  $('font-default-current').textContent = !font?.masterCount ? '当前文稿没有可用的幻灯片母版。' : `当前：${font.mixed ? '母版或文字层级的默认字号不一致' : font.points ? font.points + ' 磅' : '未明确设置默认字号'} · ${font.masterCount} 个母版`;
  $('workspace').hidden = false;
  const selectedTable = context.tables.find(t => context.targets.some(x => x.slideID === t.slideID && x.shapeID === t.shapeID));
  // Select the saved style directly. Rendering a default first would briefly
  // show a different preview before switching back to the edited style.
  populateStyles(preferredStyleID || selectedTable?.styleID || context.defaultID || context.styles[0]?.id);
  showWindowMode();
}
function previewClient() {
  const styleXML = preset('green'), root = parseXML(styleXML).documentElement;
  let info = { path: '/预览/季度汇报.pptx', name: '季度汇报.pptx（界面预览）', targets: [], tableCount: 3, tables: [], fontDefaults: { masterCount: 1, points: 18, mixed: false }, defaultID: root.getAttribute('styleId'), styles: [{ id: root.getAttribute('styleId'), name: '季度汇报 · 绿色', xml: styleXML }] };
  return { async call(method, plan) {
    if (method === 'inspect') return info;
    if (method === 'fit') return { changed: false };
    if (method === 'resize') {
      return { layout: plan.layout };
    }
    const prev = info.styles.find(s => s.id === plan.id);
    const id = plan.create || plan.name !== prev?.name || Object.keys(plan.patches || {}).length ? guid() : prev.id;
    const xml = editStyle(plan.baseXML || prev.xml, plan.name, id, plan.patches);
    info = { ...info, styles: [...info.styles.filter(s => s.id !== plan.id), { id, name: plan.name, xml }], defaultID: plan.setDefault ? id : info.defaultID };
    if (plan.defaultFontSize !== undefined) info.fontDefaults = { masterCount: 1, points: plan.defaultFontSize, mixed: false };
    return { id, changedTables: plan.scope === 'none' ? 0 : 3, changedMasters: plan.defaultFontSize !== undefined ? 1 : 0, backup: '预览模式不会修改任何文件' };
  } };
}
for (const [value, label] of PARTS) { const o = document.createElement('option'); o.value = value; o.textContent = label; $('part').append(o); }
$('part').value = 'firstRow';
addColorOptions($('fill-mode'), true); addColorOptions($('color-mode'), false);
for (const [edge, label] of EDGES) {
  const row = document.createElement('div'); row.className = 'border-row';
  row.innerHTML = `<span>${label}</span><select id="border-mode-${edge}" aria-label="${label}边框方式"><option value="inherit">继承</option><option value="none">无边框</option><option value="solid">线条</option><option value="unsupported">保留</option></select><input id="border-color-${edge}" type="color" aria-label="${label}边框颜色" value="#ffffff"><div class="border-width-control"><input id="border-width-${edge}" type="number" min="${BORDER_MIN}" max="${BORDER_MAX}" step="${BORDER_STEP}" value="1" inputmode="decimal" aria-label="${label}边框磅数"><span class="border-width-steps"><button id="border-width-up-${edge}" type="button" aria-label="${label}边框增加 0.25 磅" title="增加 0.25 磅"><svg viewBox="0 0 16 10" aria-hidden="true"><path d="m4 7 4-4 4 4"/></svg></button><button id="border-width-down-${edge}" type="button" aria-label="${label}边框减少 0.25 磅" title="减少 0.25 磅，最低 0.25 磅"><svg viewBox="0 0 16 10" aria-hidden="true"><path d="m4 3 4 4 4-4"/></svg></button></span></div><select id="border-dash-${edge}" aria-label="${label}边框线型"><option value="solid">实线</option><option value="dash">虚线</option><option value="dot">点线</option><option value="dashDot">点划线</option><option value="lgDash">长虚线</option></select>`;
  $('border-fields').append(row);
  for (const suffix of ['mode', 'color', 'dash']) $('border-' + suffix + '-' + edge).addEventListener('change', () => patchBorder(edge));
  const width = $('border-width-' + edge);
  let widthBeforeFocus = width.value;
  const commitWidth = () => {
    const value = Number(width.value);
    width.value = String(Number.isFinite(value) && width.value.trim() !== '' ? Math.max(BORDER_MIN, Math.min(BORDER_MAX, value)) : 1);
    $('border-mode-' + edge).value = 'solid';
    patchBorder(edge);
    widthBeforeFocus = width.value;
  };
  width.addEventListener('focus', () => { widthBeforeFocus = width.value; });
  width.addEventListener('change', commitWidth);
  width.addEventListener('blur', () => { if (width.value !== widthBeforeFocus) commitWidth(); });
  width.addEventListener('keydown', event => {
    if (event.key !== 'ArrowUp' && event.key !== 'ArrowDown') return;
    event.preventDefault();
    stepBorderInput(edge, event.key === 'ArrowUp' ? 1 : -1);
    widthBeforeFocus = width.value;
  });
  for (const direction of ['up', 'down']) $('border-width-' + direction + '-' + edge).addEventListener('click', () => stepBorderInput(edge, direction === 'up' ? 1 : -1));
}
$('part').addEventListener('change', renderFields);
$('style-list').addEventListener('change', () => { saveDraft(); selectStyle($('style-list').value); });
$('style-name').addEventListener('input', () => { renderPreview(); });
$('new').addEventListener('click', () => { $('new-form').hidden = !$('new-form').hidden; });
$('cancel-new').addEventListener('click', () => { $('new-form').hidden = true; });
$('create').addEventListener('click', () => addStyle(preset($('preset').value), '我的' + $('preset').selectedOptions[0].textContent));
$('duplicate').addEventListener('click', () => addStyle(editedXML(), $('style-name').value + ' 副本'));
for (const [prefix, key, inputID] of [['fill', 'fill', 'fill-color'], ['color', 'color', 'text-color']]) {
  $(prefix + '-mode').addEventListener('change', () => {
    const mode = $(prefix + '-mode').value; $(inputID).disabled = mode !== 'solid';
    if (mode !== 'unsupported') patch(key, mode === 'solid' ? $(inputID).value : mode);
  });
  $(inputID).addEventListener('input', () => patch(key, $(inputID).value));
}
for (const key of ['bold', 'italic']) $(key).addEventListener('change', () => patch(key, $(key).value));
$('flags').addEventListener('change', renderPreview);
$('all-borders').addEventListener('click', () => {
  const top = borderValue('top');
  if (top.mode === 'unsupported') return;
  const p = $('part').value; patches[p] ||= {}; patches[p].borders = Object.fromEntries(EDGES.map(([e]) => [e, { ...top }]));
  renderFields(); renderPreview();
});
setupFormZoom({
  form: $('editor-form'), level: $('zoom-level'), slider: $('zoom-slider'),
  out: $('zoom-out'), plus: $('zoom-in'),
  storage: () => (window.Application || window.wps)?.PluginStorage || window.localStorage,
});
$('set-font-default').addEventListener('change', () => { $('font-default-fields').hidden = !$('set-font-default').checked; });
if (isMacro) {
  $('macro-files').hidden = false;
  $('open-file').addEventListener('click', () => run(async () => {
    if (await client.call('openFile')) { $('status').hidden = true; await refresh(); }
  }));
  $('use-current').addEventListener('click', () => run(async () => {
    await client.call('useCurrent'); $('status').hidden = true; await refresh();
  }));
}
$('layout').addEventListener('change', () => run(async () => {
  const requested = $('layout').value;
  cancelFloatingFit();
  try {
    if (!client) throw new Error('请在 WPS 演示加载项中切换布局。');
    const result = await client.call('resize', { layout: requested, metrics: measureWindow(window) });
    layout = result.layout;
  } finally { showWindowMode(); fitToScreen(); }
}));
$('fit-window').addEventListener('click', () => {
  if (layout === 'floating') { fitToScreen(); return; }
  run(async () => {
    if (!client) throw new Error('请在 WPS 演示加载项中调整布局。');
    await client.call('resize', { layout, metrics: measureWindow(window) });
    showWindowMode();
  }, false);
});
$('apply').addEventListener('click', () => run(async () => {
  if (!selected) throw new Error('请先新建或选择一个样式。');
  $('status').hidden = true;
  $('apply').title = '正在保存并应用…';
  $('apply').setAttribute('aria-busy', 'true');
  try {
    const plan = { path: context.path, id: selected.id, name: $('style-name').value, sourceXML: selected.xml || null, baseXML: selected.create ? selected.xml : undefined, create: !!selected.create, patches, flags: flags(), scope: $('scope').value, targets: context.targets, setDefault: $('set-default').checked, clearOverrides: $('clear-overrides').checked };
    if ($('set-font-default').checked) {
      if (!$('font-default-size').value || !$('font-default-size').checkValidity()) throw new Error('默认字号需为 1–4000 磅，最多两位小数。');
      plan.defaultFontSize = Number($('font-default-size').value);
    }
    const result = await client.call('apply', plan);
    await refresh(result.id);
  } finally {
    $('apply').title = '保存并应用（手动重新打开文稿查看）';
    $('apply').removeAttribute('aria-busy');
  }
}));
if (isPreview) client = previewClient();
let screenKey;
function checkScreen() {
  showWindowMode();
  const s = window.screen;
  const key = [s.availWidth, s.availHeight, s.availLeft, s.availTop, window.devicePixelRatio].join(':');
  if (screenKey !== undefined && key !== screenKey) fitToScreen();
  screenKey = key;
}
window.addEventListener('resize', checkScreen);
window.addEventListener('focus', checkScreen);
setInterval(checkScreen, 1500);
checkScreen();
showWindowMode();
run(async () => { await refresh(); fitToScreen(); });
