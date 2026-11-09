import { binaryString } from '../../src/host.js';

export { binaryString };
export const CHUNK = 65536;
export const MAX_BYTES = 512 * 1024 * 1024;
export const ioError = (operation, path) => new Error(operation + '失败：' + path);
export function checkData(data) {
  if (typeof data !== 'string' || data.length > MAX_BYTES) throw new Error('要写入的二进制数据无效或超过 512 MiB。');
}
export function checkPath(path, platform) {
  const absolute = typeof path === 'string' && (platform === 'windows' ? /^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+)/.test(path) : path.startsWith('/'));
  if (typeof path !== 'string' || !absolute || /[\0\r\n]/.test(path)) throw new Error('文件路径必须是有效的 ' + platform + ' 绝对路径。');
  return path;
}
export function decodeUTF8(raw) {
  return decodeURIComponent(Array.from(raw, c => '%' + c.charCodeAt(0).toString(16).padStart(2, '0')).join(''));
}
export function loadLibrary(ffi, candidates, signature, label) {
  if (typeof ffi?.LoadLibrary !== 'function') throw new Error('当前 JS 宏缺少文件读写接口，也没有 ffi.LoadLibrary，无法直接修改 PPTX。');
  const failures = [];
  for (const library of [...new Set(candidates)]) {
    try {
      const api = ffi.LoadLibrary(library, signature);
      if (!api || Object.keys(signature).some(name => typeof api[name] !== 'function')) throw new Error('文件接口不完整');
      return { api, library };
    } catch (error) { failures.push(library + '：' + (error.message || String(error))); }
  }
  throw new Error('无法加载 ' + label + '。已尝试以下路径：\n' + failures.join('\n'));
}
export function windowsABI(root) {
  const size = root.Platform?.PointerSize?.() ?? root.FFI?.PointerSize;
  return size === 4 ? { abi: 'stdcall' } : {};
}
