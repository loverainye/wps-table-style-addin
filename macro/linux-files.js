import { posixFiles, POSIX_SIGNATURE } from './platform/posix.js';

export function loadLibc(ffi, signature, hints) {
  // WPS FFI can reject a soname with "Library not found" before dlopen.
  // NixOS also has no /lib/libc.so.6. Build hints come from the running WPS's
  // own libc mapping, not from the build tool's potentially different libc.
  const candidates = [...new Set([...hints,
    '/lib/x86_64-linux-gnu/libc.so.6', '/usr/lib/x86_64-linux-gnu/libc.so.6',
    '/lib/aarch64-linux-gnu/libc.so.6', '/usr/lib/aarch64-linux-gnu/libc.so.6',
    '/lib64/libc.so.6', '/usr/lib64/libc.so.6', '/usr/lib/libc.so.6', '/lib/libc.so.6',
    '/run/current-system/sw/share/nix-ld/lib/libc.so.6', 'libc.so.6',
  ])];
  const failures = [];
  for (const library of candidates) {
    try {
      const api = ffi.LoadLibrary(library, signature);
      if (!api || Object.keys(signature).some(name => typeof api[name] !== 'function')) {
        throw new Error('文件接口不完整');
      }
      return { api, library };
    } catch (error) { failures.push(library + '：' + (error.message || String(error))); }
  }
  throw new Error('无法加载 Linux libc 文件接口。已尝试以下路径：\n' + failures.join('\n'));
}

export function createLinuxFileSystem(ffi, hints = typeof __WPS_LIBC_PATHS__ === 'undefined' ? [] : __WPS_LIBC_PATHS__) {
  if (typeof ffi?.LoadLibrary !== 'function') throw new Error('当前 JS 宏缺少文件读写接口，也没有 ffi.LoadLibrary，无法直接修改 PPTX。');
  const { api, library } = loadLibc(ffi, POSIX_SIGNATURE, hints);
  return posixFiles(api, library, 'linux', 1 | 64 | 512, () => '/tmp/');
}
