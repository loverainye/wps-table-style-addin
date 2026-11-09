import { binaryString } from '../src/host.js';

const CHUNK = 65536;
const MAX_BYTES = 512 * 1024 * 1024;

function loadLibc(ffi, signature, hints) {
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

// JSA has a native FFI interface, whereas Application.FileSystem belongs to
// the add-in host. Use only public libc file APIs, never WPS private offsets.
export function createLinuxFileSystem(ffi, hints = typeof __WPS_LIBC_PATHS__ === 'undefined' ? [] : __WPS_LIBC_PATHS__) {
  if (!ffi || typeof ffi.LoadLibrary !== 'function') {
    throw new Error('当前 JS 宏缺少文件读写接口，也没有 ffi.LoadLibrary，无法直接修改 PPTX。');
  }
  const { api, library } = loadLibc(ffi, {
    open: { returnType: 'int32', parameters: ['string', 'int32', 'uint32'] },
    read: { returnType: 'int64', parameters: ['int32', 'pointer', 'uint64'] },
    write: { returnType: 'int64', parameters: ['int32', 'pointer', 'uint64'] },
    fsync: { returnType: 'int32', parameters: ['int32'] },
    close: { returnType: 'int32', parameters: ['int32'] },
    access: { returnType: 'int32', parameters: ['string', 'int32'] },
    unlink: { returnType: 'int32', parameters: ['string'] },
  }, hints);

  const checkedPath = path => {
    if (typeof path !== 'string' || !path.startsWith('/') || path.includes('\0')) {
      throw new Error('文件路径必须是有效的 Linux 绝对路径。');
    }
    return path;
  };
  const error = (operation, path) => new Error(operation + '失败：' + path);
  const withFile = (path, flags, work) => {
    checkedPath(path);
    const fd = Number(api.open(path, flags, 0o600));
    if (!Number.isInteger(fd) || fd < 0) throw error('打开文件', path);
    let failure = null;
    try { return work(fd); }
    catch (e) { failure = e; throw e; }
    finally {
      // close() must not be retried: even on an error Linux may have released
      // the descriptor, and a retry could close an unrelated WPS file.
      if (Number(api.close(fd)) !== 0 && !failure) throw error('关闭文件', path);
    }
  };
  const read = path => withFile(path, 0, fd => {
    const buffer = new Uint8Array(CHUNK), parts = [];
    let total = 0;
    while (true) {
      const size = Number(api.read(fd, buffer.buffer, BigInt(buffer.length)));
      if (!Number.isInteger(size) || size < 0 || size > buffer.length) throw error('读取文件', path);
      if (size === 0) break;
      total += size;
      if (total > MAX_BYTES) throw new Error('宏版当前支持不超过 512 MiB 的 PPTX 文件。');
      parts.push(binaryString(buffer.subarray(0, size)));
    }
    return parts.join('');
  });
  return {
    backend: 'Linux libc / WPS FFI',
    library,
    readAsBinaryString: read,
    readFileString(path) {
      const raw = read(path);
      return decodeURIComponent(Array.from(raw, c => '%' + c.charCodeAt(0).toString(16).padStart(2, '0')).join(''));
    },
    writeAsBinaryString(path, data) {
      if (typeof data !== 'string' || data.length > MAX_BYTES) throw new Error('要写入的二进制数据无效或超过 512 MiB。');
      withFile(path, 1 | 64 | 512, fd => {
        for (let at = 0; at < data.length; at += CHUNK) {
          const buffer = Uint8Array.from(data.slice(at, at + CHUNK), c => c.charCodeAt(0) & 255);
          let sent = 0;
          while (sent < buffer.length) {
            const remaining = buffer.slice(sent);
            const size = Number(api.write(fd, remaining.buffer, BigInt(remaining.length)));
            if (!Number.isInteger(size) || size <= 0 || size > remaining.length) throw error('写入文件', path);
            sent += size;
          }
        }
        if (Number(api.fsync(fd)) !== 0) throw error('同步文件到磁盘', path);
      });
    },
    Exists(path) { return Number(api.access(checkedPath(path), 0)) === 0; },
    Remove(path) { if (Number(api.unlink(checkedPath(path))) !== 0) throw error('删除临时文件', path); },
  };
}

// Adapt just the members consumed by FileHost. Do not assign properties onto
// the WPS Application proxy or change the installed add-in's filesystem.
export function macroApplication(app, root = globalThis) {
  const fs = app.FileSystem;
  if (fs && typeof fs.readAsBinaryString === 'function' && typeof fs.writeAsBinaryString === 'function') return app;
  const files = createLinuxFileSystem(root.ffi);
  return {
    FileSystem: files,
    Env: { GetTempPath: () => '/tmp/' },
    FileDialog(type) { return app.FileDialog(type); },
    get ActivePresentation() { return app.ActivePresentation; },
    get ActiveWindow() { return app.ActiveWindow; },
    get Presentations() { return app.Presentations; },
  };
}
