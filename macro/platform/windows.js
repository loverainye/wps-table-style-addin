import { CHUNK, MAX_BYTES, binaryString, checkData, checkPath, decodeUTF8, ioError, loadLibrary, windowsABI } from './common.js';

export function createWindowsFileSystem(root) {
  const fn = (returnType, parameters) => ({ returnType, parameters, ...windowsABI(root) });
  const { api, library } = loadLibrary(root.ffi, ['kernel32.dll', 'kernel32'], {
    CreateFileW: fn('pointer', ['u16string', 'uint32', 'uint32', 'pointer', 'uint32', 'uint32', 'pointer']),
    ReadFile: fn('int32', ['pointer', 'pointer', 'uint32', 'pointer', 'pointer']),
    WriteFile: fn('int32', ['pointer', 'pointer', 'uint32', 'pointer', 'pointer']),
    FlushFileBuffers: fn('int32', ['pointer']), CloseHandle: fn('int32', ['pointer']),
    GetFileAttributesW: fn('uint32', ['u16string']), DeleteFileW: fn('int32', ['u16string']),
    GetTempPathW: fn('uint32', ['uint32', 'pointer']),
  }, 'Windows 文件接口');
  const withFile = (path, write, work) => {
    checkPath(path, 'windows');
    const handle = api.CreateFileW(path, write ? 0x40000000 : 0x80000000, 7, null, write ? 2 : 3, 128, null);
    const value = handle?.ToBigInt?.();
    if (!handle || value === 0 || value === 0n || value === -1 || value === -1n || value === 0xffffffff || value === 0xffffffffn || value === 0xffffffffffffffffn) throw ioError('打开文件', path);
    let failure;
    try { return work(handle); } catch (e) { failure = e; throw e; }
    finally { if (!Number(api.CloseHandle(handle)) && !failure) throw ioError('关闭文件', path); }
  };
  const read = path => withFile(path, false, handle => {
    const buffer = new Uint8Array(CHUNK), count = new Uint32Array(1), parts = [];
    let total = 0;
    while (true) {
      count[0] = 0;
      if (!Number(api.ReadFile(handle, buffer.buffer, buffer.length, count.buffer, null)) || count[0] > buffer.length) throw ioError('读取文件', path);
      if (!count[0]) break;
      total += count[0];
      if (total > MAX_BYTES) throw new Error('宏版当前支持不超过 512 MiB 的 PPTX 文件。');
      parts.push(binaryString(buffer.subarray(0, count[0])));
    }
    return parts.join('');
  });
  return {
    platform: 'windows', library, backend: 'Windows kernel32 / WPS FFI',
    readAsBinaryString: read, readFileString: path => decodeUTF8(read(path)),
    writeAsBinaryString(path, data) {
      checkData(data);
      withFile(path, true, handle => {
        const count = new Uint32Array(1);
        for (let at = 0; at < data.length; at += CHUNK) {
          const buffer = Uint8Array.from(data.slice(at, at + CHUNK), c => c.charCodeAt(0) & 255);
          for (let sent = 0; sent < buffer.length;) {
            const remaining = buffer.slice(sent); count[0] = 0;
            if (!Number(api.WriteFile(handle, remaining.buffer, remaining.length, count.buffer, null)) || !count[0] || count[0] > remaining.length) throw ioError('写入文件', path);
            sent += count[0];
          }
        }
        if (!Number(api.FlushFileBuffers(handle))) throw ioError('同步文件到磁盘', path);
      });
    },
    Exists(path) { return Number(api.GetFileAttributesW(checkPath(path, 'windows'))) !== 0xffffffff; },
    Remove(path) { if (!Number(api.DeleteFileW(checkPath(path, 'windows')))) throw ioError('删除临时文件', path); },
    tempPath() {
      const buffer = new Uint16Array(32768), size = Number(api.GetTempPathW(buffer.length, buffer.buffer));
      if (!Number.isInteger(size) || size <= 0 || size >= buffer.length) throw new Error('无法获取 Windows 临时目录。');
      return String.fromCharCode(...buffer.subarray(0, size));
    },
  };
}
