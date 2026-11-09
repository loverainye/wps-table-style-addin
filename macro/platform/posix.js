import { CHUNK, MAX_BYTES, binaryString, checkData, checkPath, decodeUTF8, ioError } from './common.js';

export const POSIX_SIGNATURE = {
  open: { returnType: 'int32', parameters: ['string', 'int32', 'uint32'] },
  read: { returnType: 'int64', parameters: ['int32', 'pointer', 'uint64'] },
  write: { returnType: 'int64', parameters: ['int32', 'pointer', 'uint64'] },
  fsync: { returnType: 'int32', parameters: ['int32'] },
  close: { returnType: 'int32', parameters: ['int32'] },
  access: { returnType: 'int32', parameters: ['string', 'int32'] },
  unlink: { returnType: 'int32', parameters: ['string'] },
};

// Linux and Darwin share these function signatures, but not their open flags.
export function posixFiles(api, library, platform, writeFlags, tempPath, openFile = (path, flags, mode) => api.open(path, flags, mode)) {
  const withFile = (path, flags, work) => {
    checkPath(path, platform);
    const fd = Number(openFile(path, flags, 0o600));
    if (!Number.isInteger(fd) || fd < 0) throw ioError('打开文件', path);
    let failure;
    try { return work(fd); } catch (e) { failure = e; throw e; }
    finally { if (Number(api.close(fd)) !== 0 && !failure) throw ioError('关闭文件', path); }
  };
  const read = path => withFile(path, 0, fd => {
    const buffer = new Uint8Array(CHUNK), parts = [];
    let total = 0;
    while (true) {
      const size = Number(api.read(fd, buffer.buffer, BigInt(buffer.length)));
      if (!Number.isInteger(size) || size < 0 || size > buffer.length) throw ioError('读取文件', path);
      if (size === 0) break;
      total += size;
      if (total > MAX_BYTES) throw new Error('宏版当前支持不超过 512 MiB 的 PPTX 文件。');
      parts.push(binaryString(buffer.subarray(0, size)));
    }
    return parts.join('');
  });
  return {
    platform, library, backend: platform === 'linux' ? 'Linux libc / WPS FFI' : 'macOS libSystem / WPS FFI',
    tempPath, readAsBinaryString: read,
    readFileString: path => decodeUTF8(read(path)),
    writeAsBinaryString(path, data) {
      checkData(data);
      withFile(path, writeFlags, fd => {
        for (let at = 0; at < data.length; at += CHUNK) {
          const buffer = Uint8Array.from(data.slice(at, at + CHUNK), c => c.charCodeAt(0) & 255);
          let sent = 0;
          while (sent < buffer.length) {
            const remaining = buffer.slice(sent);
            const size = Number(api.write(fd, remaining.buffer, BigInt(remaining.length)));
            if (!Number.isInteger(size) || size <= 0 || size > remaining.length) throw ioError('写入文件', path);
            sent += size;
          }
        }
        if (Number(api.fsync(fd)) !== 0) throw ioError('同步文件到磁盘', path);
      });
    },
    Exists(path) { return Number(api.access(checkPath(path, platform), 0)) === 0; },
    Remove(path) { if (Number(api.unlink(checkPath(path, platform))) !== 0) throw ioError('删除临时文件', path); },
  };
}
