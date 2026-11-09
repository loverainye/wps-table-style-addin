import { loadLibrary } from './common.js';
import { POSIX_SIGNATURE, posixFiles } from './posix.js';

export const MAC_LIBRARIES = ['/usr/lib/libSystem.B.dylib', 'libSystem.B.dylib'];
export function createMacFileSystem(ffi) {
  const { api, library } = loadLibrary(ffi, MAC_LIBRARIES, {
    ...POSIX_SIGNATURE,
    open: { returnType: 'int32', parameters: ['string', 'int32'] },
    creat: { returnType: 'int32', parameters: ['string', 'uint32'] },
    getenv: { returnType: 'string', parameters: ['string'] },
  }, 'macOS libSystem 文件接口');
  // Apple arm64 passes variadic arguments on the stack. Use fixed-signature
  // creat(path, mode) for writing so no variadic open(..., mode) ABI is needed.
  return posixFiles(api, library, 'macos', 1, () => {
    const dir = api.getenv('TMPDIR');
    return typeof dir === 'string' && dir.startsWith('/') ? dir : '/tmp/';
  }, (path, flags, mode) => flags === 0 ? api.open(path, 0) : api.creat(path, mode));
}
