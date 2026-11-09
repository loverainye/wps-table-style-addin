import { loadLibrary } from './common.js';
import { POSIX_SIGNATURE, posixFiles } from './posix.js';

export const MAC_LIBRARIES = ['/usr/lib/libSystem.B.dylib', 'libSystem.B.dylib'];
export function createMacFileSystem(ffi) {
  const { api, library } = loadLibrary(ffi, MAC_LIBRARIES, {
    ...POSIX_SIGNATURE,
    getenv: { returnType: 'string', parameters: ['string'] },
  }, 'macOS libSystem 文件接口');
  return posixFiles(api, library, 'macos', 1 | 512 | 1024, () => {
    const dir = api.getenv('TMPDIR');
    return typeof dir === 'string' && dir.startsWith('/') ? dir : '/tmp/';
  });
}
