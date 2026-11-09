import { createLinuxFileSystem } from './linux-files.js';
import { createMacFileSystem, MAC_LIBRARIES } from './platform/macos.js';
import { createWindowsFileSystem } from './platform/windows.js';
import { decodeUTF8, loadLibrary, windowsABI } from './platform/common.js';
import { uuidFromBytes, normalizeUUID } from '../src/uuid.js';

function member(object, name) { try { return object?.[name]; } catch { return null; } }

export function detectPlatform(root = globalThis) {
  // Platform.OS() is supplied by WPS JSA, not by the JavaScript language.
  const value = String(root.Platform?.OS?.() || root.process?.platform || root.navigator?.platform || '').toLowerCase();
  if (/darwin|mac/.test(value)) return 'macos';
  if (/windows|^win/.test(value)) return 'windows';
  if (/linux/.test(value)) return 'linux';
  return value ? 'unsupported' : null;
}

function nativeFiles(app, root) {
  // JS macros can expose a global FileSystem while add-ins expose it on Application.
  for (const [fs, name] of [[member(app, 'FileSystem'), 'WPS Application.FileSystem'], [member(root, 'FileSystem'), 'WPS FileSystem']]) {
    if (typeof fs?.readAsBinaryString !== 'function' || typeof fs.writeAsBinaryString !== 'function') continue;
    return {
      backend: name,
      readAsBinaryString: path => fs.readAsBinaryString(path),
      writeAsBinaryString: (path, data) => fs.writeAsBinaryString(path, data),
      readFileString: path => typeof fs.readFileString === 'function' ? fs.readFileString(path) : decodeUTF8(fs.readAsBinaryString(path)),
      Exists(path) {
        if (typeof fs.Exists === 'function') return fs.Exists(path);
        if (typeof fs.existsSync === 'function') return fs.existsSync(path);
        throw new Error('WPS 文件接口缺少 Exists/existsSync。');
      },
      Remove(path) {
        if (typeof fs.Remove === 'function') return fs.Remove(path);
        if (typeof fs.unlinkSync === 'function') return fs.unlinkSync(path);
        throw new Error('WPS 文件接口缺少 Remove/unlinkSync。');
      },
      tempPath: typeof fs.tmpdir === 'function' ? () => fs.tmpdir() : undefined,
    };
  }
  return null;
}

// The editor sees one contract: files, tempPath(), uuid(), getRandomValues().
// Native WPS/JS facilities win; FFI libraries are loaded only when needed.
export function createPlatformServices(app, root = globalThis) {
  let platform = detectPlatform(root), fallback;
  const native = nativeFiles(app, root), crypto = root.crypto;
  const loadFallback = () => {
    if (fallback) return fallback;
    const factories = { linux: () => createLinuxFileSystem(root.ffi), macos: () => createMacFileSystem(root.ffi), windows: () => createWindowsFileSystem(root) };
    if (platform && !factories[platform]) throw new Error('不支持当前系统：' + platform);
    if (platform) fallback = factories[platform]();
    else {
      const failures = [];
      for (const [name, factory] of Object.entries(factories)) {
        try { fallback = factory(); platform = name; break; }
        catch (e) { failures.push(name + '：' + e.message); }
      }
      if (!fallback) throw new Error('无法识别宏文件接口。' + failures.join('\n'));
    }
    return fallback;
  };
  const files = native || loadFallback();
  let randomFill, randomBackend;
  const getRandomValues = bytes => {
    if (!ArrayBuffer.isView(bytes) || !/^(?:Uint|Int)(?:8|16|32)(?:Clamped)?Array$/.test(bytes.constructor.name)) throw new Error('随机数据缓冲区必须是整数 TypedArray。');
    if (bytes.byteLength > 65536) throw new Error('随机数据缓冲区最多为 65536 字节。');
    const view = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (typeof crypto?.getRandomValues === 'function') { randomBackend = 'JS crypto.getRandomValues'; crypto.getRandomValues(bytes); return bytes; }
    if (!randomFill) {
      let uuid;
      if (typeof crypto?.randomUUID === 'function') {
        randomBackend = 'JS crypto.randomUUID'; uuid = () => normalizeUUID(crypto.randomUUID());
      } else if (platform === 'linux') {
        randomBackend = 'Linux kernel UUID'; uuid = () => normalizeUUID(String(files.readFileString('/proc/sys/kernel/random/uuid')).replace(/\0/g, '').trim());
      } else if (platform === 'macos') {
        const { api } = loadLibrary(root.ffi, MAC_LIBRARIES, { getentropy: { returnType: 'int32', parameters: ['pointer', 'uint64'] } }, 'macOS 随机数接口');
        randomBackend = 'macOS getentropy';
        randomFill = view => {
          for (let at = 0; at < view.length; at += 256) {
            const buffer = new Uint8Array(Math.min(256, view.length - at));
            if (Number(api.getentropy(buffer.buffer, BigInt(buffer.length))) !== 0) throw new Error('macOS 获取系统随机数失败。');
            view.set(buffer, at);
          }
        };
      } else if (platform === 'windows') {
        const { api } = loadLibrary(root.ffi, ['bcrypt.dll', 'bcrypt'], {
          BCryptGenRandom: { returnType: 'int32', parameters: ['pointer', 'pointer', 'uint32', 'uint32'], ...windowsABI(root) },
        }, 'Windows 随机数接口');
        randomBackend = 'Windows BCryptGenRandom';
        randomFill = view => {
          const buffer = new Uint8Array(view.length);
          if (Number(api.BCryptGenRandom(null, buffer.buffer, buffer.length, 2)) !== 0) throw new Error('Windows 获取系统随机数失败。');
          view.set(buffer);
        };
      } else throw new Error('当前环境没有 JS Crypto，且无法识别系统。请使用提供 Platform.OS() 的 WPS JS 宏版本。');
      if (uuid) randomFill = view => {
        for (let at = 0; at < view.length;) {
          const hex = uuid().replace(/[{}-]/g, '');
          for (let i = 0; i < hex.length && at < view.length; i += 2) view[at++] = parseInt(hex.slice(i, i + 2), 16);
        }
      };
    }
    randomFill(view);
    return bytes;
  };
  return {
    get platform() { return platform || 'native'; }, files,
    get fileBackend() { return files.backend; },
    get uuidBackend() { return randomBackend || (typeof crypto?.randomUUID === 'function' ? 'JS crypto.randomUUID' : typeof crypto?.getRandomValues === 'function' ? 'JS crypto.getRandomValues' : 'system (lazy)'); },
    getRandomValues,
    uuid() {
      if (typeof crypto?.randomUUID === 'function') { randomBackend = 'JS crypto.randomUUID'; return normalizeUUID(crypto.randomUUID()); }
      return uuidFromBytes(getRandomValues(new Uint8Array(16)));
    },
    tempPath() {
      for (const source of [app.Env, root.Env]) {
        if (typeof source?.GetTempPath === 'function') {
          const path = source.GetTempPath(); if (path) return String(path);
        }
      }
      if (files.tempPath) return String(files.tempPath());
      return String(loadFallback().tempPath());
    },
  };
}

export function macroApplication(app, root = globalThis) {
  const services = createPlatformServices(app, root);
  return {
    FileSystem: services.files, Env: { GetTempPath: () => services.tempPath() }, TableStylePlatform: services,
    FileDialog(type) { return app.FileDialog(type); },
    get ActivePresentation() { return app.ActivePresentation; },
    get ActiveWindow() { return app.ActiveWindow; },
    get Presentations() { return app.Presentations; },
  };
}
