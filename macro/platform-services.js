import { createLinuxFileSystem } from './linux-files.js';
import { createMacFileSystem } from './platform/macos.js';
import { createWindowsFileSystem } from './platform/windows.js';
import { decodeUTF8 } from './platform/common.js';
import { createUUID } from '../src/uuid.js';

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

// The editor sees one contract: files, tempPath(), uuid().
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
  return {
    get platform() { return platform || 'native'; }, files,
    get fileBackend() { return files.backend; },
    get uuidBackend() { return typeof crypto?.randomUUID === 'function' ? 'JS crypto.randomUUID' : typeof crypto?.getRandomValues === 'function' ? 'JS crypto.getRandomValues' : 'JS Math.random'; },
    uuid() { return createUUID(crypto, () => (root.Math || Math).random()); },
    tempPath() {
      for (const source of [member(app, 'Env'), member(root, 'Env')]) {
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
