import { createUUID } from '../src/uuid.js';

function member(object, name) { try { return object?.[name]; } catch { return null; } }

function nativeFiles(app, root) {
  // JS macros can expose a global FileSystem while add-ins expose it on Application.
  for (const [owner, name] of [[app, 'WPS Application.FileSystem'], [root, 'WPS FileSystem']]) {
    const fs = member(owner, 'FileSystem');
    if (typeof fs?.readAsBinaryString !== 'function' || typeof fs.writeAsBinaryString !== 'function') continue;
    return {
      backend: name,
      readAsBinaryString: path => fs.readAsBinaryString(path),
      writeAsBinaryString: (path, data) => fs.writeAsBinaryString(path, data),
      readFileString: path => typeof fs.readFileString === 'function' ? fs.readFileString(path) :
        decodeURIComponent(Array.from(fs.readAsBinaryString(path), c => '%' + c.charCodeAt(0).toString(16).padStart(2, '0')).join('')),
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
    };
  }
  return null;
}

// WPS supplies file access; JavaScript supplies UUIDs. No OS detection or system libraries.
export function createPlatformServices(app, root = globalThis) {
  const files = nativeFiles(app, root), crypto = root.crypto;
  if (!files) throw new Error('当前 WPS JS 宏未提供原生二进制文件读写接口（FileSystem.readAsBinaryString / writeAsBinaryString）。请使用支持这些接口的 WPS 版本。');
  return {
    files,
    get fileBackend() { return files.backend; },
    get uuidBackend() { return typeof crypto?.randomUUID === 'function' ? 'JS crypto.randomUUID' : typeof crypto?.getRandomValues === 'function' ? 'JS crypto.getRandomValues' : 'JS Math.random'; },
    uuid() { return createUUID(crypto, () => (root.Math || Math).random()); },
  };
}

export function macroApplication(app, root = globalThis) {
  const services = createPlatformServices(app, root);
  return {
    FileSystem: services.files, TableStylePlatform: services,
    // Only the retained legacy slide renderer uses WPS's own Env object.
    get Env() { return member(app, 'Env') || member(root, 'Env'); },
    FileDialog(type) { return app.FileDialog(type); },
    get ActivePresentation() { return app.ActivePresentation; },
    get ActiveWindow() { return app.ActiveWindow; },
    get Presentations() { return app.Presentations; },
  };
}
