import { build } from 'esbuild';
import { cp, mkdir, rm, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { wpsLibcPaths } from './linux-libc.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export async function buildAddon(out = path.join(root, 'dist')) {
  await rm(out, { recursive: true, force: true });
  await mkdir(out, { recursive: true });
  await cp(path.join(root, 'public'), out, { recursive: true });
  for (const [entry, name] of [['ribbon', 'js/main'], ['editor', 'editor']]) {
    await build({ entryPoints: [path.join(root, 'src', entry + '.js')], outfile: path.join(out, name + '.js'), bundle: true, format: 'iife', platform: 'browser', target: 'chrome80', minify: true, sourcemap: true });
  }
}
export async function buildMacro(out = path.join(root, 'build', 'macro-runtime.js'), libcPaths) {
  return buildNativeRuntime('runtime.js', 'NativeTableStyles', out, libcPaths);
}
export async function buildSlideMacro(out = path.join(root, 'build', 'slide-macro-runtime.js'), libcPaths) {
  return buildNativeRuntime('slide-runtime.js', 'NativeSlideTableStyles', out, libcPaths);
}
async function buildNativeRuntime(entry, globalName, out, libcPaths) {
  await mkdir(path.dirname(out), { recursive: true });
  const hints = libcPaths ?? await wpsLibcPaths();
  await build({ entryPoints: [path.join(root, 'macro', entry)], outfile: out, bundle: true,
    define: { __WPS_LIBC_PATHS__: JSON.stringify(hints) },
    format: 'iife', globalName, platform: 'browser', target: 'chrome80',
    minifyIdentifiers: true, minifySyntax: true, minifyWhitespace: entry !== 'slide-runtime.js', legalComments: 'none' });
  // Minifiers may put literal ZIP signature control bytes in JS strings. JDE
  // stores this source in XML 1.0, so retain their JS value using escapes.
  const code = (await readFile(out, 'utf8')).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,
    c => '\\u' + c.charCodeAt(0).toString(16).padStart(4, '0'));
  await writeFile(out, code);
}
// Actual Linux WPS generates <script src="js/main.js"> in its local index.html.
// Keep the package free of index.html, per WPS offline add-in startup requirements.
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await buildAddon();
  await buildMacro();
  console.log('已生成 dist/：WPS 演示离线加载项；Linux/macOS 使用 install.sh，Windows 使用 install.ps1。');
}
