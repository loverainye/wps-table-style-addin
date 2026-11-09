import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';

// NixOS libc store paths are installation-specific. Include only paths read
// from running WPS processes, never a guessed store hash or Node's libc.
// Reading comm/maps neither attaches to nor changes any WPS process.
export async function wpsLibcPaths(proc = '/proc') {
  let entries;
  try { entries = await readdir(proc); } catch { return []; }
  const found = await Promise.allSettled(entries.filter(name => /^\d+$/.test(name)).map(async pid => {
    const base = path.join(proc, pid);
    if ((await readFile(path.join(base, 'comm'), 'utf8')).trim() !== 'wpp') return [];
    const maps = await readFile(path.join(base, 'maps'), 'utf8');
    return maps.split('\n').flatMap(line => {
      const match = line.match(/\s(\/[^\n]*\/libc(?:-[\d.]+)?\.so(?:\.6)?)$/);
      return match ? [match[1]] : [];
    });
  }));
  return [...new Set(found.flatMap(result => result.status === 'fulfilled' ? result.value : []))].sort();
}
