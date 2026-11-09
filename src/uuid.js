let provider;
export function setUUIDProvider(fn) { provider = fn; }
export function uuidFromBytes(bytes) {
  const b = Uint8Array.from(bytes);
  if (b.length !== 16) throw new Error('UUID 必须包含 16 字节。');
  b[6] = (b[6] & 15) | 64; b[8] = (b[8] & 63) | 128;
  const s = Array.from(b, x => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `{${s.slice(0, 8)}-${s.slice(8, 12)}-${s.slice(12, 16)}-${s.slice(16, 20)}-${s.slice(20)}}`;
}
export function normalizeUUID(value) {
  const s = String(value).replace(/^\{(.*)\}$/, '$1');
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(s)) throw new Error('系统返回了无效的 UUID v4。');
  return '{' + s.toUpperCase() + '}';
}
export function createUUID(crypto, random = Math.random) {
  if (typeof crypto?.randomUUID === 'function') return normalizeUUID(crypto.randomUUID());
  const bytes = new Uint8Array(16);
  if (typeof crypto?.getRandomValues === 'function') crypto.getRandomValues(bytes);
  else for (let i = 0; i < bytes.length; i++) bytes[i] = Math.floor(random() * 256);
  return uuidFromBytes(bytes);
}
export function guid() {
  if (provider) return normalizeUUID(provider());
  return createUUID(globalThis.crypto);
}
