import { posix } from 'node:path';
import { inflateSync } from 'fflate';
import { openSync, closeSync, fstatSync, readSync } from 'node:fs';
import { parseXml, validateXml } from './xml.js';

export type XmlNode = Record<string, unknown>;
export function tag(node: XmlNode): string { return Object.keys(node).find(k => k !== ':@') ?? ''; }
export function children(node: XmlNode | undefined): XmlNode[] {
  if (!node) return [];
  const value = node[tag(node)];
  return Array.isArray(value) ? value as XmlNode[] : [];
}
export function child(node: XmlNode | undefined, name: string): XmlNode | undefined { return children(node).find(n => tag(n) === name); }
export function attr(node: XmlNode | undefined, name: string): string {
  const attributes = node?.[':@'] as Record<string, unknown> | undefined;
  if (!attributes) return '';
  if (Object.hasOwn(attributes, name)) return String(attributes[name]);
  const local = name.split(':').pop()!;
  const key = Object.keys(attributes).find(key => key.endsWith(`:${local}`));
  return key ? String(attributes[key]) : '';
}
export function* descendants(nodes: XmlNode[], name: string): Generator<XmlNode> {
  for (const node of nodes) {
    if (tag(node) === name) yield node;
    yield* descendants(children(node), name);
  }
}
export function text(node: XmlNode | undefined): string {
  if (!node) return '';
  if (Object.hasOwn(node, '#text')) return String(node['#text']);
  return children(node).map(text).join('');
}
const utf8 = new TextDecoder('utf-8', { fatal: true });
interface ZipEntry { offset: number; size: number; packed: number; method: number; crc: number }
export class Archive {
  private readonly files = new Map<string, ZipEntry>();
  private fd?: number;
  private data?: Uint8Array;
  private size: number;
  constructor(source: Uint8Array | string) {
    if (typeof source === 'string') {
      this.fd = openSync(source, 'r');
      try {
        const info = fstatSync(this.fd);
        if (!info.isFile() || info.size > 32 << 20) throw new Error('Office 预览只支持不超过 32 MiB 的普通文件');
        this.size = info.size;
      } catch (error) { this.close(); throw error; }
    } else { this.data = source; this.size = source.length; }
    try { this.index(); } catch (error) { this.close(); throw error; }
  }
  close(): void { if (this.fd !== undefined) { closeSync(this.fd); this.fd = undefined; } this.data = undefined; }
  private read(offset: number, length: number): Buffer {
    if (offset < 0 || length < 0 || offset+length > this.size) throw new Error('Office ZIP 范围无效');
    if (this.data) return Buffer.from(this.data.buffer, this.data.byteOffset+offset, length);
    if (this.fd === undefined) throw new Error('Office 文件已关闭');
    const result = Buffer.allocUnsafe(length);
    let used = 0;
    while (used < length) {
      const n = readSync(this.fd, result, used, length-used, offset+used);
      if (!n) throw new Error('文件正在变化，请刷新预览');
      used += n;
    }
    return result;
  }
  private index(): void {
    const tail = this.read(Math.max(0, this.size-65557), Math.min(this.size, 65557));
    let end = tail.length-22;
    for (; end >= 0; end--) if (tail.readUInt32LE(end) === 0x06054b50 && end+22+tail.readUInt16LE(end+20) === tail.length) break;
    if (end < 0 || tail.readUInt16LE(end+4) || tail.readUInt16LE(end+6)) throw new Error('Office ZIP 无效或不支持分卷');
    const count = tail.readUInt16LE(end+10), size = tail.readUInt32LE(end+12), offset = tail.readUInt32LE(end+16);
    if (count > 4096 || offset+size > this.size-tail.length+end) throw new Error('Office ZIP 目录超过限制或无效');
    const directory = this.read(offset, size);
    let position = 0, total = 0;
    for (let i = 0; i < count; i++) {
      if (position+46 > size || directory.readUInt32LE(position) !== 0x02014b50) throw new Error('Office ZIP 目录无效');
      const flags = directory.readUInt16LE(position+8), method = directory.readUInt16LE(position+10);
      const crc = directory.readUInt32LE(position+16), packed = directory.readUInt32LE(position+20), unpacked = directory.readUInt32LE(position+24);
      const n = directory.readUInt16LE(position+28), extra = directory.readUInt16LE(position+30), comment = directory.readUInt16LE(position+32);
      const local = directory.readUInt32LE(position+42);
      if (position+46+n+extra+comment > size || flags & 1 || ![0, 8].includes(method)) throw new Error('Office ZIP 加密或压缩格式不支持');
      total += unpacked;
      if (total > 64 << 20 || unpacked > 16 << 20) throw new Error('Office 文件解压后超过预览限制（64 MiB / 单项 16 MiB）');
      const name = directory.toString('utf8', position+46, position+46+n);
      if (this.files.has(name)) throw new Error('Office ZIP 包含重复文件');
      this.files.set(name, { offset: local, size: unpacked, packed, method, crc });
      position += 46+n+extra+comment;
    }
  }
  private load(path: string): Uint8Array | undefined {
    const entry = this.files.get(path);
    if (!entry) return undefined;
    const header = this.read(entry.offset, 30);
    if (header.readUInt32LE(0) !== 0x04034b50) throw new Error('Office ZIP 文件头无效');
    const start = entry.offset+30+header.readUInt16LE(26)+header.readUInt16LE(28);
    const packed = this.read(start, entry.packed);
    const result = entry.method === 0 ? packed : inflateSync(packed, { out: new Uint8Array(entry.size) });
    if (result.length !== entry.size || crc32(result) !== entry.crc) throw new Error('Office ZIP 校验失败，请刷新预览');
    return result;
  }
  source(path: string, required = true): string {
    const entry = this.files.get(path);
    if (entry && entry.size > 8 << 20) throw new Error('单个 Office XML 超过 8 MiB 预览限制');
    const data = this.load(path);
    if (!data) { if (required) throw new Error(`Office 文件缺少 ${path}`); return ''; }
    const source = utf8.decode(data);
    validateXml(source, path);
    return source;
  }
  xml(path: string, required = true): XmlNode[] { return parseXml(this.source(path, required)); }
  relations(part: string): Map<string, string> {
    const result = new Map<string, string>();
    const path = posix.join(posix.dirname(part), '_rels', `${posix.basename(part)}.rels`);
    for (const node of descendants(this.xml(path, false), 'Relationship')) {
      if (attr(node, 'TargetMode') === 'External') continue;
      const target = attr(node, 'Target');
      const normalized = posix.normalize(target.startsWith('/') ? target.slice(1) : posix.join(posix.dirname(part), target));
      if (normalized.startsWith('../') || /[\\\0?#]/.test(normalized)) continue;
      result.set(attr(node, 'Id'), normalized);
    }
    return result;
  }
  image(part: string | undefined): string {
    if (!part) return '';
    const entry = this.files.get(part);
    if (!entry || entry.size > 256 << 10) return '';
    const data = this.load(part);
    if (!data || data.length > 256 << 10) return '';
    const extension = posix.extname(part).toLowerCase();
    const mime: Record<string, string> = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp' };
    return mime[extension] ? `data:${mime[extension]};base64,${Buffer.from(data.buffer, data.byteOffset, data.length).toString('base64')}` : '';
  }
}

const crcTable = Uint32Array.from({ length: 256 }, (_, i) => {
  for (let bit = 0; bit < 8; bit++) i = i & 1 ? 0xedb88320 ^ (i >>> 1) : i >>> 1;
  return i >>> 0;
});
function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
