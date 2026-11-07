import { posix } from 'node:path';
import { unzipSync } from 'fflate';
import { XMLParser, XMLValidator } from 'fast-xml-parser';

export type XmlNode = Record<string, unknown>;
const parser = new XMLParser({
  preserveOrder: true, ignoreAttributes: false, attributeNamePrefix: '',
  removeNSPrefix: false, transformTagName: (name: string) => name.split(':').pop()!,
  parseTagValue: false, trimValues: false, ignoreDeclaration: true,
  ignorePiTags: true,
});
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
export class Archive {
  private readonly files: Record<string, Uint8Array>;
  constructor(data: Uint8Array) {
    let total = 0, entries = 0;
    this.files = unzipSync(data, { filter(file) {
      total += file.originalSize;
      if (++entries > 4096 || total > 64 << 20 || file.originalSize > 16 << 20) throw new Error('Office 文件解压后超过预览限制（64 MiB / 4096 项）');
      return /\.(xml|rels|png|jpe?g|gif|webp)$/i.test(file.name);
    } });
  }
  xml(path: string, required = true): XmlNode[] {
    const data = this.files[path];
    if (!data) { if (required) throw new Error(`Office 文件缺少 ${path}`); return []; }
    if (data.length > 8 << 20) throw new Error('单个 Office XML 超过 8 MiB 预览限制');
    const source = utf8.decode(data);
    if (/<!DOCTYPE|<!ENTITY/i.test(source) || XMLValidator.validate(source) !== true) throw new Error(`Office XML 无效：${path}`);
    return parser.parse(source) as XmlNode[];
  }
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
    const data = this.files[part];
    if (!data || data.length > 256 << 10) return '';
    const extension = posix.extname(part).toLowerCase();
    const mime: Record<string, string> = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif', '.webp': 'image/webp' };
    return mime[extension] ? `data:${mime[extension]};base64,${Buffer.from(data.buffer, data.byteOffset, data.length).toString('base64')}` : '';
  }
}
