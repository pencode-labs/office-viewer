import { XMLParser, XMLValidator } from 'fast-xml-parser';
import type { XmlNode } from './archive.js';

const parser = new XMLParser({
  preserveOrder: true, ignoreAttributes: false, attributeNamePrefix: '',
  removeNSPrefix: false, transformTagName: (name: string) => name.split(':').pop()!,
  parseTagValue: false, trimValues: false, ignoreDeclaration: true, ignorePiTags: true,
});
export function validateXml(source: string, path: string): void {
  if (/<!DOCTYPE|<!ENTITY/i.test(source) || XMLValidator.validate(source) !== true) throw new Error(`Office XML 无效：${path}`);
}
export function parseXml(source: string): XmlNode[] { return parser.parse(source) as XmlNode[]; }
export interface XmlRange { start: number; end: number; head: number; name: string }

// Index byte-independent string ranges, without constructing an XML object
// tree. Comments, CDATA and quoted '>' are consumed as a single token.
export function* elements(source: string, name: string, begin = 0, end = source.length): Generator<XmlRange> {
  const token = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?([\w:.-]+)(?:[^<>"']|"[^"]*"|'[^']*')*>/g;
  token.lastIndex = begin;
  let start = -1, depth = 0, head = 0;
  for (let match; (match = token.exec(source));) {
    if (match.index >= end) break;
    const full = match[0], tag = match[1]?.split(':').pop();
    if (!tag) continue;
    const closing = full.startsWith('</'), empty = full.endsWith('/>');
    if (start < 0) {
      if (closing || tag !== name) continue;
      start = match.index; head = token.lastIndex;
      if (empty) { yield { start, head, end: head, name }; start = -1; }
      else depth = 1;
    } else {
      if (closing) depth--;
      else if (!empty) depth++;
      if (!depth) { yield { start, head, end: token.lastIndex, name }; start = -1; }
    }
  }
}
const attributes = new Map<string, RegExp>();
export function rangeAttribute(source: string, range: XmlRange, name: string): string {
  let pattern = attributes.get(name);
  if (!pattern) { pattern = new RegExp(`(?:\\s|:)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`); attributes.set(name, pattern); }
  return pattern.exec(source.slice(range.start, range.head))?.slice(1).find(v => v !== undefined) ?? '';
}

// Direct children of a source fragment, with matching closing-tag boundaries.
export function* directElements(source: string): Generator<XmlRange & { tail: number }> {
  const token = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<\/?([\w:.-]+)(?:[^<>"']|"[^"]*"|'[^']*')*>/g;
  let start = 0, head = 0, depth = 0, name = '';
  for (let match; (match = token.exec(source));) {
    if (!match[1]) continue;
    const closing = match[0].startsWith('</'), empty = match[0].endsWith('/>');
    if (!depth && !closing) {
      start = match.index; head = token.lastIndex; name = match[1].split(':').pop()!;
    }
    if (closing) depth--; else if (!empty) depth++;
    if (!depth) yield { start, head, tail: empty ? head : match.index, end: token.lastIndex, name };
  }
}
