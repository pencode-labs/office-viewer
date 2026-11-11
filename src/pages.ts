import { Archive, attr, child, children, descendants, tag, text } from './archive.js';
import type { XmlNode } from './archive.js';
import { enabled, escape } from './render.js';
import { directElements, elements, parseXml } from './xml.js';

type Token = { html: string; close?: string; leave?: boolean };
const PageSize = 128 << 10;
function* run(node: XmlNode, archive: Archive, relations: Map<string, string>): Generator<Token> {
  const props = child(node, 'rPr');
  let open = '', close = '';
  for (const [name, element] of [['b', 'strong'], ['i', 'em'], ['u', 'u'], ['strike', 's']]) {
    if (enabled(child(props, name!)) && (name !== 'u' || attr(child(props, 'u'), 'val') !== 'none')) { open += `<${element}>`; close = `</${element}>` + close; }
  }
  for (const c of children(node)) {
    if (tag(c) === 't') {
      const value = text(c);
      for (let offset = 0; offset < value.length;) {
        let end = Math.min(value.length, offset+4096);
        if (end < value.length && /[\uD800-\uDBFF]/.test(value[end-1]!)) end--;
        yield { html: open+escape(value.slice(offset, end))+close }; offset = end;
      }
    } else if (['br', 'cr'].includes(tag(c))) yield { html: '<br />' };
    else if (tag(c) === 'tab') yield { html: '&#160;&#160;&#160;&#160;' };
    else if (['drawing', 'pict'].includes(tag(c))) yield { html: picture(c, archive, relations) };
  }
}
function picture(node: XmlNode, archive: Archive, relations: Map<string, string>): string {
  const blip = descendants([node], 'blip').next().value;
  const src = archive.image(relations.get(attr(blip, 'embed')));
  return src ? `<img src="${src}" alt="文档图片" />` : '<span class="muted">[图片不可预览]</span>';
}
function* paragraph(node: XmlNode, archive: Archive, relations: Map<string, string>, word: boolean): Generator<Token> {
  const props = child(node, 'pPr'), heading = /^heading([1-6])$/i.exec(attr(child(props, 'pStyle'), 'val'));
  const element = heading ? `h${heading[1]}` : 'p';
  const alignment = attr(child(props, word ? 'jc' : 'pPr'), word ? 'val' : 'algn') || attr(props, 'algn');
  const align: Record<string, string> = { left: 'left', l: 'left', center: 'center', ctr: 'center', right: 'right', r: 'right', both: 'justify', just: 'justify' };
  const list = word ? !!child(props, 'numPr') : !!child(props, 'buChar') || !!child(props, 'buAutoNum');
  yield { html: `<${element}${align[alignment] ? ` style="text-align:${align[alignment]}"` : ''}${list ? ' class="list-item"' : ''}>`, close: `</${element}>` };
  if (list) yield { html: '• ' };
  let populated = false;
  for (const c of children(node)) {
    if (['r', 'fld'].includes(tag(c))) { yield* run(c, archive, relations); populated = true; }
    else if (tag(c) === 'br') { yield { html: '<br />' }; populated = true; }
    else if (['hyperlink', 'ins', 'sdt'].includes(tag(c))) {
      for (const r of descendants([c], 'r')) { yield* run(r, archive, relations); populated = true; }
    }
  }
  if (!populated) yield { html: '<br />' };
  yield { html: `</${element}>`, leave: true };
}
// Walk ranges in the XML source. Only a requested paragraph or picture becomes
// an object tree; large tables are traversed row by row, including nested ones.
function* tokens(source: string, archive: Archive, relations: Map<string, string>, word: boolean, depth = 0): Generator<Token> {
  for (const range of directElements(source)) {
    const name = range.name;
    if (name === 'p') yield* paragraph(parseXml(source.slice(range.start, range.end))[0]!, archive, relations, word);
    else if (name === 'pic') yield { html: picture(parseXml(source.slice(range.start, range.end))[0]!, archive, relations) };
    else if (['tbl', 'tr', 'tc'].includes(name)) {
      if (depth > 24) { yield { html: '<p>[嵌套表格超过预览层数]</p>' }; continue; }
      let cell = '<td>';
      if (name === 'tc') {
        const content = source.slice(range.head, range.tail);
        const properties = directElements(content).next().value;
        if (properties?.name === 'tcPr') {
          const span = Number(attr(descendants(parseXml(content.slice(properties.start, properties.end)), 'gridSpan').next().value, 'val'));
          if (span > 1 && span <= 256) cell = `<td colspan="${span}">`;
        }
      }
      const open = name === 'tbl' ? '<table><tbody>' : name === 'tr' ? '<tr>' : cell;
      const close = name === 'tbl' ? '</tbody></table>' : name === 'tr' ? '</tr>' : '</td>';
      yield { html: open, close };
      yield* tokens(source.slice(range.head, range.tail), archive, relations, word, depth+1);
      yield { html: close, leave: true };
    } else if (['document', 'body', 'sdt', 'sdtContent', 'sld', 'cSld', 'spTree', 'sp', 'txBody', 'graphicFrame', 'graphic', 'graphicData', 'grpSp'].includes(name)) {
      if (depth < 128) yield* tokens(source.slice(range.head, range.tail), archive, relations, word, depth+1);
    }
  }
}
function* pages(source: string, archive: Archive, relations: Map<string, string>, word: boolean): Generator<string> {
  const stack: { html: string; close: string }[] = [];
  let html = '', size = 0;
  for (const token of tokens(source, archive, relations, word)) {
    if (size && size + token.html.length > PageSize && !token.leave) {
      yield html + stack.toReversed().map(t => t.close).join('');
      html = stack.map(t => t.html).join(''); size = 0;
    }
    html += token.html; size += token.html.length;
    if (token.close) stack.push({ html: token.html, close: token.close });
    if (token.leave) stack.pop();
  }
  if (size) yield html;
}
export class PagedPreview {
  private readonly source: string;
  private readonly relations: Map<string, string>;
  private iterator: Generator<string>;
  private generated = 0;
  private readonly cache = new Map<number, string>();
  private end?: number;
  index = 0;
  constructor(private readonly archive: Archive, part: string, private readonly word: boolean) {
    this.source = archive.source(part);
    this.relations = archive.relations(part);
    this.iterator = pages(this.source, archive, this.relations, word);
    if (word && !elements(this.source, 'body').next().value) throw new Error('Word 文件缺少正文');
  }
  private get(index: number): string | undefined {
    if (this.cache.has(index)) return this.cache.get(index);
    if (this.end !== undefined && index >= this.end) return undefined;
    if (index < this.generated) { this.iterator = pages(this.source, this.archive, this.relations, this.word); this.generated = 0; }
    while (this.generated <= index) {
      const result = this.iterator.next();
      if (result.done) { this.end = this.generated; return undefined; }
      this.cache.set(this.generated++, result.value);
      if (this.cache.size > 3) this.cache.delete(this.cache.keys().next().value!);
    }
    return this.cache.get(index);
  }
  get next(): boolean { return this.get(this.index+1) !== undefined; }
  move(delta: number): boolean {
    const index = this.index+delta;
    if (index < 0 || this.get(index) === undefined) return false;
    this.index = index; return true;
  }
  render(): string { return this.get(this.index) ?? '<p class="muted">此页面没有可预览的正文。</p>'; }
}
