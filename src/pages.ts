import { Archive, attr, child, children, descendants, tag, text } from './archive.js';
import type { XmlNode } from './archive.js';
import { enabled, escape } from './render.js';
import { directElements, elements, parseXml } from './xml.js';

type Token = { html: string; close?: string; leave?: boolean };
type TextFormat = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  css: string;
};
const PageSize = 128 << 10;
const colors: Record<string, string> = {
  black: '#000000', blue: '#0000ff', cyan: '#00ffff', darkBlue: '#00008b', darkCyan: '#008b8b',
  darkGray: '#a9a9a9', darkGreen: '#006400', darkMagenta: '#8b008b', darkRed: '#8b0000',
  darkYellow: '#b8860b', gray: '#808080', green: '#008000', magenta: '#ff00ff', red: '#ff0000',
  white: '#ffffff', yellow: '#ffff00',
};
function hexColor(value: string): string | undefined {
  if (/^[\da-f]{6}$/i.test(value)) return `#${value.toLowerCase()}`;
  return colors[value];
}
function points(value: string, divisor: number): string | undefined {
  if (!value) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? `${number / divisor}pt` : undefined;
}
function emuPoints(value: string): string | undefined { return points(value, 12700); }
function last(nodes: XmlNode[], name: string): XmlNode | undefined {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const node = child(nodes[i], name);
    if (node) return node;
  }
  return undefined;
}
function lastValue(nodes: XmlNode[], name: string, attribute: string): string {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const node = child(nodes[i], name);
    if (node) return attr(node, attribute);
  }
  return '';
}
function bool(nodes: XmlNode[], ...names: string[]): boolean {
  for (const name of names) {
    const node = last(nodes, name);
    if (node) return enabled(node);
    for (let i = nodes.length - 1; i >= 0; i--) {
      const value = attr(nodes[i], name);
      if (value) return !['0', 'false', 'off', 'none'].includes(value);
    }
  }
  return false;
}
function cssStyle(entries: [string, string | undefined][]): string {
  return entries.filter((entry): entry is [string, string] => !!entry[1]).map(([name, value]) => `${name}:${value};`).join('');
}
function format(nodes: XmlNode[], word: boolean): TextFormat {
  const fonts = last(nodes, 'rFonts');
  const latin = last(nodes, 'latin');
  const font = attr(fonts, word ? 'ascii' : 'latin') || attr(fonts, 'hAnsi') || attr(fonts, 'eastAsia') || attr(latin, 'typeface');
  const runProperties = nodes.at(-1);
  const size = lastValue(nodes, 'sz', 'val') || lastValue(nodes, 'rPr', 'sz') || attr(runProperties, 'sz');
  const fontSize = size ? points(size, word ? 2 : 100) : undefined;
  const colorNode = last(nodes, 'color');
  const fill = last(nodes, 'solidFill');
  const fillColor = hexColor(attr(last(fill ? [fill] : [], 'srgbClr'), 'val'));
  const color = hexColor(attr(colorNode, 'val')) ?? fillColor;
  const highlight = hexColor(lastValue(nodes, 'highlight', 'val')) ?? hexColor(lastValue(nodes, 'shd', 'fill'));
  const vertical = lastValue(nodes, 'vertAlign', 'val') || attr(runProperties, 'baseline');
  const baseline = Number(lastValue(nodes, 'baseline', 'val')) || Number(attr(runProperties, 'baseline'));
  const baselineValue = vertical === 'superscript' || baseline > 0 ? 'super' : vertical === 'subscript' || baseline < 0 ? 'sub' : undefined;
  const caps = bool(nodes, 'caps') ? 'uppercase' : bool(nodes, 'smallCaps') ? 'small-caps' : undefined;
  const underline = last(nodes, 'u');
  const underlineValue = underline ? attr(underline, 'val') !== 'none' : bool(nodes, 'u');
  const underlineStyle = attr(underline, 'val') || attr(nodes.at(-1), 'u');
  const strike = bool(nodes, 'strike', 'dstrike');
  return {
    bold: bool(nodes, 'b', 'bCs'), italic: bool(nodes, 'i', 'iCs'), underline: !!underlineValue, strike,
    css: cssStyle([
      ['font-family', font ? `'${font.replace(/[\\'\"\r\n]/g, '')}'` : undefined],
      ['font-size', fontSize], ['color', color], ['background-color', highlight], ['vertical-align', baselineValue],
      ['text-transform', caps], ['text-decoration-style', underlineValue && ['double', 'dbl'].includes(underlineStyle) ? 'double' : undefined],
      ['font-variant', caps === 'small-caps' ? 'small-caps' : undefined],
    ]),
  };
}
class WordStyles {
  private readonly styles = new Map<string, XmlNode>();
  constructor(archive: Archive) {
    const source = archive.source('word/styles.xml', false);
    if (source) for (const style of descendants(parseXml(source), 'style')) {
      const id = attr(style, 'styleId');
      if (id) this.styles.set(id, style);
    }
  }
  private chain(id: string, seen = new Set<string>()): XmlNode[] {
    const style = this.styles.get(id);
    if (!style || seen.has(id)) return [];
    seen.add(id);
    return [...this.chain(attr(child(style, 'basedOn'), 'val'), seen), style];
  }
  runProps(id: string): XmlNode[] { return this.chain(id).map(style => child(style, 'rPr')).filter((node): node is XmlNode => !!node); }
  paragraphProps(id: string): XmlNode[] { return this.chain(id).map(style => child(style, 'pPr')).filter((node): node is XmlNode => !!node); }
  isHeading(id: string): boolean { return this.chain(id).some(style => /^heading[1-6]$/i.test(attr(style, 'styleId'))); }
}
function paragraphStyle(nodes: XmlNode[], word: boolean): string {
  const direct = nodes.at(-1);
  const alignment = word ? lastValue(nodes, 'jc', 'val') : attr(direct, 'algn');
  const align: Record<string, string> = { left: 'left', l: 'left', center: 'center', ctr: 'center', right: 'right', r: 'right', both: 'justify', just: 'justify', distribute: 'justify' };
  const indent = last(nodes, 'ind');
  const spacing = last(nodes, 'spacing') ?? direct;
  const marginLeft = word ? points(attr(indent, 'left'), 20) : emuPoints(attr(nodes.length ? nodes[nodes.length - 1] : undefined, 'marL'));
  const marginRight = word ? points(attr(indent, 'right'), 20) : emuPoints(attr(nodes.length ? nodes[nodes.length - 1] : undefined, 'marR'));
  const firstLine = points(attr(indent, 'firstLine'), 20);
  const hanging = points(attr(indent, 'hanging'), 20);
  const indentValue = firstLine ? firstLine : hanging ? `-${hanging}` : undefined;
  const before = word ? points(attr(spacing, 'before'), 20) : spacingPoints(spacing, 'spcBef');
  const after = word ? points(attr(spacing, 'after'), 20) : spacingPoints(spacing, 'spcAft');
  const line = word ? lineHeight(attr(spacing, 'line'), attr(spacing, 'lineRule')) : presentationLine(spacing);
  return cssStyle([
    ['text-align', align[alignment]], ['margin-left', marginLeft], ['margin-right', marginRight], ['text-indent', indentValue],
    ['margin-top', before], ['margin-bottom', after], ['line-height', line],
  ]);
}
function spacingPoints(node: XmlNode | undefined, name: string): string | undefined {
  const spacing = child(node, name), pointsNode = child(spacing, 'spcPts');
  return points(attr(pointsNode, 'val'), 100);
}
function lineHeight(value: string, rule: string): string | undefined {
  if (!value) return undefined;
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return undefined;
  return rule === 'auto' || !rule ? String(number / 240) : `${number / 20}pt`;
}
function presentationLine(node: XmlNode | undefined): string | undefined {
  const line = child(node, 'lnSpc'), percent = child(line, 'spcPct'), pointsNode = child(line, 'spcPts');
  if (percent) return String(Number(attr(percent, 'val')) / 100000);
  return points(attr(pointsNode, 'val'), 100);
}
function* run(node: XmlNode, archive: Archive, relations: Map<string, string>, word: boolean, styles?: WordStyles, defaults: XmlNode[] = []): Generator<Token> {
  const direct = child(node, 'rPr');
  const styleID = attr(child(direct, 'rStyle'), 'val');
  const props = [...defaults, ...(styles?.runProps(styleID) ?? []), ...(direct ? [direct] : [])];
  const formatValue = format(props, word);
  let open = '', close = '';
  if (formatValue.css) { open += `<span style="${formatValue.css}">`; close = '</span>' + close; }
  if (formatValue.bold) { open += '<strong>'; close = '</strong>' + close; }
  if (formatValue.italic) { open += '<em>'; close = '</em>' + close; }
  if (formatValue.underline) { open += '<u>'; close = '</u>' + close; }
  if (formatValue.strike) { open += '<s>'; close = '</s>' + close; }
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
function* paragraph(node: XmlNode, archive: Archive, relations: Map<string, string>, word: boolean, styles?: WordStyles): Generator<Token> {
  const props = child(node, 'pPr');
  const styleID = attr(child(props, 'pStyle'), 'val');
  const styleProps = styles?.paragraphProps(styleID) ?? [];
  const paragraphProps = [...styleProps, ...(props ? [props] : [])];
  const heading = /^heading([1-6])$/i.exec(styleID) ?? (styles?.isHeading(styleID) ? /^heading([1-6])$/i.exec(styleID) : undefined);
  const element = heading ? `h${heading[1]}` : 'p';
  const list = word ? !!child(props, 'numPr') : !!child(props, 'buChar') || !!child(props, 'buAutoNum');
  const style = paragraphStyle(paragraphProps, word);
  const runDefaults = word ? [] : [child(props, 'defRPr')].filter((value): value is XmlNode => !!value);
  yield { html: `<${element}${style ? ` style="${style}"` : ''}${list ? ' class="list-item"' : ''}>`, close: `</${element}>` };
  if (list) yield { html: `${escape(attr(child(child(props, 'buChar'), 'char') || child(props, 'buChar'), 'char') || '•')} ` };
  let populated = false;
  for (const c of children(node)) {
    if (['r', 'fld'].includes(tag(c))) { yield* run(c, archive, relations, word, styles, runDefaults); populated = true; }
    else if (tag(c) === 'br') { yield { html: '<br />' }; populated = true; }
    else if (['hyperlink', 'ins', 'sdt'].includes(tag(c))) {
      for (const r of descendants([c], 'r')) { yield* run(r, archive, relations, word, styles, runDefaults); populated = true; }
    }
  }
  if (!populated) yield { html: '<br />' };
  yield { html: `</${element}>`, leave: true };
}
// Walk ranges in the XML source. Only a requested paragraph or picture becomes
// an object tree; large tables are traversed row by row, including nested ones.
function* tokens(source: string, archive: Archive, relations: Map<string, string>, word: boolean, styles?: WordStyles, depth = 0): Generator<Token> {
  for (const range of directElements(source)) {
    const name = range.name;
    if (name === 'p') yield* paragraph(parseXml(source.slice(range.start, range.end))[0]!, archive, relations, word, styles);
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
      yield* tokens(source.slice(range.head, range.tail), archive, relations, word, styles, depth+1);
      yield { html: close, leave: true };
    } else if (['document', 'body', 'sdt', 'sdtContent', 'sld', 'cSld', 'spTree', 'sp', 'txBody', 'graphicFrame', 'graphic', 'graphicData', 'grpSp'].includes(name)) {
      if (depth < 128) yield* tokens(source.slice(range.head, range.tail), archive, relations, word, styles, depth+1);
    }
  }
}
function* pages(source: string, archive: Archive, relations: Map<string, string>, word: boolean, styles?: WordStyles): Generator<string> {
  const stack: { html: string; close: string }[] = [];
  let html = '', size = 0;
  for (const token of tokens(source, archive, relations, word, styles)) {
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
  private readonly styles?: WordStyles;
  private iterator: Generator<string>;
  private generated = 0;
  private readonly cache = new Map<number, string>();
  private end?: number;
  index = 0;
  constructor(private readonly archive: Archive, part: string, private readonly word: boolean) {
    this.source = archive.source(part);
    this.relations = archive.relations(part);
    this.styles = word ? new WordStyles(archive) : undefined;
    this.iterator = pages(this.source, archive, this.relations, word, this.styles);
    if (word && !elements(this.source, 'body').next().value) throw new Error('Word 文件缺少正文');
  }
  private get(index: number): string | undefined {
    if (this.cache.has(index)) return this.cache.get(index);
    if (this.end !== undefined && index >= this.end) return undefined;
    if (index < this.generated) { this.iterator = pages(this.source, this.archive, this.relations, this.word, this.styles); this.generated = 0; }
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
  render(): string {
    const content = this.get(this.index);
    if (!this.word) return content ?? '';
    return `<div class="word-page"><div class="word-page-content">${content ?? '<p class="muted">此页面没有可预览的正文。</p>'}</div></div>`;
  }
}
