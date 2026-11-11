import { Archive, attr, child, children, descendants, tag, text } from './archive.js';
import type { XmlNode } from './archive.js';
import { escape } from './render.js';
import { elements, parseXml, rangeAttribute } from './xml.js';
import type { XmlRange } from './xml.js';
export interface Sheet { name: string; part: string }
function columnPixels(value: string, fallback: number): number {
  const width = Number(value);
  return value && Number.isFinite(width) && width > 0 && width <= 255 ? Math.floor((256 * width + 18) / 256 * 7) : fallback;
}
function rowPixels(value: string, fallback: number): number {
  const height = Number(value);
  return value && Number.isFinite(height) && height > 0 && height <= 409 ? height * 4 / 3 : fallback;
}
export function columnNumber(reference: string): number {
  const match = /^([A-Z]{1,3})[1-9]\d{0,6}$/i.exec(reference);
  if (!match) return 0;
  let number = 0;
  for (const c of match[1]!.toUpperCase()) number = number * 26 + c.charCodeAt(0) - 64;
  return number <= 16384 ? number : 0;
}
export function columnName(number: number): string {
  let name = '';
  for (; number > 0; number = Math.floor((number - 1) / 26)) name = String.fromCharCode(65 + (number - 1) % 26) + name;
  return name;
}
export class Workbook {
  readonly sheets: Sheet[] = [];
  private readonly shared: string;
  private readonly sharedRanges: number[] = [];
  private readonly sharedCache = new Map<number, string>();
  private readonly formats: number[];
  readonly alignments: string[];
  private readonly custom = new Map<number, string>();
  private readonly date1904: boolean;
  constructor(readonly archive: Archive) {
    const root = archive.xml('xl/workbook.xml'), relations = archive.relations('xl/workbook.xml');
    const date = attr(descendants(root, 'workbookPr').next().value, 'date1904');
    this.date1904 = date === '1' || date === 'true';
    for (const sheet of descendants(root, 'sheet')) {
      if (attr(sheet, 'state') === 'hidden' || attr(sheet, 'state') === 'veryHidden') continue;
      const part = relations.get(attr(sheet, 'r:id'));
      if (part) this.sheets.push({ name: attr(sheet, 'name'), part });
    }
    if (!this.sheets.length) throw new Error('Excel 文件没有可预览的工作表');
    this.shared = archive.source('xl/sharedStrings.xml', false);
    for (const range of elements(this.shared, 'si')) this.sharedRanges.push(range.start, range.end);
    const styles = archive.xml('xl/styles.xml', false);
    for (const n of descendants(styles, 'numFmt')) this.custom.set(Number(attr(n, 'numFmtId')), attr(n, 'formatCode'));
    const xfs = descendants(styles, 'cellXfs').next().value as XmlNode | undefined;
    const cells = children(xfs).filter(n => tag(n) === 'xf');
    this.formats = cells.map(n => Number(attr(n, 'numFmtId')));
    this.alignments = cells.map(n => {
      const alignment = child(n, 'alignment');
      const horizontal = attr(alignment, 'horizontal'), vertical = attr(alignment, 'vertical');
      let css = '';
      if (['left', 'center', 'right', 'justify'].includes(horizontal)) css += `text-align:${horizontal};`;
      if (['top', 'center', 'bottom'].includes(vertical)) css += `vertical-align:${vertical === 'center' ? 'middle' : vertical};`;
      if (['1', 'true'].includes(attr(alignment, 'wrapText'))) css += 'white-space:pre-wrap;overflow-wrap:break-word;';
      return css;
    });
  }
  private sharedValue(index: number): string {
    if (!Number.isInteger(index) || index < 0 || index*2 >= this.sharedRanges.length) return '';
    const cached = this.sharedCache.get(index);
    if (cached !== undefined) return cached;
    const nodes = parseXml(this.shared.slice(this.sharedRanges[index*2]!, this.sharedRanges[index*2+1]!));
    const value = Array.from(descendants(nodes, 't'), text).join('');
    if (this.sharedCache.size >= 2048) this.sharedCache.delete(this.sharedCache.keys().next().value!);
    this.sharedCache.set(index, value); return value;
  }
  value(cell: XmlNode): string {
    const value = text(child(cell, 'v')), type = attr(cell, 't');
    if (type === 's') return this.sharedValue(Number(value));
    if (type === 'inlineStr') return Array.from(descendants(children(cell), 't'), text).join('');
    if (type === 'b') return value === '1' ? 'TRUE' : 'FALSE';
    if (type === 'str' || type === 'e' || type === 'd') return value;
    if (!value && child(cell, 'f')) return '[公式无缓存值]';
    const number = Number(value), format = this.formats[Number(attr(cell, 's'))] ?? 0;
    const code = (this.custom.get(format) ?? '').replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '');
    if (value && Number.isFinite(number)) {
      const date = format >= 14 && format <= 22 || format >= 45 && format <= 47 || /[ydhs]/i.test(code);
      if (date) {
        if (number === 60 && !this.date1904) return '1900-02-29';
        const epoch = this.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, number < 60 ? 31 : 30);
        const date = new Date(epoch + Math.round(number * 86400000));
        if (Number.isFinite(date.getTime())) {
          const timeOnly = format >= 18 && format <= 21 || format >= 45 && format <= 47 || (!/[yd]/i.test(code) && /[hs]/i.test(code));
          return timeOnly ? date.toISOString().slice(11, 19) : date.toISOString().slice(0, format === 22 || /[hs]/i.test(code) ? 19 : 10).replace('T', ' ');
        }
      }
      if (format === 9 || format === 10 || code.includes('%')) return `${(number * 100).toFixed(format === 9 ? 0 : 2)}%`;
    }
    return value;
  }
  open(index: number): SheetData { return new SheetData(this, this.sheets[index]!); }
}
export interface GridWindow {
  request: number; offset: number; firstRow: number; endRow: number;
  firstColumn: number; endColumn: number; all: boolean; metadata: boolean;
}
type Cell = [number, number, string, number];
let sourceID = 0;
export class SheetData {
  readonly source = `office-sheet-${++sourceID}`;
  private readonly xml: string;
  // Three integers per populated row, rather than a permanent XML object tree.
  private readonly rows: number[] = [];
  private readonly widths = new Map<number, number>();
  private readonly heights = new Map<number, number>();
  private readonly defaultWidth: number;
  private readonly defaultHeight: number;
  private transfer?: { request: number; offset: number; iterator: Generator<string>; pending: Buffer; used: number; done: boolean };
  readonly deferred: boolean;
  private readonly inline: Cell[] = [];
  rowCount = 0;
  columnCount = 0;
  constructor(private readonly book: Workbook, readonly sheet: Sheet) {
    const source = this.xml = book.archive.source(sheet.part);
    const format = elements(source, 'sheetFormatPr').next().value;
    this.defaultWidth = columnPixels(format ? rangeAttribute(source, format, 'defaultColWidth') : '', 64);
    this.defaultHeight = rowPixels(format ? rangeAttribute(source, format, 'defaultRowHeight') : '', 20);
    for (const col of elements(source, 'col')) {
      const min = Number(rangeAttribute(source, col, 'min')), max = Number(rangeAttribute(source, col, 'max'));
      if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min || max > 16384) continue;
      const width = columnPixels(rangeAttribute(source, col, 'width'), this.defaultWidth);
      for (let c = min; c <= max; c++) this.widths.set(c, width);
    }
    let previous = 0;
    const rows = new Map<number, [number, number]>();
    for (const row of elements(source, 'row')) {
      const number = Number(rangeAttribute(source, row, 'r')) || previous + 1;
      if (!Number.isInteger(number) || number < 1 || number > 1048576) continue;
      previous = number;
      const height = rangeAttribute(source, row, 'ht');
      if (height) this.heights.set(number, rowPixels(height, this.defaultHeight));
      let column = 0;
      for (const cell of elements(source, 'c', row.head, row.end)) {
        const ref = rangeAttribute(source, cell, 'r');
        column = ref ? columnNumber(ref) : column + 1;
        if (column > 0 && column <= 16384) this.columnCount = Math.max(this.columnCount, column);
      }
      rows.set(number, [row.start, row.end]);
      this.rowCount = Math.max(this.rowCount, number);
    }
    for (const [number, range] of [...rows].sort((a, b) => a[0] - b[0])) this.rows.push(number, ...range);
    // Inline only genuinely small payloads, including shared-string expansion.
    // Large sheets never format their offscreen cells during initial opening.
    let budget = book.alignments.reduce((n, css) => n+css.length*6+8, 0) + (this.widths.size+this.heights.size)*32;
    this.deferred = source.length > 64 << 10 || budget > 64 << 10;
    if (!this.deferred) {
      for (const cell of this.cells()) {
        budget += cell[2].length*6+64;
        if (budget > 64 << 10) { this.deferred = true; this.inline.length = 0; break; }
        this.inline.push(cell);
      }
    }
  }
  private *cells(window?: GridWindow): Generator<Cell> {
    let low = 0, high = this.rows.length / 3;
    if (window && !window.all) {
      while (low < high) {
        const mid = (low + high) >>> 1;
        if (this.rows[mid*3]! < window.firstRow) low = mid+1; else high = mid;
      }
    }
    for (let i = low*3; i < this.rows.length; i += 3) {
      const r = this.rows[i]!;
      if (window && !window.all && r >= window.endRow) break;
      let column = 0;
      const cells = new Map<number, XmlRange>();
      for (const cell of elements(this.xml, 'c', this.rows[i+1]!, this.rows[i+2]!)) {
        const ref = rangeAttribute(this.xml, cell, 'r');
        column = ref ? columnNumber(ref) : column+1;
        if (column < 1 || column > 16384 || window && !window.all && (column < window.firstColumn || column >= window.endColumn)) continue;
        cells.set(column, cell);
      }
      for (const [c, range] of [...cells].sort((a, b) => a[0]-b[0])) {
        const cell = parseXml(this.xml.slice(range.start, range.end))[0]!;
        const value = this.book.value(cell);
        if (value) yield [r, c, value, this.book.alignments[Number(attr(cell, 's'))] ? Number(attr(cell, 's'))+1 : 0];
      }
    }
  }
  private metadata() {
    return {
      styles: ['', ...this.book.alignments],
      rowSizes: Array.from(this.heights, ([r, size]): [number, number] => [r, Math.max(21, size)]).filter(([, size]) => size !== Math.max(21, this.defaultHeight)),
      columnSizes: Array.from(this.widths).filter(([, size]) => size !== this.defaultWidth),
    };
  }
  private *json(window: GridWindow): Generator<string> {
    yield '{';
    if (window.metadata) {
      const metadata = this.metadata();
      for (const [key, values] of Object.entries(metadata)) {
        yield `"${key}":[`;
        let separator = '';
        for (const value of values) { yield separator+JSON.stringify(value); separator = ','; }
        yield '],';
      }
    }
    yield '"cells":[';
    let separator = '';
    for (const cell of this.cells(window)) { yield separator+JSON.stringify(cell); separator = ','; }
    yield ']}';
  }
  chunk(window: GridWindow): string {
    for (const value of [window.request, window.offset, window.firstRow, window.endRow, window.firstColumn, window.endColumn]) {
      if (!Number.isSafeInteger(value) || value < 0) throw new Error('无效的表格窗口');
    }
    if (!window.offset) this.transfer = { request: window.request, offset: 0, iterator: this.json(window), pending: Buffer.alloc(0), used: 0, done: false };
    const transfer = this.transfer;
    if (!transfer || transfer.request !== window.request || transfer.offset !== window.offset) return '';
    const bytes = Buffer.allocUnsafe(64 << 10);
    let length = 0;
    while (length < bytes.length) {
      if (transfer.used === transfer.pending.length) {
        const next = transfer.iterator.next();
        if (next.done) { transfer.done = true; break; }
        transfer.pending = Buffer.from(next.value); transfer.used = 0;
      }
      const count = Math.min(bytes.length-length, transfer.pending.length-transfer.used);
      transfer.pending.copy(bytes, length, transfer.used, transfer.used+count);
      transfer.used += count; length += count;
    }
    const payload = { source: this.source, request: window.request, offset: window.offset, done: transfer.done, data: bytes.subarray(0, length).toString('base64') };
    transfer.offset += length;
    if (transfer.done) this.transfer = undefined;
    return `<div data-pencode-grid-chunk="${escape(JSON.stringify(payload))}"></div>`;
  }
  render(): string {
    const grid = {
      rows: Math.max(100, this.rowCount), columns: Math.max(26, this.columnCount),
      rowHeight: Math.max(21, this.defaultHeight), columnWidth: this.defaultWidth,
      ...this.deferred ? { source: this.source, styles: [''], cells: [] } : { ...this.metadata(), cells: this.inline },
    };
    return `<div class="sheet-grid" data-ref="office-grid" data-pencode-grid="${escape(JSON.stringify(grid))}"></div>`;
  }
}
