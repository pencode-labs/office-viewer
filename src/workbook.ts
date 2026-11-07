import { Archive, attr, child, children, descendants, tag, text } from './archive.js';
import type { XmlNode } from './archive.js';
import { escape } from './render.js';
export const PageRows = 50, PageColumns = 20;
export interface Sheet { name: string; part: string }
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
  readonly shared: string[];
  private readonly formats: number[];
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
    this.shared = Array.from(descendants(archive.xml('xl/sharedStrings.xml', false), 'si'), n => Array.from(descendants([n], 't'), text).join(''));
    const styles = archive.xml('xl/styles.xml', false);
    for (const n of descendants(styles, 'numFmt')) this.custom.set(Number(attr(n, 'numFmtId')), attr(n, 'formatCode'));
    const xfs = descendants(styles, 'cellXfs').next().value as XmlNode | undefined;
    this.formats = children(xfs).filter(n => tag(n) === 'xf').map(n => Number(attr(n, 'numFmtId')));
  }
  value(cell: XmlNode): string {
    const value = text(child(cell, 'v')), type = attr(cell, 't');
    if (type === 's') return this.shared[Number(value)] ?? '';
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
export class SheetData {
  readonly rows = new Map<number, Map<number, XmlNode>>();
  rowCount = 0;
  columnCount = 0;
  constructor(private readonly book: Workbook, readonly sheet: Sheet) {
    let previous = 0;
    const root = book.archive.xml(sheet.part);
    for (const row of descendants(root, 'row')) {
      const number = Number(attr(row, 'r')) || previous + 1;
      if (!Number.isInteger(number) || number < 1 || number > 1048576) continue;
      previous = number;
      const cells = new Map<number, XmlNode>();
      let column = 0;
      for (const cell of children(row).filter(n => tag(n) === 'c')) {
        const ref = attr(cell, 'r');
        column = ref ? columnNumber(ref) : column + 1;
        if (column < 1 || column > 16384) continue;
        cells.set(column, cell); this.columnCount = Math.max(this.columnCount, column);
      }
      if (cells.size) { this.rows.set(number, cells); this.rowCount = Math.max(this.rowCount, number); }
    }
  }
  render(rowPage: number, columnPage: number): string {
    if (!this.rows.size) return '<p class="muted">此工作表为空。</p>';
    const startRow = rowPage * PageRows + 1, endRow = Math.min(this.rowCount, startRow + PageRows - 1);
    const startColumn = columnPage * PageColumns + 1, endColumn = Math.min(this.columnCount, startColumn + PageColumns - 1);
    let html = '<table class="sheet"><thead><tr><th aria-label="行号"></th>';
    for (let c = startColumn; c <= endColumn; c++) html += `<th>${columnName(c)}</th>`;
    html += '</tr></thead><tbody>';
    for (let r = startRow; r <= endRow; r++) {
      html += `<tr><th>${r}</th>`;
      const row = this.rows.get(r);
      for (let c = startColumn; c <= endColumn; c++) {
        const cell = row?.get(c), value = cell ? this.book.value(cell) : '';
        const display = value.length > 512 ? value.slice(0, 512) + '…' : value;
        html += `<td>${escape(display)}</td>`;
      }
      html += '</tr>';
    }
    return html + '</tbody></table>';
  }
}
