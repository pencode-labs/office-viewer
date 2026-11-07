import { open } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { basename, extname } from 'node:path';
import { Archive, attr, descendants } from './archive.js';
import { escape, slidePage, wordPage } from './render.js';
import { PageColumns, PageRows, Workbook } from './workbook.js';
import type { SheetData } from './workbook.js';
import type { EditorMessage, EditorReply, Input } from '@pencode/peb';
const css = readFileSync(new URL('../resources/viewer.css', import.meta.url), 'utf8');

const MaxFile = 32 << 20;
export async function readOffice(path: string): Promise<Buffer> {
  const file = await open(path, 'r');
  try {
    const info = await file.stat();
    if (!info.isFile() || info.size > MaxFile) throw new Error('Office 预览只支持不超过 32 MiB 的普通文件');
    // Read at most the initial size plus one byte, even if another process grows it.
    const buffer = Buffer.allocUnsafe(info.size + 1);
    let length = 0;
    while (length < buffer.length) {
      const { bytesRead } = await file.read(buffer, length, buffer.length - length, null);
      if (!bytesRead) break;
      length += bytesRead;
    }
    if (length > info.size) throw new Error('文件正在变化，请刷新预览');
    return buffer.subarray(0, length);
  } finally { await file.close(); }
}
export class OfficeDocument {
  private readonly archive: Archive;
  private readonly book?: Workbook;
  private sheet?: SheetData;
  private readonly slides: string[] = [];
  private readonly word: string;
  private index = 0;
  private rowPage = 0;
  private columnPage = 0;
  constructor(readonly path: string, data: Uint8Array) {
    this.archive = new Archive(data);
    const extension = extname(path).toLowerCase();
    this.word = extension === '.docx' ? wordPage(this.archive) : '';
    if (extension === '.xlsx') { this.book = new Workbook(this.archive); this.sheet = this.book.open(0); }
    if (extension === '.pptx') {
      const relations = this.archive.relations('ppt/presentation.xml');
      for (const slide of descendants(this.archive.xml('ppt/presentation.xml'), 'sldId')) {
        const part = relations.get(attr(slide, 'r:id'));
        if (part) this.slides.push(part);
      }
      if (!this.slides.length) throw new Error('PowerPoint 文件没有可预览的幻灯片');
    }
    if (!['.docx', '.xlsx', '.pptx'].includes(extension)) throw new Error('请选择 .docx、.xlsx 或 .pptx 文件');
  }
  event(ref: string): boolean {
    if (ref.startsWith('sheet-') && this.book) {
      const index = Number(ref.slice(6));
      if (!Number.isInteger(index) || index < 0 || index >= this.book.sheets.length || index === this.index) return false;
      const sheet = this.book.open(index);
      this.index = index; this.sheet = sheet; this.rowPage = this.columnPage = 0; return true;
    }
    const oldIndex = this.index, oldRow = this.rowPage, oldColumn = this.columnPage;
    switch (ref) {
      case 'previous': this.index = Math.max(0, this.index - 1); break;
      case 'next': this.index = Math.min(this.slides.length - 1, this.index + 1); break;
      case 'rows-previous': this.rowPage = Math.max(0, this.rowPage - 1); break;
      case 'rows-next': if (this.sheet) this.rowPage = Math.min(Math.max(0, Math.ceil(this.sheet.rowCount / PageRows) - 1), this.rowPage + 1); break;
      case 'columns-previous': this.columnPage = Math.max(0, this.columnPage - 1); break;
      case 'columns-next': if (this.sheet) this.columnPage = Math.min(Math.max(0, Math.ceil(this.sheet.columnCount / PageColumns) - 1), this.columnPage + 1); break;
    }
    return oldIndex !== this.index || oldRow !== this.rowPage || oldColumn !== this.columnPage;
  }
  view(): Input<EditorReply> {
    const events: { ref: string; kind: string }[] = [];
    const button = (ref: string, label: string, disabled = false, active = false): string => {
      events.push({ ref, kind: 'click' });
      return `<button type="button" data-ref="${ref}"${disabled ? ' disabled' : ''}${active ? ' aria-pressed="true"' : ''}>${escape(label)}</button>`;
    };
    let navigation = '', body = this.word;
    if (this.book && this.sheet) {
      navigation = '<nav class="sheet-tabs" aria-label="工作表">' + this.book.sheets.map((s, i) => button(`sheet-${i}`, s.name, false, i === this.index)).join('') + '</nav>';
      navigation += '<div class="paging">' + button('rows-previous', '上一页', this.rowPage === 0) + button('rows-next', '下一页', (this.rowPage + 1) * PageRows >= this.sheet.rowCount);
      navigation += `<span>${this.sheet.rowCount} 行 · ${this.sheet.columnCount} 列 · 第 ${this.rowPage + 1} / ${Math.max(1, Math.ceil(this.sheet.rowCount / PageRows))} 页</span>`;
      navigation += button('columns-previous', '向左', this.columnPage === 0) + button('columns-next', '向右', (this.columnPage + 1) * PageColumns >= this.sheet.columnCount) + '</div>';
      body = this.sheet.render(this.rowPage, this.columnPage);
    } else if (this.slides.length) {
      navigation = '<nav class="paging" aria-label="幻灯片">' + button('previous', '上一张', this.index === 0) + `<span>幻灯片 ${this.index + 1} / ${this.slides.length}</span>` + button('next', '下一张', this.index === this.slides.length - 1) + '</nav>';
      body = slidePage(this.archive, this.slides[this.index]!);
    }
    const html = `<section class="office-viewer"><header><strong>${escape(basename(this.path))}</strong><span class="muted">只读预览</span>${button('reload', '刷新')}</header>${navigation}<article class="office-content${this.book ? ' workbook' : ''}${this.slides.length ? ' slide' : ''}" data-ref="office-content" data-pencode-selectable data-pencode-searchable aria-label="Office 预览">${body}</article></section>`;
    if (Buffer.byteLength(html) > 1800 << 10) throw new Error('当前页面内容超过预览限制（约 1.8 MiB），请拆分文档后重试');
    return { view: { html, css, events } };
  }
}
export function createViewer(): (message: EditorMessage) => Promise<Input<EditorReply>> {
  let document: OfficeDocument | undefined;
  return async message => {
    if (message.editor !== 'office') throw new Error('未知的 Office 编辑工具');
    if (message.event.kind === 'click' && message.event.ref !== 'reload' && document?.path === message.path) {
      return document.event(message.event.ref) ? document.view() : {};
    }
    if (!['open', 'documentchange'].includes(message.event.kind) && message.event.ref !== 'reload') return {};
    try {
      const next = new OfficeDocument(message.path, await readOffice(message.path));
      const reply = next.view(); document = next; return reply;
    } catch (error) {
      document = undefined;
      const detail = error instanceof Error ? error.message : String(error);
      return { view: { html: `<section class="office-viewer"><h2>无法预览 Office 文件</h2><p>${escape(detail)}</p><p class="muted">支持未加密的 .docx、.xlsx 和 .pptx。旧版 .doc、.xls、.ppt 请先转换为新格式。</p><button type="button" data-ref="reload">重试</button></section>`, css, events: [{ ref: 'reload', kind: 'click' }] } };
    }
  };
}
