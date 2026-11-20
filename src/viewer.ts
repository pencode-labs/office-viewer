import { readFileSync } from 'node:fs';
import { extname } from 'node:path';
import { MaxDocument } from '@pencode/peb';
import type { SlidePreview } from './slide-preview.js';
import { Archive, attr, descendants, text } from './archive.js';
import { escape } from './render.js';
import { PagedPreview } from './pages.js';
import { Workbook } from './workbook.js';
import { elements, parseXml } from './xml.js';
import type { SheetData, GridWindow } from './workbook.js';
import type { EditorMessage, EditorReply, Input } from '@pencode/peb';
const css = readFileSync(new URL('../resources/viewer.css', import.meta.url), 'utf8');
function iconButton(ref: string, label: string, icon: string, disabled = false): string {
  return `<button class="icon-button" type="button" data-ref="${ref}" title="${label}" aria-label="${label}"${disabled ? ' disabled' : ''}><span class="office-icon ${icon}" aria-hidden="true"></span></button>`;
}

function slideSummary(source: string): string {
  let summary = '';
  let remaining = 100;
  for (const range of elements(source, 't')) {
    const value = Array.from(text(parseXml(source.slice(range.start, range.end))[0]).trim().slice(0, remaining * 2)).slice(0, remaining);
    if (!value.length) continue;
    summary += `<span>${escape(value.join(''))}</span>`;
    remaining -= value.length;
    if (!remaining) break;
  }
  return summary;
}

export class OfficeDocument {
  private readonly archive: Archive;
  private readonly book?: Workbook;
  private sheet?: SheetData;
  private readonly slides: string[] = [];
  private preview?: PagedPreview;
  private word?: Awaited<ReturnType<typeof import('./word-preview.js').renderWord>>;
  private presentation?: SlidePreview;
  private thumbnails?: string[];
  private slideWidth = 960;
  private slideHeight = 540;
  private slideText = false;
  private index = 0;
  constructor(readonly path: string, data: Uint8Array | string = path) {
    this.archive = new Archive(data);
    const extension = extname(path).toLowerCase();
    try {
      if (extension === '.docx') this.archive.source('word/document.xml');
      if (extension === '.xlsx') { this.book = new Workbook(this.archive); this.sheet = this.book.open(0); }
      if (extension === '.pptx') {
        const presentation = this.archive.xml('ppt/presentation.xml');
        const size = descendants(presentation, 'sldSz').next().value;
        const width = Number(attr(size, 'cx')), height = Number(attr(size, 'cy'));
        if (Number.isSafeInteger(width) && Number.isSafeInteger(height) && width > 0 && height > 0) {
          this.slideWidth = width;
          this.slideHeight = height;
        }
        const relations = this.archive.relations('ppt/presentation.xml');
        for (const slide of descendants(presentation, 'sldId')) {
          const part = relations.get(attr(slide, 'r:id'));
          if (part) this.slides.push(part);
        }
        if (!this.slides.length) throw new Error('PowerPoint 文件没有可预览的幻灯片');
      }
      if (!['.docx', '.xlsx', '.pptx'].includes(extension)) throw new Error('请选择 .docx、.xlsx 或 .pptx 文件');
    } catch (error) { this.archive.close(); throw error; }
  }
  static async open(path: string, data: Uint8Array | string = path): Promise<OfficeDocument> {
    const document = new OfficeDocument(path, data);
    try {
      const extension = extname(path).toLowerCase();
      if (extension === '.docx') {
        const { renderWord } = await import('./word-preview.js');
        document.word = await renderWord(document.archive.rendererPackage('word'));
      } else if (extension === '.pptx') {
        const { SlidePreview } = await import('./slide-preview.js');
        document.presentation = new SlidePreview(document.archive.rendererPackage('ppt'), document.slideWidth, document.slideHeight);
        await document.prepareSlide();
      }
    } catch {
      // Keep bounded text navigation available for files the layout renderer cannot open.
    }
    return document;
  }

  async prepareSlide(): Promise<void> {
    if (!this.presentation || this.slideText) return;
    try { await this.presentation.show(this.index, this.slides.length); }
    catch { this.presentation = undefined; }
  }

  close(): void { this.archive.close(); }
  chunk(ref: string, value: string): Input<EditorReply> {
    if (this.sheet?.source !== ref) return {};
    const html = this.sheet.chunk(JSON.parse(value) as GridWindow);
    return html ? { view: { html } } : {};
  }
  event(ref: string): boolean {
    if (ref.startsWith('sheet-') && this.book) {
      const index = Number(ref.slice(6));
      if (!Number.isInteger(index) || index < 0 || index >= this.book.sheets.length || index === this.index) return false;
      const sheet = this.book.open(index);
      this.index = index; this.sheet = sheet; return true;
    }
    if (ref === 'page-previous') return this.preview?.move(-1) ?? false;
    if (ref === 'page-next') return this.preview?.move(1) ?? false;
    if (!this.slides.length) return false;
    if (ref === 'slide-text') {
      this.slideText = !this.slideText;
      this.preview = undefined;
      return true;
    }
    let index = this.index;
    if (ref === 'previous') index--;
    else if (ref === 'next') index++;
    else if (/^slide-\d+$/.test(ref)) index = Number(ref.slice(6));
    if (index < 0 || index >= this.slides.length || index === this.index) return false;
    this.index = index;
    this.preview = undefined;
    return true;
  }

  private slideThumbnails(): string[] {
    return this.thumbnails ??= this.slides.map(part => {
      try { return slideSummary(this.archive.source(part)); }
      catch { return ''; }
    });
  }

  view(): Input<EditorReply> {
    const events: { ref: string; kind: string }[] = [];
    const button = (ref: string, label: string, icon: string, disabled = false): string => {
      events.push({ ref, kind: 'click' });
      return iconButton(ref, label, icon, disabled);
    };
    let navigation = '', sidebar = '', body = '';
    const kind = this.book ? 'excel' : this.slides.length ? 'powerpoint' : 'word';
    let contentClass = 'word';
    if (this.book && this.sheet) {
      navigation = '<nav class="sheet-tabs" aria-label="工作表">' + this.book.sheets.map((s, i) => {
        events.push({ ref: `sheet-${i}`, kind: 'click' });
        return `<button type="button" data-ref="sheet-${i}" title="${escape(s.name)}" aria-pressed="${i === this.index}">${escape(s.name)}</button>`;
      }).join('') + '</nav>';
      body = this.sheet.render();
      contentClass = 'workbook';
    } else if (this.slides.length) {
      const image = this.presentation?.image(this.index);
      const textMode = this.slideText || !image;
      contentClass = textMode ? 'slide-text' : 'slide';
      const ratio = this.slideHeight / this.slideWidth;
      const thumbnailHeight = Math.round(Math.max(28, Math.min(144, 136 * ratio)));
      const summaries = this.slideThumbnails();
      let budget = Math.min(320 << 10, MaxDocument - (image?.length ?? 0) - 320 * summaries.length - 4096);
      sidebar = '<nav class="presentation-list" data-ref="slide-list" data-pencode-scroll="preserve" aria-label="幻灯片缩略图">' + summaries.map((summary, index) => {
        const ref = `slide-${index}`, label = `第 ${index + 1} 张幻灯片`;
        events.push({ ref, kind: 'click' });
        const thumbnail = this.presentation?.thumbnail(index);
        const hasImage = thumbnail && thumbnail.length <= budget;
        const preview = hasImage ? `<img src="${thumbnail}" alt="">` : summary;
        const size = Buffer.byteLength(preview);
        const contents = size <= budget ? preview : '';
        if (contents) budget -= size;
        return `<button class="presentation-entry" type="button" data-ref="${ref}" aria-label="${label}" aria-pressed="${index === this.index}"><span class="slide-number" aria-hidden="true">${index + 1}</span><span class="presentation-thumbnail" style="height:${thumbnailHeight}px" aria-hidden="true">${contents}</span></button>`;
      }).join('') + '</nav>';
      if (image || this.slideText) {
        events.push({ ref: 'slide-text', kind: 'click' });
        navigation = `<button class="slide-mode" type="button" data-ref="slide-text" aria-pressed="${textMode}">文本视图</button>`;
      }
      if (textMode) {
        this.preview ??= new PagedPreview(this.archive, this.slides[this.index]!, false);
        body = `<div class="presentation-text">${this.preview.render()}</div>`;
      } else {
        this.preview = undefined;
        body = `<img class="presentation-image" src="${image}" alt="第 ${this.index + 1} 张幻灯片">`;
      }
    } else if (this.word) {
      body = this.word.html;
    } else {
      this.preview ??= new PagedPreview(this.archive, 'word/document.xml', true);
      body = `<div class="word-fallback">${this.preview.render()}</div>`;
    }
    if (this.preview && (this.preview.index || this.preview.next)) {
      navigation += '<nav class="paging" aria-label="内容分段">' + button('page-previous', '上一段', 'left', this.preview.index === 0) + `<span>第 ${this.preview.index+1} 段</span>` + button('page-next', '下一段', 'right', !this.preview.next) + '</nav>';
    }
    const article = `<article class="office-content ${contentClass}" data-ref="office-content" data-pencode-selectable data-pencode-searchable data-pencode-preserve-whitespace aria-label="Office 预览">${body}</article>`;
    let content = article;
    if (sidebar) {
      const controls = navigation ? `<footer class="presentation-controls">${navigation}</footer>` : '';
      content = `<div class="presentation-workspace"><aside class="presentation-sidebar">${sidebar}${controls}</aside><div class="presentation-stage">${article}</div></div>`;
    } else if (navigation) {
      content = `<header class="office-toolbar"><div class="office-toolbar-navigation">${navigation}</div></header>` + article;
    }
    return { view: { html: `<section class="office-viewer ${kind}">${content}</section>`, css: this.word ? this.word.css + css : css, events } };
  }
}
export function createViewer(signal: AbortSignal): (message: EditorMessage) => Promise<Input<EditorReply>> {
  signal.throwIfAborted();
  let document: OfficeDocument | undefined;
  signal.addEventListener('abort', () => { document?.close(); document = undefined; }, { once: true });
  return async message => {
    signal.throwIfAborted();
    if (message.editor !== 'office') throw new Error('未知的 Office 编辑工具');
    try {
      if (message.event.kind === 'grid' && document?.path === message.path) return document.chunk(message.event.ref, message.event.value);
      if (message.event.kind === 'click' && message.event.ref !== 'reload' && document?.path === message.path) {
        if (!document.event(message.event.ref)) return {};
        await document.prepareSlide();
        return document.view();
      }
      if (!['open', 'documentchange'].includes(message.event.kind) && message.event.ref !== 'reload') return {};
      document?.close(); document = undefined;
      const next = await OfficeDocument.open(message.path);
      try { const reply = next.view(); document = next; return reply; }
      catch (error) { next.close(); throw error; }
    } catch (error) {
      document?.close(); document = undefined;
      const detail = error instanceof Error ? error.message : String(error);
      return { view: { html: `<section class="office-viewer"><header class="office-toolbar"><div class="office-actions"><button type="button" data-ref="reload">重试</button></div></header><article class="office-content office-error"><h2>无法预览</h2><p>${escape(detail)}</p></article></section>`, css, events: [{ ref: 'reload', kind: 'click' }] } };
    }
  };
}
