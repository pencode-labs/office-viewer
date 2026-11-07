import { Archive, attr, child, children, descendants, tag, text } from './archive.js';
import type { XmlNode } from './archive.js';
export function escape(value: string): string {
  return value.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
}
export function enabled(node: XmlNode | undefined): boolean { return !!node && !['0', 'false', 'off'].includes(attr(node, 'val')); }
function image(node: XmlNode, archive: Archive, relations: Map<string, string>): string {
  const blip = descendants([node], 'blip').next().value as XmlNode | undefined;
  const src = archive.image(relations.get(attr(blip, 'embed')));
  return src ? `<img src="${src}" alt="文档图片" />` : '<span class="muted">[图片不可预览]</span>';
}
function run(node: XmlNode, archive: Archive, relations: Map<string, string>): string {
  const props = child(node, 'rPr');
  let body = '';
  for (const c of children(node)) {
    switch (tag(c)) {
      case 't': body += escape(text(c)); break;
      case 'br': case 'cr': body += '<br />'; break;
      case 'tab': body += '&#160;&#160;&#160;&#160;'; break;
      case 'drawing': case 'pict': body += image(c, archive, relations); break;
    }
  }
  if (enabled(child(props, 'b'))) body = `<strong>${body}</strong>`;
  if (enabled(child(props, 'i'))) body = `<em>${body}</em>`;
  const u = child(props, 'u');
  if (enabled(u) && attr(u, 'val') !== 'none') body = `<u>${body}</u>`;
  if (enabled(child(props, 'strike'))) body = `<s>${body}</s>`;
  return body;
}
export function paragraph(node: XmlNode, archive: Archive, relations: Map<string, string>, word = true): string {
  const props = child(node, 'pPr');
  const style = attr(child(props, 'pStyle'), 'val');
  const heading = /^heading([1-6])$/i.exec(style);
  const element = heading ? `h${heading[1]}` : 'p';
  const alignment = attr(child(props, word ? 'jc' : 'pPr'), word ? 'val' : 'algn') || attr(props, 'algn');
  const align: Record<string, string> = { left: 'left', l: 'left', center: 'center', ctr: 'center', right: 'right', r: 'right', both: 'justify', just: 'justify' };
  const list = word ? child(props, 'numPr') !== undefined : child(props, 'buChar') !== undefined || child(props, 'buAutoNum') !== undefined;
  let content = '';
  for (const c of children(node)) {
    switch (tag(c)) {
      case 'r': case 'fld': content += run(c, archive, relations); break;
      case 'br': content += '<br />'; break;
      case 'hyperlink': case 'ins': case 'sdt':
        for (const r of descendants([c], 'r')) content += run(r, archive, relations);
        break;
    }
  }
  const inline = align[alignment] ? ` style="text-align:${align[alignment]}"` : '';
  return `<${element}${inline}${list ? ' class="list-item"' : ''}>${list ? '• ' : ''}${content || '<br />'}</${element}>`;
}
export function table(node: XmlNode, archive: Archive, relations: Map<string, string>, word = true, depth = 0): string {
  if (depth > 8) return '<p>[嵌套表格超过预览层数]</p>';
  let html = '<table><tbody>';
  for (const row of children(node).filter(n => tag(n) === 'tr')) {
    html += '<tr>';
    for (const cell of children(row).filter(n => tag(n) === 'tc')) {
      const span = Number(attr(child(child(cell, 'tcPr'), 'gridSpan'), 'val'));
      html += `<td${span > 1 && span <= 256 ? ` colspan="${span}"` : ''}>`;
      const contents = word ? children(cell) : children(child(cell, 'txBody'));
      for (const c of contents) {
        if (tag(c) === 'p') html += paragraph(c, archive, relations, word);
        if (tag(c) === 'tbl') html += table(c, archive, relations, word, depth + 1);
      }
      html += '</td>';
    }
    html += '</tr>';
  }
  return html + '</tbody></table>';
}
export function wordPage(archive: Archive): string {
  const part = 'word/document.xml', relations = archive.relations(part);
  const body = descendants(archive.xml(part), 'body').next().value as XmlNode | undefined;
  if (!body) throw new Error('Word 文件缺少正文');
  let html = '';
  for (const c of children(body)) {
    if (tag(c) === 'p') html += paragraph(c, archive, relations);
    if (tag(c) === 'tbl') html += table(c, archive, relations);
    if (tag(c) === 'sdt') {
      for (const p of descendants([c], 'p')) html += paragraph(p, archive, relations);
    }
  }
  return html || '<p class="muted">此文档没有可预览的正文。</p>';
}
export function slidePage(archive: Archive, part: string): string {
  const relations = archive.relations(part);
  let html = '';
  const tree = descendants(archive.xml(part), 'spTree').next().value as XmlNode | undefined;
  function shapes(nodes: XmlNode[]): void {
    for (const shape of nodes) {
      switch (tag(shape)) {
        case 'sp':
          for (const p of descendants([shape], 'p')) html += paragraph(p, archive, relations, false);
          break;
        case 'pic': html += image(shape, archive, relations); break;
        case 'graphicFrame':
          for (const t of descendants([shape], 'tbl')) html += table(t, archive, relations, false);
          break;
        case 'grpSp': shapes(children(shape)); break;
      }
    }
  }
  shapes(children(tree));
  return html || '<p class="muted">此幻灯片没有可预览的文本或图片。</p>';
}
