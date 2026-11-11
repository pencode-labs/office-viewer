import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { zipSync, strToU8 } from 'fflate';
import { decode, encode, EditorReplyCodec, EditorMessageCodec, HelloCodec, Kind, Reader, Writer } from '@pencode/peb';
import { Archive } from '../src/archive.js';
import { OfficeDocument } from '../src/viewer.js';
import type { GridWindow } from '../src/workbook.js';

const zip = (files: Record<string, string>) => zipSync(Object.fromEntries(Object.entries(files).map(([name, data]) => [name, strToU8(data)])));
const html = (document: OfficeDocument) => document.view().view!.html!;
interface GridPayload {
  source: string; rows: number; columns: number; styles: string[];
  columnSizes: [number, number][]; cells: [number, number, string, number][];
}
interface GridChunk { source: string; request: number; offset: number; done: boolean; data: string }
function attribute<T>(view: string, name: string): T {
  const value = view.match(new RegExp(`${name}="([^"]*)"`))![1]!;
  return JSON.parse(value.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&'));
}
const workbook = (sheet: string, shared?: string) => zip({
  'xl/workbook.xml': '<workbook><sheets><sheet name="large" r:id="r1"/><sheet name="other" r:id="r2"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="r1" Target="worksheets/data.xml"/><Relationship Id="r2" Target="worksheets/other.xml"/></Relationships>',
  'xl/styles.xml': '<styleSheet><cellXfs><xf/><xf><alignment horizontal="center"/></xf></cellXfs></styleSheet>',
  'xl/worksheets/data.xml': sheet,
  'xl/worksheets/other.xml': '<worksheet><sheetData/></worksheet>',
  ...shared === undefined ? {} : { 'xl/sharedStrings.xml': shared },
});
function fetchGrid(document: OfficeDocument, source: string, options: Partial<GridWindow> = {}) {
  const request = { request: 1, offset: 0, firstRow: 1, endRow: 12, firstColumn: 1, endColumn: 3, all: false, metadata: true, ...options };
  const chunks: Buffer[] = [];
  for (;;) {
    const reply = document.chunk(source, JSON.stringify(request));
    assert.ok(encode(EditorReplyCodec, reply).length < 128 << 10);
    const chunk = attribute<GridChunk>(reply.view!.html!, 'data-pencode-grid-chunk');
    assert.equal(chunk.request, request.request); assert.equal(chunk.offset, request.offset);
    const data = Buffer.from(chunk.data, 'base64');
    assert.ok(data.length <= 64 << 10); chunks.push(data);
    if (chunk.done) break;
    assert.ok(data.length > 0); request.offset += data.length;
  }
  const bytes = Buffer.concat(chunks);
  return { data: JSON.parse(bytes.toString('utf8')) as GridPayload, replies: chunks.length, bytes: bytes.length };
}
test('large Excel fetches both axes on demand and streams full search/copy data beyond 1.8 MiB', () => {
  const sheet = '<worksheet><cols><col min="2" max="2" width="40"/></cols><sheetData>' + Array.from({ length: 6000 }, (_, i) => `<row r="${i+1}" ht="30"><c r="A${i+1}" t="inlineStr" s="1"><is><t>行${i+1} ${'中文🙂 &amp;&lt;&quot;'.repeat(24)}</t></is></c><c r="AD${i+1}"><v>${i+1}</v></c></row>`).join('') + '</sheetData></worksheet>';
  const document = new OfficeDocument('/large.xlsx', workbook(sheet));
  const initial = html(document), data = attribute<GridPayload>(initial, 'data-pencode-grid');
  assert.ok(initial.length < 3000); assert.equal(data.rows, 6000); assert.equal(data.columns, 30);
  assert.deepEqual(data.cells, []);
  const visible = fetchGrid(document, data.source);
  assert.equal(visible.data.cells.length, 11);
  assert.ok(visible.data.cells.every((c) => c[0] < 12 && c[1] === 1));
  assert.deepEqual(visible.data.columnSizes, [[2, 280]]);
  assert.equal(visible.data.styles[visible.data.cells[0]![3]], 'text-align:center;');
  const distant = fetchGrid(document, data.source, { request: 2, firstRow: 5999, endRow: 6001, firstColumn: 30, endColumn: 31, metadata: false });
  assert.deepEqual(distant.data.cells, [[5999, 30, '5999', 0], [6000, 30, '6000', 0]]);
  const full = fetchGrid(document, data.source, { request: 3, all: true });
  assert.equal(full.data.cells.length, 12000);
  assert.ok(full.bytes > 1800 << 10); assert.ok(full.replies > 30);
  assert.match(full.data.cells.at(-2)![2], /^行6000 中文🙂 &<"/);
  assert.equal(document.event('sheet-1'), true);
  assert.deepEqual(document.chunk(data.source, '{}'), {}, 'ignore stale sheet requests');
});
test('small sheet with an oversized shared string also defers its transfer', () => {
  const value = '🙂中文&'.repeat(100000);
  const sheet = '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>';
  const document = new OfficeDocument('/shared.xlsx', workbook(sheet, `<sst><si><t>${value.replaceAll('&', '&amp;')}</t></si></sst>`));
  const first = attribute<GridPayload>(html(document), 'data-pencode-grid');
  assert.deepEqual(first.cells, []);
  assert.equal(fetchGrid(document, first.source).data.cells[0]![2], value);
});
test('Word and PowerPoint render bounded balanced segments with backward navigation', () => {
  const paragraphs = Array.from({ length: 4500 }, (_, i) => `<w:p><w:r><w:t>${String(i).padStart(5, '0')} ${'长文本🙂&amp;&lt;'.repeat(24)}</w:t></w:r></w:p>`);
  const docx = zip({ 'word/document.xml': `<w:document><w:body><w:tbl>${paragraphs.map(p => `<w:tr><w:tc>${p}</w:tc></w:tr>`).join('')}</w:tbl></w:body></w:document>` });
  const pptx = zip({
    'ppt/presentation.xml': '<presentation><sldIdLst><sldId r:id="r1"/></sldIdLst></presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="r1" Target="slides/s.xml"/></Relationships>',
    'ppt/slides/s.xml': `<sld><cSld><spTree><sp><txBody>${paragraphs.join('')}</txBody></sp></spTree></cSld></sld>`,
  });
  for (const [path, data] of [['large.docx', docx], ['large.pptx', pptx]] as const) {
    const document = new OfficeDocument(path, data), first = html(document);
    assert.match(first, /00000/); assert.doesNotMatch(first, /04499/);
    let total = 0, count = 0;
    do {
      const reply = document.view(), page = reply.view!.html!;
      assert.ok(encode(EditorReplyCodec, reply).length < 1 << 20);
      for (const name of ['table', 'tr', 'td', 'p']) {
        assert.equal((page.match(new RegExp(`<${name}>`, 'g')) ?? []).length, (page.match(new RegExp(`</${name}>`, 'g')) ?? []).length);
      }
      total += Buffer.byteLength(page); count++;
    } while (document.event('page-next'));
    assert.ok(total > 1800 << 10); assert.ok(count > 5); assert.match(html(document), /04499/);
    while (document.event('page-previous')) { /* revisit earlier source ranges */ }
    assert.equal(html(document), first);
  }
});
test('an oversized paragraph splits without dropping text or surrogate pairs', () => {
  const value = '🙂<&'.repeat(100000);
  const document = new OfficeDocument('paragraph.docx', zip({ 'word/document.xml': `<document><body><p><r><rPr><b/></rPr><t>${value.replaceAll('&', '&amp;').replaceAll('<', '&lt;')}</t></r></p></body></document>` }));
  let collected = '';
  do {
    const page = html(document);
    assert.ok(Buffer.byteLength(page) < 1 << 20);
    collected += Array.from(page.matchAll(/<strong>(.*?)<\/strong>/gs), m => m[1]).join('');
    assert.doesNotMatch(page, /\uFFFD/);
  } while (document.event('page-next'));
  assert.equal(collected.replaceAll('&lt;', '<').replaceAll('&amp;', '&'), value);
});

test('archive reads and validates members only when they are requested', async () => {
  const value = zipSync({ 'word/document.xml': strToU8('<document><body><p><r><t>visible</t></r></p></body></document>'), 'unused.xml': strToU8('unrequested-member') }, { level: 0 });
  const bytes = Buffer.from(value), position = bytes.indexOf('unrequested-member');
  assert.ok(position > 0); bytes[position] = bytes[position]! ^ 1;
  const directory = await mkdtemp(join(tmpdir(), 'office-lazy-'));
  const path = join(directory, 'lazy.docx');
  try {
    await writeFile(path, bytes);
    const document = new OfficeDocument(path);
    assert.match(html(document), /visible/); document.close();
    const archive = new Archive(path);
    try { assert.throws(() => archive.source('unused.xml'), /ZIP 校验/); }
    finally { archive.close(); }
    assert.throws(() => archive.source('word/document.xml'), /已关闭/);
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test('actual Node PEB process streams an oversized cell without enlarging protocol limits', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'office-stream-'));
  const child = spawn(process.execPath, ['dist/office-viewer.mjs'], { stdio: ['pipe', 'pipe', 'inherit'] });
  const exited = new Promise<number | null>((resolve, reject) => { child.once('exit', resolve); child.once('error', reject); });
  const reader = new Reader(child.stdout), writer = new Writer(child.stdin);
  try {
    await writer.send(Kind.Initialize, 1, encode(HelloCodec, { host: 'pencode', capabilities: 511 }));
    assert.equal((await reader.read())?.kind, Kind.Register);
    const value = '中文🙂'.repeat(240000), path = join(directory, 'large.xlsx');
    await writeFile(path, workbook('<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>', `<sst><si><t>${value}</t></si></sst>`));
    let id = 2;
    const exchange = async (kind: string, ref = '', eventValue = '') => {
      await writer.send(Kind.EditorRequest, id, encode(EditorMessageCodec, { editor: 'office', path, binary: true, event: { kind, ref, value: eventValue } }));
      const frame = await reader.read();
      assert.equal(frame?.kind, Kind.EditorResponse); assert.equal(frame?.id, id++);
      assert.ok(frame!.payload.length < 128 << 10);
      return decode(EditorReplyCodec, frame!.payload);
    };
    const first = await exchange('open'), source = attribute<GridPayload>(first.view!.html, 'data-pencode-grid').source;
    const chunks: Buffer[] = [];
    let offset = 0;
    for (;;) {
      const reply = await exchange('grid', source, JSON.stringify({ request: 1, offset, firstRow: 1, endRow: 2, firstColumn: 1, endColumn: 2, all: true, metadata: true }));
      const chunk = attribute<GridChunk>(reply.view!.html, 'data-pencode-grid-chunk'), bytes = Buffer.from(chunk.data, 'base64');
      assert.equal(chunk.offset, offset); chunks.push(bytes); offset += bytes.length;
      if (chunk.done) break;
    }
    assert.ok(offset > 1800 << 10);
    assert.equal(JSON.parse(Buffer.concat(chunks).toString('utf8')).cells[0][2], value);
    await writer.send(Kind.Shutdown, 0); assert.equal(await exited, 0);
  } finally { child.kill(); await rm(directory, { recursive: true, force: true }); }
});

test('an empty self-closing Word body remains previewable', () => {
  const document = new OfficeDocument('empty.docx', zip({ 'word/document.xml': '<document><body/></document>' }));
  assert.match(html(document), /没有可预览的正文/);
});
