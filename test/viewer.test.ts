import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import test from 'node:test';
import { zipSync, strToU8 } from 'fflate';
import { decode, encode, EditorMessageCodec, EditorReplyCodec, HelloCodec, ContributionsCodec, Kind, Reader, Writer } from '@pencode/peb';
import { createViewer, OfficeDocument } from '../src/viewer.js';
import { columnName, columnNumber } from '../src/workbook.js';
function zip(files: Record<string, string>): Uint8Array { return zipSync(Object.fromEntries(Object.entries(files).map(([name, data]) => [name, strToU8(data)]))); }
const word = (value: string) => zip({ 'word/document.xml': `<w:document xmlns:w="word"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${value}</w:t></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t> plain</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格内容</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>` });
function html(document: OfficeDocument): string { return document.view().view?.html ?? ''; }
interface Grid { rows: number; columns: number; rowHeight: number; columnWidth: number; rowSizes: number[][]; columnSizes: number[][]; styles: string[]; cells: [number, number, string, number][] }
function grid(view: string): Grid {
  const payload = view.match(/data-pencode-grid="([^"]*)"/)![1]!;
  return JSON.parse(payload.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')) as Grid;
}
const message = (path: string, kind = 'open', ref = '') => decode(EditorMessageCodec, encode(EditorMessageCodec, { editor: 'office', path, binary: true, event: { kind, ref } }));
test('Word preserves inline order, formatting, tables and escapes file text', () => {
  const document = new OfficeDocument('/测试.docx', word('你好 &lt;script&gt;'));
  const view = document.view();
  assert.match(view.view!.html!, /<h1><strong>你好 &lt;script&gt;<\/strong> plain<\/h1>/);
  assert.match(view.view!.html!, /<td><p>表格内容<\/p><\/td>/);
  assert.equal(view.content, undefined);
  assert.doesNotMatch(view.view!.html!, /<script>/);
  assert.match(view.view!.html!, /data-pencode-searchable/);
});
test('Word styles and PowerPoint run properties render as text layout styles', () => {
  const docx = zip({
    'word/styles.xml': '<w:styles><w:style w:styleId="BodyText"><w:rPr><w:rFonts w:ascii="仿宋"/><w:sz w:val="28"/><w:color w:val="C00000"/></w:rPr></w:style></w:styles>',
    'word/document.xml': '<w:document><w:body><w:p><w:pPr><w:pStyle w:val="BodyText"/><w:jc w:val="center"/><w:ind w:left="720" w:firstLine="360"/><w:spacing w:before="240" w:after="120" w:line="360"/></w:pPr><w:r><w:rPr><w:rStyle w:val="BodyText"/><w:b/><w:u w:val="double"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>样式文字</w:t></w:r></w:p></w:body></w:document>',
  });
  const wordView = html(new OfficeDocument('/styled.docx', docx));
  assert.match(wordView, /font-family:'仿宋';font-size:14pt;color:#c00000;/);
  assert.match(wordView, /font-weight|<strong>/);
  assert.match(wordView, /text-decoration-style:double/);
  assert.match(wordView, /vertical-align:super/);
  assert.match(wordView, /text-align:center/);
  assert.match(wordView, /margin-left:36pt/);
  const pptx = zip({
    'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId r:id="one"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="one" Target="slides/slide.xml"/></Relationships>',
    'ppt/slides/slide.xml': '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1" i="1" u="dbl" sz="2800" baseline="30000"><a:latin typeface="Aptos"/><a:solidFill><a:srgbClr val="00A2FF"/></a:solidFill></a:rPr><a:t>PPT 样式</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  });
  const pptView = html(new OfficeDocument('/styled.pptx', pptx));
  assert.match(pptView, /font-family:'Aptos';font-size:28pt;color:#00a2ff;vertical-align:super/);
  assert.match(pptView, /<strong><em><u>PPT 样式<\/u><\/em><\/strong>/);
  assert.match(pptView, /text-align:center/);
});
test('Excel resolves relationships, strings, cached formulas, dates and all sparse rows and columns without paging', () => {
  const data = zip({
    'xl/workbook.xml': '<workbook><sheets><sheet name="数据 &amp; 汇总" r:id="r1"/><sheet name="另一个表" r:id="r2"/><sheet name="hidden" state="hidden" r:id="r3"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="r1" Target="worksheets/data.xml"/><Relationship Id="r2" Target="worksheets/empty.xml"/></Relationships>',
    'xl/sharedStrings.xml': '<sst><si><r><t>共享</t></r><r><t>字符串</t></r></si></sst>',
    'xl/styles.xml': '<styleSheet><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="10"/></cellXfs></styleSheet>',
    'xl/worksheets/data.xml': '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>00123</t></is></c><c r="C1"><f>1+1</f><v>2</v></c><c r="D1" s="1"><v>1</v></c><c r="E1" s="2"><v>0.125</v></c><c r="F1"><f>NOW()</f></c></row><row r="51"><c r="U51" t="inlineStr"><is><t>第二页右侧</t></is></c></row></sheetData></worksheet>',
    'xl/worksheets/empty.xml': '<worksheet><sheetData/></worksheet>',
  });
  const document = new OfficeDocument('/book.xlsx', data);
  const first = html(document);
  assert.deepEqual(grid(first).cells.map(cell => cell.slice(0, 3)), [[1, 1, '共享字符串'], [1, 2, '00123'], [1, 3, '2'], [1, 4, '1900-01-01'], [1, 5, '12.50%'], [1, 6, '[公式无缓存值]'], [51, 21, '第二页右侧']]);
  assert.ok(first.includes('数据 &amp; 汇总'));
  assert.doesNotMatch(first, /<t[dh]\b/);
  assert.doesNotMatch(first, /sheet-2/);
  assert.doesNotMatch(first, /paging|rows-next|columns-next|office-actions|office-readonly|data-ref="reload"/);
  assert.doesNotMatch(first, /office-file-mark|book\.xlsx|只读预览/);
  assert.match(first, /data-ref="sheet-0"[^>]*aria-pressed="true"/);
  assert.match(first, /data-pencode-selectable data-pencode-searchable data-pencode-preserve-whitespace/);
  assert.match(first, /data-pencode-highlight-rows/);
  assert.match(first, /data-pencode-highlight-cells/);
  assert.equal(document.event('next'), false);
  assert.equal(document.event('sheet-1'), true);
  assert.equal(grid(html(document)).columns, 26);
  assert.equal(grid(html(document)).rows, 100);
  assert.deepEqual(grid(html(document)).cells, []);
  assert.doesNotMatch(html(document), /共享字符串|第二页右侧/);
  assert.match(html(document), /data-ref="sheet-1"[^>]*aria-pressed="true"/);
  assert.match(html(document), /data-ref="sheet-0"[^>]*aria-pressed="false"/);
  assert.equal(document.event('sheet-99'), false);
  assert.equal(columnNumber('XFD1048576'), 16384); assert.equal(columnName(16384), 'XFD');
  assert.equal(columnNumber('XFE1'), 0);
});
test('Excel draws blank cells beyond small sheets and preserves larger sheet extents', () => {
  const workbook = (reference: string) => zip({
    'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" r:id="r1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="r1" Target="worksheets/data.xml"/></Relationships>',
    'xl/worksheets/data.xml': `<worksheet><sheetData><row r="${reference.replace(/\D/g, '')}"><c r="${reference}" t="inlineStr"><is><t>value</t></is></c></row></sheetData></worksheet>`,
  });
  for (const [reference, rows, columns] of [['C3', 100, 26], ['AA101', 101, 27]] as const) {
    const view = html(new OfficeDocument('/grid.xlsx', workbook(reference)));
    const data = grid(view);
    assert.equal(data.rows, rows);
    assert.equal(data.columns, columns);
    assert.deepEqual(data.cells, [[Number(reference.replace(/\D/g, '')), columnNumber(reference), 'value', 0]]);
    assert.doesNotMatch(view, /<t[drh]\b/);
  }
});
test('Excel uses saved column widths, row heights and alignment without expanding long-text columns', () => {
  const data = zip({
    'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" r:id="r1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="r1" Target="worksheets/data.xml"/></Relationships>',
    'xl/styles.xml': '<styleSheet><cellXfs><xf/><xf><alignment horizontal="center" vertical="center"/></xf><xf><alignment wrapText="1"/></xf></cellXfs></styleSheet>',
    'xl/worksheets/data.xml': `<worksheet><sheetFormatPr defaultColWidth="9" defaultRowHeight="15.75"/><cols><col min="1" max="1" width="9"/><col min="2" max="2" width="98"/><col min="3" max="4" width="36"/><col min="5" max="5" width="NaN"/><col min="0" max="16385" width="100"/></cols><sheetData><row r="1" ht="30"><c r="A1" t="inlineStr"><is><t>id</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>${'long text '.repeat(100)}</t></is></c></row><row r="2"><c r="B2" s="2" t="inlineStr"><is><t>wrapped text</t></is></c></row></sheetData></worksheet>`,
  });
  const view = html(new OfficeDocument('/dimensions.xlsx', data));
  const dataGrid = grid(view);
  assert.deepEqual(dataGrid.columnSizes, [[2, 686], [3, 252], [4, 252]]);
  assert.equal(dataGrid.columnWidth, 63);
  assert.equal(dataGrid.rowHeight, 21);
  assert.deepEqual(dataGrid.rowSizes, [[1, 40]]);
  assert.equal(dataGrid.styles[dataGrid.cells[1]![3]], 'text-align:center;vertical-align:middle;');
  assert.equal(dataGrid.styles[dataGrid.cells[2]![3]], 'white-space:pre-wrap;overflow-wrap:break-word;');
  assert.doesNotMatch(view, /NaN|Infinity/);
});
test('Excel retains full cell text and represents distant sparse cells without dense HTML', () => {
  const workbook = (row: number, value: string) => zip({
    'xl/workbook.xml': '<workbook><sheets><sheet name="Sheet1" r:id="r1"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="r1" Target="worksheets/data.xml"/></Relationships>',
    'xl/worksheets/data.xml': `<worksheet><sheetData><row r="${row}"><c r="A${row}" t="inlineStr"><is><t>${value}</t></is></c></row></sheetData></worksheet>`,
  });
  const value = '  ' + '长文本'.repeat(300) + '\n末尾搜索词  ';
  assert.equal(grid(html(new OfficeDocument('/long.xlsx', workbook(1, value)))).cells[0]![2], value);
  const sparse = html(new OfficeDocument('/sparse.xlsx', workbook(1048576, 'last')));
  assert.equal(grid(sparse).rows, 1048576);
  assert.deepEqual(grid(sparse).cells, [[1048576, 1, 'last', 0]]);
  assert.ok(sparse.length < 2000);
});
test('PowerPoint follows presentation order and excludes external images', () => {
  const data = zip({
    'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId r:id="second" id="257"/><p:sldId id="256" r:id="first"/></p:sldIdLst></p:presentation>',
    'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="first" Target="slides/slide1.xml"/><Relationship Id="second" Target="slides/slide2.xml"/></Relationships>',
    'ppt/slides/slide1.xml': '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>第一张</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    'ppt/slides/slide2.xml': '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>第二张</a:t></a:r></a:p></p:txBody></p:sp><p:pic><a:blip r:embed="external"/></p:pic></p:spTree></p:cSld></p:sld>',
    'ppt/slides/_rels/slide2.xml.rels': '<Relationships><Relationship Id="external" Target="https://example.test/private.png" TargetMode="External"/></Relationships>',
  });
  const document = new OfficeDocument('/slides.pptx', data);
  const first = document.view();
  assert.deepEqual(first.view?.events, [{ ref: 'slide-0', kind: 'click' }, { ref: 'slide-1', kind: 'click' }]);
  assert.match(first.view!.html!, /data-ref="slide-0"[^>]*aria-pressed="true"/);
  assert.match(first.view!.html!, /<article[^>]*>.*第二张/);
  assert.doesNotMatch(first.view!.html!, /https:/);
  assert.equal(document.event('previous'), false);
  assert.equal(document.event('slide-1'), true);
  assert.match(html(document), /data-ref="slide-1"[^>]*aria-pressed="true"/);
  assert.match(html(document), /<article[^>]*>.*第一张/);
  assert.equal(document.event('slide-99'), false);
  assert.equal(document.event('next'), false);
});
test('Excel refresh follows the selected sheet across edits and reordering', async () => {
  const abort = new AbortController(), handle = createViewer(abort.signal);
  const directory = await mkdtemp(join(tmpdir(), 'pencode-excel-'));
  const path = join(directory, '系统信息.xlsx');
  const save = (names: string[], revision: string) => writeFile(path, zip({
    'xl/workbook.xml': '<workbook><sheets>' + names.map((name, i) => `<sheet name="${name}" r:id="r${i}"/>`).join('') + '</sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships>' + names.map((_, i) => `<Relationship Id="r${i}" Target="worksheets/s${i}.xml"/>`).join('') + '</Relationships>',
    ...Object.fromEntries(names.map((name, i) => [`xl/worksheets/s${i}.xml`, `<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>${name} ${revision}</t></is></c></row></sheetData></worksheet>`])),
  }));
  try {
    await save(['目录', '数据'], 'v1');
    await handle(message(path));
    assert.equal(grid((await handle(message(path, 'click', 'sheet-1'))).view!.html!).cells[0]![2], '数据 v1');
    await save(['目录', '新增', '数据'], 'v2');
    assert.equal(grid((await handle(message(path, 'documentchange'))).view!.html!).cells[0]![2], '数据 v2');
    await save(['目录'], 'v3');
    assert.equal(grid((await handle(message(path, 'documentchange'))).view!.html!).cells[0]![2], '目录 v3');
  } finally { abort.abort(); await rm(directory, { recursive: true, force: true }); }
});
test('viewer retries corrupt files and refreshes without modifying the Office source', async () => {
  const abort = new AbortController();
  const directory = await mkdtemp(join(tmpdir(), 'pencode-office-'));
  try {
    const path = join(directory, 'a.docx'), handle = createViewer(abort.signal);
    await writeFile(path, 'not a zip');
    assert.match((await handle(message(path))).view!.html!, /无法预览/);
    const before = word('原始内容'); await writeFile(path, before);
    assert.match((await handle(message(path, 'click', 'reload'))).view!.html!, /原始内容/);
    assert.deepEqual(await readFile(path), Buffer.from(before));
    assert.deepEqual(await handle(message(path, 'click', 'unknown')), {});
    await writeFile(path, word('新内容'));
    assert.match((await handle(message(path, 'documentchange'))).view!.html!, /新内容/);
    abort.abort();
    await assert.rejects(handle(message(path)));
  } finally { await rm(directory, { recursive: true, force: true }); }
});
test('rejects DTDs, malformed XML and missing document parts', () => {
  for (const xml of ['<!DOCTYPE a [<!ENTITY e "x">]><document>&e;</document>', '<document><body></document>']) {
    assert.throws(() => new OfficeDocument('test.docx', zip({ 'word/document.xml': xml })), /XML/);
  }
  assert.throws(() => new OfficeDocument('test.docx', zip({})), /缺少/);
});
test('bundled JavaScript serves PEB requests under the actual host runtime', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pencode-office-peer-'));
  const child = spawn(process.execPath, ['dist/office-viewer.mjs'], { stdio: ['pipe', 'pipe', 'inherit'] });
  const exited = new Promise<number | null>((resolve, reject) => { child.once('exit', resolve); child.once('error', reject); });
  const reader = new Reader(child.stdout), writer = new Writer(child.stdin);
  try {
    await writer.send(Kind.Initialize, 1, encode(HelloCodec, { host: 'pencode', capabilities: 511 }));
    const registered = await reader.read(); assert.equal(registered?.kind, Kind.Register);
    const contributions = decode(ContributionsCodec, registered!.payload);
    assert.equal(contributions.actions[0]?.default, true); assert.equal(contributions.actionService, false);
    const docx = join(directory, 'test.docx');
    const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const relationships = 'http://schemas.openxmlformats.org/package/2006/relationships';
    await writeFile(docx, zip({
      '_rels/.rels': `<Relationships xmlns="${relationships}"><Relationship Id="r1" Type="${r}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'word/document.xml': `<w:document xmlns:w="${w}"><w:body>
        <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>真实进程</w:t></w:r></w:p>
        <w:p><w:r><w:br w:type="page"/></w:r></w:p>
        <w:p><w:r><w:t>第二页</w:t></w:r></w:p>
        <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1080" w:right="1680" w:bottom="1008" w:left="1180"/></w:sectPr>
      </w:body></w:document>`,
    }));
    const pptx = join(directory, 'test.pptx');
    const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const slide = (content: string) => `<p:sld xmlns:p="${p}" xmlns:a="${a}">
      <p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="E8F0FB"/></a:solidFill></p:bgPr></p:bg>
      <p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="3200"><a:latin typeface="Arial"/><a:ea typeface="Missing Office Font"/></a:rPr><a:t>${content}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
    await writeFile(pptx, zip({
      '_rels/.rels': `<Relationships xmlns="${relationships}"><Relationship Id="r1" Type="${r}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
      '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>',
      'ppt/presentation.xml': `<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="256" r:id="r1"/><p:sldId id="257" r:id="r2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
      'ppt/_rels/presentation.xml.rels': `<Relationships xmlns="${relationships}"><Relationship Id="r1" Type="${r}/slide" Target="slides/slide1.xml"/><Relationship Id="r2" Type="${r}/slide" Target="slides/slide2.xml"/></Relationships>`,
      'ppt/slides/slide1.xml': slide('PPT 甲'),
      'ppt/slides/slide2.xml': slide('PPT 乙'),
    }));
    let id = 1;
    const exchange = async (path: string, kind = 'open', ref = '') => {
      await writer.send(Kind.EditorRequest, ++id, encode(EditorMessageCodec, message(path, kind, ref)));
      const response = await reader.read();
      assert.equal(response?.id, id);
      assert.equal(response?.kind, Kind.EditorResponse);
      const reply = decode(EditorReplyCodec, response!.payload);
      assert.equal(reply.content, null);
      return reply.view!.html;
    };
    const wordView = await exchange(docx);
    assert.equal((wordView.match(/class="word-paper"/g) ?? []).length, 2);
    assert.match(wordView, /width:793\.73px;/);
    assert.match(wordView, /padding-left:78\.67px/);
    assert.match(wordView, /text-align:center[^>]*>.*真实进程/);
    const png = (view: string) => Buffer.from(view.match(/class="presentation-image" src="data:image\/png;base64,([^"]+)"/)![1]!, 'base64');
    const slideView = await exchange(pptx);
    assert.doesNotMatch(slideView, /office-toolbar/);
    assert.match(slideView, /<footer class="presentation-controls">.*data-ref="slide-text".*<\/footer><\/aside>/);
    const thumbnails = [...slideView.matchAll(/class="presentation-thumbnail"[^>]*><img src="([^"]+)"/g)];
    assert.equal(thumbnails.length, 2);
    assert.ok(thumbnails[0]![1] !== thumbnails[1]![1], 'different Chinese glyphs must produce different thumbnails');
    const firstSlide = png(slideView);
    assert.equal(firstSlide.readUInt32BE(16) / firstSlide.readUInt32BE(20), 4 / 3);
    const secondSlide = await exchange(pptx, 'click', 'slide-1');
    assert.match(secondSlide, /data-ref="slide-1"[^>]*aria-pressed="true"/);
    assert.ok(!png(secondSlide).equals(firstSlide), 'different Chinese glyphs must produce different slides');
    assert.match(await exchange(pptx, 'click', 'slide-text'), /presentation-text.*PPT 乙/);
    assert.deepEqual(png(await exchange(pptx, 'click', 'slide-text')), png(secondSlide));
    await exchange(pptx, 'click', 'slide-text');
    assert.match(await exchange(pptx, 'click', 'slide-0'), /data-ref="slide-text" aria-pressed="true"/);
    assert.deepEqual(png(await exchange(pptx, 'click', 'slide-text')), firstSlide);
    await writer.send(Kind.Shutdown, 0); assert.equal(await exited, 0);
  } finally { child.kill(); await rm(directory, { recursive: true, force: true }); }
});

test('one Office process keeps navigation independent for two views of the same file', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'pencode-office-sessions-'));
  const child = spawn(process.execPath, ['dist/office-viewer.mjs'], { stdio: ['pipe', 'pipe', 'inherit'] });
  const exited = new Promise<number | null>((resolve, reject) => { child.once('exit', resolve); child.once('error', reject); });
  const reader = new Reader(child.stdout), writer = new Writer(child.stdin);
  try {
    const path = join(directory, 'same.pptx');
    await writeFile(path, zip({
      'ppt/presentation.xml': '<p:presentation><p:sldIdLst><p:sldId r:id="one"/><p:sldId r:id="two"/></p:sldIdLst></p:presentation>',
      'ppt/_rels/presentation.xml.rels': '<Relationships><Relationship Id="one" Target="slides/slide1.xml"/><Relationship Id="two" Target="slides/slide2.xml"/></Relationships>',
      'ppt/slides/slide1.xml': '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>独立第一张</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      'ppt/slides/slide2.xml': '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>独立第二张</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
    }));
    const request = async (session: number, kind = 'open', ref = '') => {
      await writer.send(Kind.EditorRequest, 2, encode(EditorMessageCodec, message(path, kind, ref)), session);
      const response = await reader.read(); assert.equal(response?.session, session); assert.equal(response?.kind, Kind.EditorResponse);
      return decode(EditorReplyCodec, response!.payload);
    };
    for (const session of [1, 2]) {
      await writer.send(Kind.Initialize, 1, encode(HelloCodec, { capabilities: 1023 }), session);
      assert.equal((await reader.read())?.session, session);
      assert.match((await request(session)).view!.html, /<article[^>]*>.*独立第一张/s);
      if (session === 1) assert.match((await request(1, 'click', 'slide-1')).view!.html, /<article[^>]*>.*独立第二张/s);
    }
    assert.equal((await request(1, 'click', 'slide-1')).view, null);
    await writer.send(Kind.Shutdown, 0, undefined, 1);
    assert.match((await request(2, 'click', 'slide-1')).view!.html, /<article[^>]*>.*独立第二张/s);
    await writer.send(Kind.Shutdown, 0); assert.equal(await exited, 0);
  } finally { child.kill(); await rm(directory, { recursive: true, force: true }); }
});
