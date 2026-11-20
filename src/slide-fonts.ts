import { execFile } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { promisify } from 'node:util';
import { collectUsedFonts, DEFAULT_FONT_MAPPING, getMappedFont } from 'pptx-glimpse';
import type { FontBuffer } from 'pptx-glimpse';

const matchFont = promisify(execFile);

export async function slideFonts(input: Uint8Array): Promise<FontBuffer[] | undefined> {
  if (process.platform !== 'linux') return undefined;

  // Fontconfig selects the files this deck needs without parsing every font
  // installed on the host before replying to the editor's open request.
  const { fonts } = collectUsedFonts(input);
  const families = [...new Set(['sans-serif', ...fonts.slice(0, 12)])];
  const matches = await Promise.allSettled(families.map(async name => {
    const substitute = getMappedFont(name, DEFAULT_FONT_MAPPING);
    const family = name.replace(/[\\,:=]/g, '\\$&');
    const pattern = substitute ? family + ',' + substitute : family;
    const { stdout } = await matchFont('fc-match', ['-f', '%{file}', '--', pattern], { timeout: 800 });
    return { name, path: stdout.trim() };
  }));
  const files = new Map<string, Promise<Buffer>>();
  const buffers: FontBuffer[] = [];
  for (const match of matches) {
    if (match.status !== 'fulfilled' || !match.value.path) continue;
    const { name, path } = match.value;
    if (!files.has(path)) files.set(path, readFile(path));
    try { buffers.push({ name, data: await files.get(path)! }); }
    catch { /* The font may have been removed since Fontconfig indexed it. */ }
  }
  return buffers;
}
