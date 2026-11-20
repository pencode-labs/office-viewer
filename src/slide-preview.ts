import { readFile } from 'node:fs/promises';
import { initWasm, Resvg, type CustomFontsOptions } from '@resvg/resvg-wasm';
import { readPptx } from '@pptx-glimpse/document';
import { renderPptxSourceModelToSvg } from 'pptx-glimpse';
import { slideFonts } from './slide-fonts.js';

const wasm = new URL('../dist/resvg.wasm', import.meta.url);
let rasterizer: Promise<void> | undefined;
let fallbackFont: Promise<Buffer> | undefined;
const thumbnailRadius = 3;
const displaySize = 1360;
const previewBudget = 1100 << 10;

type SlideImage = { full: string; thumbnail: string };

export class SlidePreview {
  private readonly source: ReturnType<typeof readPptx>;
  private readonly fonts: ReturnType<typeof slideFonts>;
  private readonly images = new Map<number, SlideImage>();

  constructor(input: Uint8Array, private readonly width: number, private readonly height: number) {
    this.source = readPptx(input);
    this.fonts = slideFonts(input);
  }

  async show(index: number, count: number): Promise<void> {
    const slides: number[] = [];
    for (let i = Math.max(0, index - thumbnailRadius); i <= Math.min(index + thumbnailRadius, count - 1); i++) {
      if (!this.images.has(i) || (i === index && !this.images.get(i)!.full)) slides.push(i + 1);
    }
    if (!slides.length) return;
    const [fonts, fallback] = await Promise.all([
      this.fonts,
      fallbackFont ??= readFile(new URL('../resources/fonts/NotoSansSC-Regular.ttf', import.meta.url)),
      rasterizer ??= readFile(wasm).then(bytes => initWasm(bytes)),
    ]);
    const { slides: rendered } = await renderPptxSourceModelToSvg(this.source, { slides, fonts, textOutput: 'text' });
    for (const { slideNumber, svg } of rendered) {
      const i = slideNumber - 1;
      try {
        // resvg needs the SVG's subset fonts as bytes. Keeping text until this
        // step lets it choose a fallback for individual missing glyphs.
        const fontBuffers: Uint8Array[] = Array.from(
          svg.matchAll(/src:url\(data:font\/otf;base64,([A-Za-z0-9+/=]+)\)/g),
          match => Buffer.from(match[1]!, 'base64'),
        );
        fontBuffers.push(fallback);
        const font = { fontBuffers, defaultFontFamily: 'Noto Sans SC', sansSerifFamily: 'Noto Sans SC' };
        const thumbnail = this.images.get(i)?.thumbnail || rasterize(svg, 192, this.width >= this.height, font);
        let full = '';
        if (i === index) {
          for (const size of [displaySize, 960, 640]) {
            full = rasterize(svg, size, this.width >= this.height, font);
            if (full.length <= previewBudget) break;
            full = '';
          }
        }
        this.images.set(i, { full, thumbnail });
      } catch {
        this.images.set(i, { full: '', thumbnail: '' });
      }
    }
    for (const [i, image] of this.images) {
      if (Math.abs(i - index) > thumbnailRadius) this.images.delete(i);
      else if (i !== index) image.full = '';
    }
  }

  image(index: number): string { return this.images.get(index)?.full ?? ''; }
  thumbnail(index: number): string { return this.images.get(index)?.thumbnail ?? ''; }
}

function rasterize(svg: string, size: number, landscape: boolean, font: CustomFontsOptions): string {
  const renderer = new Resvg(svg, { fitTo: { mode: landscape ? 'width' : 'height', value: size }, background: '#fff', font });
  try {
    const image = renderer.render();
    try {
      return 'data:image/png;base64,' + Buffer.from(image.asPng()).toString('base64');
    } finally { image.free(); }
  } finally { renderer.free(); }
}
