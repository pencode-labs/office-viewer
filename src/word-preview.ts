import { resolveObjectURL, type Blob } from 'node:buffer';
import { DOMParser as XMLDOMParser, XMLSerializer } from '@xmldom/xmldom';
import { renderAsync } from 'docx-preview';
import { parseHTML } from 'linkedom';
import { imageSize } from 'image-size';
import { MaxDocument } from '@pencode/peb';

class DOMParser extends XMLDOMParser {
  override parseFromString(...args: Parameters<XMLDOMParser['parseFromString']>) {
    const document = super.parseFromString(...args);
    return Object.assign(document, { firstElementChild: document.documentElement });
  }
}

const { document, Node } = parseHTML('<!doctype html><html></html>');
// docx-preview uses a DOM to build the view. These implementations only parse
// and serialize markup in the extension process.
Object.assign(globalThis, { document, Node, DOMParser, XMLSerializer });

function cssPixels(css: string): string {
  // Onix uses CSS pixels for absolute lengths. DOCX uses points at 72 per inch.
  return css.replace(/"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|(-?(?:\d+(?:\.\d+)?|\.\d+))pt\b/g,
    (token, points: string | undefined) => points === undefined ? token : `${(Number(points) * 4 / 3).toFixed(2)}px`);
}

export async function renderWord(input: Uint8Array): Promise<{ html: string; css: string } | undefined> {
  const body = document.createElement('div');
  const styles = document.createElement('div');
  await renderAsync(input, body, styles, {
    className: 'word-paper', ignoreFonts: true, ignoreLastRenderedPageBreak: false,
    renderAltChunks: false,
  });
  const wrapper = body.querySelector('.word-paper-wrapper');
  if (!wrapper) return;

  let pixels = 0;
  const images = new Map<string, string>();
  for (const img of body.querySelectorAll('img')) {
    const src = img.getAttribute('src') ?? '';
    let data = images.get(src);
    if (data === undefined) {
      data = '';
      const blob = resolveObjectURL(src) as Blob | undefined;
      if (blob) URL.revokeObjectURL(src);
      if (blob && blob.size <= 256 << 10 && images.size < 64) {
        const bytes = Buffer.from(await blob.arrayBuffer());
        try {
          const { width, height, type } = imageSize(bytes);
          if (width <= 2048 && height <= 2048 && pixels + width * height <= 4 << 20 && (type === 'png' || type === 'jpg')) {
            data = 'data:image/' + (type === 'jpg' ? 'jpeg' : 'png') + ';base64,' + bytes.toString('base64');
            pixels += width * height;
          }
        } catch { /* Unsupported or damaged embedded image. */ }
      }
      images.set(src, data);
    }
    if (data) img.setAttribute('src', data);
    else img.replaceWith(document.createTextNode('[图片不可预览]'));
  }

  // Pencode owns element IDs and link actions. Keep the text and formatting.
  for (const element of body.querySelectorAll('[id], a[href]')) {
    element.removeAttribute('id');
    element.removeAttribute('href');
  }
  for (const element of body.querySelectorAll('[style]')) {
    element.setAttribute('style', cssPixels(element.getAttribute('style')!));
  }
  const html = wrapper.outerHTML;
  const css = cssPixels(styles.textContent);
  if (Buffer.byteLength(html) > MaxDocument - 4096 || Buffer.byteLength(css) > MaxDocument / 2) return;
  return { html, css };
}
