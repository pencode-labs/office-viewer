import { attr } from './archive.js';
import type { XmlNode } from './archive.js';
export function escape(value: string): string {
  return value.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
}
export function enabled(node: XmlNode | undefined): boolean { return !!node && !['0', 'false', 'off'].includes(attr(node, 'val')); }
