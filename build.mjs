import { copyFile } from 'node:fs/promises';
import { build } from 'esbuild';

await build({
  entryPoints: ['src/main.ts'], outfile: 'dist/office-viewer.mjs', bundle: true,
  platform: 'node', format: 'esm', target: 'node22', legalComments: 'eof',
  banner: { js: `import { createRequire as createBundleRequire } from 'node:module';
const require = createBundleRequire(import.meta.url);` },
});
await copyFile('node_modules/@resvg/resvg-wasm/index_bg.wasm', 'dist/resvg.wasm');
