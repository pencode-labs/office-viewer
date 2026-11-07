import { build } from 'esbuild';
await build({
  entryPoints: ['src/main.ts'], outfile: 'dist/office-viewer.mjs', bundle: true,
  platform: 'node', format: 'esm', target: 'node22', legalComments: 'eof',
  loader: { '.css': 'text' },
});
