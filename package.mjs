import { cp, mkdir, rm } from 'node:fs/promises';
await rm('package', { recursive: true, force: true });
await mkdir('package');
for (const path of ['dist', 'resources', 'package.json', 'LICENSE', 'README.md']) {
  await cp(path, `package/${path}`, { recursive: true });
}
