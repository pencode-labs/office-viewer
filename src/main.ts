import { ActionLocation, Match, serveExtension } from '@pencode/peb';
import { createViewer } from './viewer.js';
await serveExtension({ actions: [{
  id: 'office', text: 'Office Viewer', tooltip: 'Office 预览',
  icon: 'resources/office.svg', location: ActionLocation.OpenWith, default: true,
  selectors: ['.docx', '.xlsx', '.pptx', '.DOCX', '.XLSX', '.PPTX'].map(pattern => ({ match: Match.FileSuffix, pattern })),
}] }, { newEditor: createViewer }).catch(error => { console.error(error); process.exitCode = 1; });
