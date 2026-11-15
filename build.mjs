// Build step: copies the app into ./dist and adds the pdf.js engine (installed by npm) into ./dist/vendor
import { cpSync, mkdirSync, existsSync, rmSync, readFileSync } from 'node:fs';

const OUT = 'dist';
const APP = [
  'index.html', 'style.css', 'app.js', 'pptx.js', 'ppt-legacy.js', 'readers.js', 'jszip.min.js',
  'sw.js', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'
];

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT + '/vendor', { recursive: true });

for (const f of APP) {
  if (!existsSync(f)) throw new Error('Missing app file: ' + f);
  cpSync(f, `${OUT}/${f}`);
}

const pdf = 'node_modules/pdfjs-dist';
if (!existsSync(pdf)) throw new Error('pdfjs-dist is not installed. Run "npm install" first.');
const legacy = existsSync(`${pdf}/legacy/build/pdf.min.mjs`) ? `${pdf}/legacy/build` : `${pdf}/build`;
cpSync(`${legacy}/pdf.min.mjs`, `${OUT}/vendor/pdf.min.mjs`);
cpSync(`${legacy}/pdf.worker.min.mjs`, `${OUT}/vendor/pdf.worker.min.mjs`);
for (const d of ['cmaps', 'standard_fonts', 'wasm', 'iccs']) {
  if (existsSync(`${pdf}/${d}`)) cpSync(`${pdf}/${d}`, `${OUT}/vendor/${d}`, { recursive: true });
}
const v = JSON.parse(readFileSync(`${pdf}/package.json`, 'utf8')).version;
console.log(`Build OK -> ${OUT}/ (pdf.js ${v}, using ${legacy})`);
