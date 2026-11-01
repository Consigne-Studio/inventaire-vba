// Imprime un Markdown en PDF A4 à la charte du rapport. Par défaut docs/instructions-client.md (fichier livré après la commande).
// Usage : node scripts/instructions.mjs [source.md sortie.pdf "Titre"]
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { versHtml } from '../src/rapport/markdown.mjs';
import { imprimerPdf } from '../src/rapport/pdf.mjs';

const R = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const b64 = async (n) => (await readFile(path.join(R, 'polices', n))).toString('base64');
/** Imprime `src` (Markdown) en `sortie` (PDF) et écrit à côté la version HTML lisible à l'écran. */
export async function imprimerMarkdown(src, sortie, titre, remplacer = (t) => t) {
// Titres # retirés (le titre est dans l'en-tête), titres ## ramenés au niveau ####.
const md = remplacer(await readFile(src, 'utf8')).replace(/^# .*$/gm, '').replace(/^## /gm, '#### ');
const html = `<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Inventaire VBA · ${titre}</title><style>
@font-face{font-family:Archivo;src:url(data:font/woff2;base64,${await b64('Archivo-var.woff2')}) format('woff2');font-weight:100 900;font-stretch:62% 125%}
@font-face{font-family:'Martian Mono';src:url(data:font/woff2;base64,${await b64('MartianMono-var.woff2')}) format('woff2');font-weight:100 800;font-stretch:75% 112.5%}
@font-face{font-family:Atkinson;src:url(data:font/woff2;base64,${await b64('AtkinsonHyperlegibleNext-var.woff2')}) format('woff2');font-weight:200 800}
@page{size:A4;margin:18mm 18mm 20mm;@bottom-right{content:'Inventaire VBA · DEV-607 · ' counter(page) ' / ' counter(pages);font:500 7.5px 'Martian Mono',monospace;color:#5C636A;letter-spacing:.06em}}
body{font:13.5px/1.55 Atkinson,sans-serif;color:#1B1E21;margin:0;background:#fff}
@media screen{body{max-width:760px;margin:32px auto;padding:0 max(16px,env(safe-area-inset-left,0px));font-size:15.5px}}
header{display:flex;gap:14px;align-items:center;border-bottom:1.5px solid #1B1E21;padding-bottom:12px;margin-bottom:18px}
header .t{font:800 40px/.9 Archivo,sans-serif;font-stretch:62%;text-transform:uppercase}
header .d{font:500 9px/1.7 'Martian Mono',monospace;font-stretch:87.5%;letter-spacing:.07em;text-transform:uppercase;color:#5C636A}
h4,h5{font:800 19px/1.1 Archivo,sans-serif;font-stretch:75%;text-transform:uppercase;margin:18px 0 6px;break-after:avoid}
p{margin:0 0 8px}ul{margin:0 0 8px 18px;padding:0}li{margin:0 0 5px}
code{font-family:'Martian Mono',monospace;font-stretch:87.5%;font-size:.85em;background:#E6E8EA;padding:0 3px;word-break:break-word}
</style></head><body><header><svg width="64" height="64" viewBox="0 0 64 64"><circle cx="32" cy="32" r="30.6" fill="#2BB673" stroke="#1B1E21" stroke-width="2.8"/><line x1="1.4" y1="32" x2="62.6" y2="32" stroke="#1B1E21" stroke-width="2.8"/><text x="32" y="26" text-anchor="middle" font-family="Archivo" font-stretch="62%" font-weight="800" font-size="19" fill="#1B1E21">DEV</text><text x="32" y="49" text-anchor="middle" font-family="Martian Mono" font-stretch="87.5%" font-size="12.5" fill="#1B1E21">607</text></svg>
<div><div class="t">${titre}</div><div class="d">Inventaire des macros VBA · Consigne</div></div></header>${versHtml(md)}</body></html>`;
const h = sortie.replace(/\.pdf$/, '.html');
await writeFile(h, html);
await imprimerPdf(h, sortie);
return { html: h, pdf: sortie };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [src = path.join(R, 'docs', 'instructions-client.md'), sortie = path.join(R, 'docs', 'instructions-client.pdf'), titre = 'Après votre commande'] = process.argv.slice(2);
  console.log((await imprimerMarkdown(src, sortie, titre)).pdf);
}
