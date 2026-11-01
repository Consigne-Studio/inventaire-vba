// Captures du rapport d'exemple pour le README (docs/img/*.png), la page de vente et les visuels (docs/img/*.jpg),
// plus les pages 1 et 8 du diagnostic livré (PDF rendu par Chromium).
// Usage : node scripts/exemple.mjs && node scripts/captures.mjs   (Playwright installé globalement, poppler-utils)
import path from 'node:path';
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const R = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const require = createRequire(path.join(execSync('npm root -g').toString().trim(), 'x.js'));
const { chromium } = require('playwright');
const ex = path.join(R, 'exemples', 'exemple-industrie');
const sortie = (n) => path.join(R, 'docs', 'img', n);
const b = await chromium.launch();
const ouvrir = async (fichier, w, h, scale) => {
  const p = await b.newPage({ viewport: { width: w, height: h }, deviceScaleFactor: scale, colorScheme: 'light' });
  await p.goto('file://' + fichier); await p.evaluate(() => document.fonts.ready);
  return p;
};
const SECTIONS = ['echeances', 'regles', 'courriel', 'fichiers', 'estimation', 'copies'];
// PNG à 2x pour le README, JPEG à 1200 px de large pour la page.
for (const [scale, type] of [[2, 'png'], [1200 / 1132, 'jpeg']]) {
  const p = await ouvrir(path.join(ex, 'rapport.html'), 1280, 800, scale);
  const ext = type === 'png' ? 'png' : 'jpg';
  const opts = type === 'jpeg' ? { type, quality: 82 } : { type };
  await p.screenshot({ path: sortie(`rapport.${ext}`), ...opts });
  await p.evaluate(() => { const d = document.querySelector('#regles details'); if (d) d.open = true; });
  for (const id of SECTIONS) {
    const bb = await p.locator(`#${id}`).boundingBox();
    await p.screenshot({ path: sortie(`rapport-${id}.${ext}`), clip: { x: bb.x, y: bb.y, width: bb.width, height: Math.min(bb.height, 1100) }, fullPage: true, ...opts });
  }
  await p.close();
}
const m = await ouvrir(path.join(ex, 'rapport.html'), 390, 844, 2);
await m.screenshot({ path: sortie('rapport-mobile.png') });
const sw = await m.evaluate(() => document.documentElement.scrollWidth);
await m.close();
await b.close();
// Pages du PDF livré (poppler).
const pdf = path.join(ex, 'diagnostic-livre', 'rapport.pdf');
for (const n of [1, 8]) execSync(`pdftoppm -r 110 -jpeg -jpegopt quality=82 -singlefile -f ${n} -l ${n} "${pdf}" "${sortie(`diagnostic-p${n}`)}"`);
console.log(`captures écrites dans docs/img (largeur mobile ${sw} px)`);
