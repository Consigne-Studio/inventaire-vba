// Impression PDF du rapport via Chromium (Playwright), si installé sur le poste (localement ou avec npm i -g).
import { createRequire } from 'node:module';
import { execSync } from 'node:child_process';
import path from 'node:path';

async function chargerPlaywright() {
  try { return await import('playwright'); } catch { /* essai suivant : installation globale */ }
  try {
    const racine = execSync('npm root -g', { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], shell: true, windowsHide: true }).trim();
    return createRequire(path.join(racine, 'noop.js'))('playwright');
  } catch {
    throw new Error('--pdf demande Playwright : npm i -g playwright && npx playwright install chromium, ou ouvrez rapport.html et imprimez-le en PDF depuis le navigateur');
  }
}

export async function imprimerPdf(html, pdf) {
  const pw = await chargerPlaywright();
  const navigateur = await pw.chromium.launch(process.env.PW_CHROMIUM ? { executablePath: process.env.PW_CHROMIUM } : {});
  try {
    const page = await navigateur.newPage();
    await page.goto(`file://${html.replace(/\\/g, '/')}`);
    await page.evaluate(() => document.fonts.ready);
    await page.pdf({ path: pdf, format: 'A4', printBackground: true, margin: { top: '14mm', bottom: '14mm', left: '12mm', right: '12mm' } });
  } finally { await navigateur.close(); }
}
