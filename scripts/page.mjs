// Assemble les pages autonomes (polices et images intégrées).
//   node scripts/page.mjs
//     vente/dist/page.html et index.html : aperçu de la page de vente (mentions légales non renseignées)
//     docs/index.html : documentation générée depuis README.md (GitHub Pages, dossier /docs)
//     docs/exemple/ : rapport d'exemple interactif et diagnostic livré
//   node scripts/page.mjs --publication <dossier>
//     site prêt pour Cloudflare Pages dans <dossier>/site/ (index.html, cgv.html, cgv.pdf, cgv-remediation.html et .pdf),
//     CGV lues dans <dossier>/cgv.md et dans le dossier voisin *-remediation-courriels-mise-en-vente/cgv.md.
//     Lit STUDIO_NOM_LEGAL, STUDIO_SIRET, STUDIO_ADRESSE_PRO, STUDIO_TELEPHONE et STUDIO_SUPPORT_EMAIL, refuse s'il en manque.
//   node scripts/page.mjs --cgv-modele <dossier>
//     CGV en HTML gardant les {{STUDIO_…}} (<dossier>/livrables/<produit>-cgv.html, produit lu dans publication.json),
//     rempli par launch/publier.mjs à l'envoi.
import { readFile, writeFile, mkdir, copyFile, rm, readdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { e } from '../src/rapport/markdown.mjs';
import { imprimerMarkdown } from './instructions.mjs';

const R = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const DEPOT = 'https://github.com/Consigne-Studio/inventaire-vba/blob/main/';
const b64 = async (rel, type) => `data:${type};base64,${(await readFile(path.join(R, rel))).toString('base64')}`;
const polices = {
  __ARCHIVO__: await b64('polices/Archivo-var.woff2', 'font/woff2'),
  __MARTIAN__: await b64('polices/MartianMono-var.woff2', 'font/woff2'),
  __ATKINSON__: await b64('polices/AtkinsonHyperlegibleNext-var.woff2', 'font/woff2'),
};
const remplir = (html, rempl) => {
  for (const [k, v] of Object.entries(rempl)) html = html.split(k).join(v);
  const reste = html.match(/__[A-Z0-9_]+__/);
  if (reste) throw new Error(`Marqueur non remplacé : ${reste[0]}`);
  return html;
};
const document = (corps) => `<!doctype html>\n<html lang="fr">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">\n</head>\n<body>\n${corps}\n</body>\n</html>\n`;

const VARIABLES = { NOM_LEGAL: 'STUDIO_NOM_LEGAL', SIRET: 'STUDIO_SIRET', ADRESSE: 'STUDIO_ADRESSE_PRO', TELEPHONE: 'STUDIO_TELEPHONE', EMAIL: 'STUDIO_SUPPORT_EMAIL' };
const args = process.argv.slice(2);
const opt = (n) => { const k = args.indexOf(n); return k >= 0 ? path.resolve(args[k + 1] ?? '') : null; };

if (opt('--cgv-modele')) {
  const dossier = opt('--cgv-modele');
  const { produit } = JSON.parse(await readFile(path.join(dossier, 'publication.json'), 'utf8'));
  await mkdir(path.join(dossier, 'livrables'), { recursive: true });
  const { pdf } = await imprimerMarkdown(path.join(dossier, 'cgv.md'), path.join(dossier, 'livrables', `${produit}-cgv.pdf`), 'Conditions générales de vente');
  await rm(pdf);
  console.log(`CGV modèle : livrables/${produit}-cgv.html`);
  process.exit(0);
}
const dossierPub = opt('--publication');
let mentions = Object.fromEntries(Object.keys(VARIABLES).map((k) => [k, k === 'EMAIL' ? 'adresse de support' : 'renseigné à la mise en ligne']));
if (dossierPub) {
  const manquent = Object.values(VARIABLES).filter((v) => !process.env[v]?.trim());
  if (manquent.length) { console.error(`Variables manquantes : ${manquent.join(', ')} (LANCEMENT.md, étape 1).`); process.exit(1); }
  mentions = Object.fromEntries(Object.entries(VARIABLES).map(([k, v]) => [k, process.env[v].trim()]));
}

// Page de vente
const images = { __IMG_DIAG_P1__: 'diagnostic-p1', __IMG_REGLES__: 'rapport-regles', __IMG_ECHEANCES__: 'rapport-echeances', __IMG_ESTIMATION__: 'rapport-estimation', __IMG_COURRIEL__: 'rapport-courriel' };
// Hauteur d'un JPEG (segment SOF), pour les attributs width/height des captures dont la hauteur varie.
const hauteurJpeg = (buf) => { for (let i = 2; i < buf.length;) { const m = buf[i + 1]; if (m >= 0xc0 && m <= 0xc3) return buf.readUInt16BE(i + 5); i += 2 + buf.readUInt16BE(i + 2); } throw new Error('JPEG sans SOF'); };
const rempl = { ...polices, __BASE__: dossierPub ? '' : 'https://inventaire-vba.pages.dev/', __H_COURRIEL__: String(hauteurJpeg(await readFile(path.join(R, 'docs', 'img', 'rapport-courriel.jpg')))) };
for (const [k, v] of Object.entries(mentions)) rempl[`__${k}__`] = e(v);
for (const [k, n] of Object.entries(images)) rempl[k] = await b64(`docs/img/${n}.jpg`, 'image/jpeg');
const vente = remplir(await readFile(path.join(R, 'vente', 'page.html'), 'utf8'), rempl);
if (dossierPub) {
  const site = path.join(dossierPub, 'site');
  await rm(site, { recursive: true, force: true });
  await mkdir(site, { recursive: true });
  await writeFile(path.join(site, 'index.html'), document(vente));
  const valeurs = Object.fromEntries(Object.entries(VARIABLES).map(([k, v]) => [v, mentions[k]]));
  const remplacer = (t) => {
    const r = t.replace(/\{\{(STUDIO_[A-Z_]+)\}\}/g, (m, k) => valeurs[k] ?? m);
    const reste = r.match(/\{\{[^}]+\}\}/);
    if (reste) throw new Error(`CGV : « ${reste[0]} » non renseigné`);
    return r;
  };
  await imprimerMarkdown(path.join(dossierPub, 'cgv.md'), path.join(site, 'cgv.pdf'), 'Conditions générales de vente', remplacer);
  // La page lie aussi les CGV de la remédiation de l'envoi de courriels, rangées dans le dossier de publication voisin.
  const parent = path.dirname(dossierPub);
  const voisin = (await readdir(parent)).filter((n) => n.endsWith('-remediation-courriels-mise-en-vente')).sort().pop();
  if (!voisin) { console.error(`CGV de la remédiation introuvables : aucun dossier *-remediation-courriels-mise-en-vente dans ${parent}.`); process.exit(1); }
  await imprimerMarkdown(path.join(parent, voisin, 'cgv.md'), path.join(site, 'cgv-remediation.pdf'), 'Conditions générales de vente', remplacer);
  console.log(`site : ${path.relative(process.cwd(), site)}/ (index.html, cgv.html, cgv.pdf, cgv-remediation.html, cgv-remediation.pdf)`);
  process.exit(0);
}
await mkdir(path.join(R, 'vente', 'dist'), { recursive: true });
await writeFile(path.join(R, 'vente', 'dist', 'page.html'), vente);
await writeFile(path.join(R, 'vente', 'dist', 'index.html'), document(vente));

// Documentation : README.md rendu (titres, listes, tableaux, blocs de code, images, liens).
const lien = (href) => href.startsWith('#') || /^https?:/.test(href) ? href : href.startsWith('docs/') ? href.slice(5).replace(/\.md$/, '.html') : href.startsWith('exemples/exemple-industrie/rapport.html') ? 'exemple/rapport.html' : DEPOT + href;
const ancre = (t) => t.toLowerCase().replace(/[`*]/g, '').trim().replace(/[^\p{L}\p{N}\s-]/gu, '').replace(/\s+/g, '-');
const enLigne = (s) => e(s)
  .replace(/!\[([^\]]*)\]\(([^)\s]+)\)/g, (_, alt, src) => `<figure><img src="${lien(src)}" alt="${alt}" loading="lazy"></figure>`)
  .replace(/`([^`]+)`/g, '<code>$1</code>')
  .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
  .replace(/(^|[\s(])\*([^*\s][^*]*)\*/g, '$1<em>$2</em>')
  .replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (_, t, h) => `<a href="${lien(h.replace(/&amp;/g, '&'))}">${t}</a>`);
function markdown(md) {
  const out = []; const toc = [];
  const L = md.split(/\r?\n/);
  for (let i = 0; i < L.length; i++) {
    const l = L[i];
    if (/^<!--.*-->$/.test(l.trim()) || !l.trim()) continue;
    if (l.startsWith('```')) { const code = []; while (++i < L.length && !L[i].startsWith('```')) code.push(L[i]); out.push(`<pre><code>${e(code.join('\n'))}</code></pre>`); continue; }
    const t = /^(#{1,3})\s+(.*)$/.exec(l);
    if (t) { const n = t[1].length; const id = ancre(t[2]); if (n === 2) toc.push([id, t[2]]); out.push(n === 1 ? `<h1 class="titre">${enLigne(t[2])}</h1>` : `<h${n} id="${id}">${enLigne(t[2])}</h${n}>`); continue; }
    if (l.startsWith('|')) {
      const rows = []; while (i < L.length && L[i].startsWith('|')) rows.push(L[i++]); i--;
      const cells = (r) => r.replace(/^\||\|$/g, '').split('|').map((c) => c.trim());
      out.push(`<div class="tab"><table><thead><tr>${cells(rows[0]).map((c) => `<th>${enLigne(c)}</th>`).join('')}</tr></thead><tbody>${rows.slice(2).map((r) => `<tr>${cells(r).map((c) => `<td>${enLigne(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`);
      continue;
    }
    if (/^\s*([*-]|\d+\.)\s+/.test(l)) {
      const ol = /^\s*\d+\./.test(l); const items = [];
      while (i < L.length && /^\s*([*-]|\d+\.)\s+/.test(L[i])) items.push(L[i++].replace(/^\s*([*-]|\d+\.)\s+/, '')); i--;
      out.push(`<${ol ? 'ol' : 'ul'}>${items.map((x) => `<li>${enLigne(x)}</li>`).join('')}</${ol ? 'ol' : 'ul'}>`); continue;
    }
    const para = [l]; while (i + 1 < L.length && L[i + 1].trim() && !/^(#|\||```|\s*([*-]|\d+\.)\s|<!--)/.test(L[i + 1])) para.push(L[++i]);
    out.push(`<p>${enLigne(para.join(' '))}</p>`);
  }
  return { html: out.join('\n'), toc };
}
const { html, toc } = markdown(await readFile(path.join(R, 'README.md'), 'utf8'));
const entete = vente.slice(vente.indexOf('<header'), vente.indexOf('</header>') + 9).replace(/<nav[\s\S]*?<\/nav>/, '<nav class="data" aria-label="Liens"><a href="exemple/rapport.html">Rapport d\'exemple</a><a href="https://github.com/Consigne-Studio/inventaire-vba">Code source</a><a href="https://github.com/Consigne-Studio/inventaire-vba/releases/latest">Télécharger</a></nav>');
const gabarit = remplir(await readFile(path.join(R, 'scripts', 'gabarits', 'docs.html'), 'utf8'), polices);
const docs = `${gabarit}\n${entete}\n<main class="bande doc" id="haut"><div class="int"><nav class="data" aria-label="Sommaire">${toc.map(([id, t]) => `<a href="#${id}">${e(t.replace(/`/g, ''))}</a>`).join('')}</nav><article>${html}</article></div></main>`;
await writeFile(path.join(R, 'docs', 'index.html'), document(docs));
// Article : nouvel Outlook et SMTP AUTH (même gabarit que la documentation).
{
  const a = markdown(await readFile(path.join(R, 'docs', 'nouvel-outlook-smtp-auth.md'), 'utf8'));
  const page = `${gabarit}\n${entete}\n<main class="bande doc" id="haut"><div class="int"><nav class="data" aria-label="Sommaire"><a href="./">Documentation</a>${a.toc.map(([id, t]) => `<a href="#${id}">${e(t.replace(/`/g, ''))}</a>`).join('')}</nav><article>${a.html}</article></div></main>`;
  await writeFile(path.join(R, 'docs', 'nouvel-outlook-smtp-auth.html'), document(page).replace(/<title>[^<]*<\/title>/, '<title>Nouvel Outlook et SMTP AUTH : ce qui casse, et quand · inventaire-vba</title>'));
}
await mkdir(path.join(R, 'docs', 'exemple'), { recursive: true });
await copyFile(path.join(R, 'exemples', 'exemple-industrie', 'rapport.html'), path.join(R, 'docs', 'exemple', 'rapport.html'));
await copyFile(path.join(R, 'exemples', 'exemple-industrie', 'diagnostic-livre', 'rapport.pdf'), path.join(R, 'docs', 'exemple', 'diagnostic-livre.pdf'));
console.log(`vente : ${Math.round(vente.length / 1024)} Ko · docs : ${Math.round(docs.length / 1024)} Ko`);
