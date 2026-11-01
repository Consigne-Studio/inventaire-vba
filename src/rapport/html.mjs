// Rapport HTML autonome : un seul fichier, polices intégrées, aucune ressource externe, imprimable.
// Charte « Consigne » : gris process, une seule couleur de signal (vert marche, gamme outillage) portée par la bulle.
import { REGLES, PRIORITES } from '../regles.mjs';
import { e, enLigne, versHtml, sections } from './markdown.mjs';
import { archivo, martian, atkinson } from './polices-data.mjs';

const C = { process: '#D4D7DA', releve: '#F3F4F5', carter: '#9AA0A6', gris: '#5C636A', graphite: '#1B1E21', signal: '#2BB673' };
const nf = (n, d = 0) => Number(n).toLocaleString('fr-FR', { maximumFractionDigits: d, minimumFractionDigits: d }).replace(/ | /g, ' ');
const date = (iso) => (iso ? iso.slice(0, 10) : '—');

// Échéances datées, dans l'ordre du calendrier. Les règles rattachées comptent les fichiers touchés.
export const ECHEANCES = [
  { axe: '64 bits', quand: 'Aujourd\'hui', detail: 'Office 64 bits est installé par défaut depuis Office 2019 et Microsoft 365.', regles: ['B64-001', 'B64-002', 'B64-003', 'B64-004'] },
  { axe: 'ActiveX', quand: 'Depuis avril 2025', detail: 'Contrôles ActiveX désactivés par défaut (Microsoft 365 2504, Office 2024).', regles: ['ACX-001'] },
  { axe: 'Signatures V3', quand: 'Dès la stratégie activée', detail: '« Only trust VBA macros that use V3 signatures » : les signatures anciennes sont ignorées.', regles: ['SIG-001', 'SIG-002'] },
  { axe: 'SMTP AUTH', quand: 'Fin décembre 2026', detail: 'Exchange Online désactive par défaut l\'envoi SMTP par mot de passe (authentification de base).', regles: ['MEL-001'] },
  { axe: 'VBScript', quand: 'Vers 2026-2027', detail: 'Phase 2 : VBScript désactivé par défaut dans Windows. Touche les .vbs et VB6 partout, RegExp sur Office LTSC et MSI.', regles: ['VBS-001', 'VBS-002', 'VBS-003', 'VBS-010', 'VB6-010'] },
  { axe: 'Nouvel Outlook', quand: 'Mars 2027', detail: 'Phase « opt-out » en entreprise ; le nouvel Outlook n\'a ni COM, ni VBA, ni MAPI.', regles: ['OUT-001', 'OUT-002', 'OUT-003'] },
  { axe: 'Parc mixte', quand: 'Dès qu\'un poste n\'est pas en 2508+', detail: 'RegExp natif : ne compile pas sur Office LTSC, 2021, 2019.', regles: ['VBS-004'] },
  { axe: 'Retraits déjà faits', quand: 'Déjà en vigueur', detail: 'Internet Explorer, DDE, macros Excel 4.0.', regles: ['IE-001', 'DDE-001', 'XLM-001', 'XLM-002'] },
];

const CSS = `
@font-face{font-family:Archivo;src:url(data:font/woff2;base64,${archivo}) format('woff2');font-weight:100 900;font-stretch:62% 125%}
@font-face{font-family:'Martian Mono';src:url(data:font/woff2;base64,${martian}) format('woff2');font-weight:100 800;font-stretch:75% 112.5%}
@font-face{font-family:Atkinson;src:url(data:font/woff2;base64,${atkinson}) format('woff2');font-weight:200 800}
*{box-sizing:border-box;margin:0;padding:0}
html{background:${C.process}}
body{font:15px/1.5 Atkinson,system-ui,sans-serif;color:${C.graphite};-webkit-font-smoothing:antialiased;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.titre{font-family:Archivo,'Arial Narrow',sans-serif;font-weight:800;font-stretch:62%;text-transform:uppercase;line-height:.9;letter-spacing:-.005em}
.data,th{font-family:'Martian Mono',ui-monospace,monospace;font-stretch:87.5%;font-weight:500;text-transform:uppercase;letter-spacing:.06em;font-variant-numeric:tabular-nums}
code,.mono{font-family:'Martian Mono',ui-monospace,monospace;font-stretch:87.5%;font-size:.82em}
code{background:${C.process};padding:0 .3em;word-break:break-word}
a{color:inherit;overflow-wrap:anywhere}
.page{max-width:1180px;margin:0 auto;padding:0 24px 64px}
.planche{background-color:${C.process};background-image:linear-gradient(rgba(27,30,33,.075) 1px,transparent 1px),linear-gradient(90deg,rgba(27,30,33,.075) 1px,transparent 1px);background-size:16px 16px}
header.tete{border:1.5px solid ${C.graphite};margin:24px 0 0;padding:28px 28px 24px;display:grid;grid-template-columns:auto minmax(0,1fr) minmax(290px,360px);gap:28px;align-items:center}
header.tete h1{font-size:72px}
header.tete .sous{font-size:18px;margin-top:10px;max-width:640px}
.cartouche{border:1.5px solid ${C.graphite};background:${C.releve};min-width:290px;align-self:end}
.cartouche .l{display:grid;grid-template-columns:1.6fr 1fr}
.cartouche .l+.l{border-top:1px solid ${C.graphite}}
.cartouche .c{padding:4px 8px 6px;min-width:0}
.cartouche .c+.c{border-left:1px solid ${C.graphite}}
.cartouche .k{font-size:7.5px;color:${C.gris}}
.cartouche .v{font-size:11px;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
nav.sommaire{position:sticky;top:0;z-index:5;background:${C.releve};border:1.5px solid ${C.graphite};border-top:0;display:flex;flex-wrap:wrap}
nav.sommaire a{font:500 10px 'Martian Mono',monospace;font-stretch:87.5%;text-transform:uppercase;letter-spacing:.06em;text-decoration:none;padding:9px 12px;border-right:1px solid ${C.graphite}}
nav.sommaire a:hover{background:${C.process}}
section.bloc{background:${C.releve};border:1.5px solid ${C.graphite};margin-top:24px;padding:24px 28px 28px}
h2{font-family:Archivo,sans-serif;font-weight:800;font-stretch:62%;text-transform:uppercase;font-size:36px;line-height:1.05;margin:0 0 14px;display:flex;gap:12px;align-items:baseline}
h2 .n{font-family:'Martian Mono',monospace;font-stretch:87.5%;font-size:11px;font-weight:600;border:1.5px solid ${C.graphite};padding:2px 6px;transform:translateY(-7px)}
h3{font-family:Archivo,sans-serif;font-weight:800;font-stretch:75%;text-transform:uppercase;font-size:19px;margin:22px 0 8px}
h4{font-size:15px;margin:12px 0 6px}
p{margin:0 0 10px;max-width:820px}
ul,ol{margin:0 0 10px 20px}
.note{font-size:12.5px;color:${C.gris}}
.kpi{display:grid;grid-template-columns:repeat(4,1fr);border-top:1.5px solid ${C.graphite};border-bottom:1.5px solid ${C.graphite};margin:4px 0 6px}
.kpi>div{padding:12px 0 12px 14px}
.kpi>div+div{border-left:1px solid ${C.graphite}}
.kpi .v{font-size:56px}
.kpi .k{font-size:9px;margin-top:7px;color:${C.gris}}
table{width:100%;border-collapse:collapse;font-size:13px}
th{font-size:8.5px;text-align:left;padding:6px;border-bottom:1.5px solid ${C.graphite};vertical-align:bottom;color:${C.gris};white-space:nowrap}
td{padding:7px 6px;border-bottom:1px solid rgba(27,30,33,.2);vertical-align:top}
td.num,th.num{text-align:right;font-family:'Martian Mono',monospace;font-stretch:87.5%;font-variant-numeric:tabular-nums;white-space:nowrap}
th[data-tri]{cursor:pointer}
th[data-tri]:hover{color:${C.graphite}}
.chemin{font-family:'Martian Mono',monospace;font-stretch:87.5%;font-size:11.5px;overflow-wrap:anywhere}
td.date{white-space:nowrap}
.sev{display:inline-block;width:28px;height:20px;line-height:17px;text-align:center;font:600 10px 'Martian Mono',monospace;font-stretch:87.5%;border:1px solid ${C.graphite};background:${C.releve}}
.sev.p1{background:${C.graphite};color:${C.releve}}
.sev.p2{border-width:2.5px;line-height:15px}
.sev.p4{border-style:dashed;color:${C.gris}}
.sev.p5{border-color:${C.carter};color:${C.carter}}
.ids{font:500 10px 'Martian Mono',monospace;font-stretch:87.5%;letter-spacing:.03em}
.repartition{display:flex;height:34px;border:1.5px solid ${C.graphite};margin:6px 0 8px}
.repartition>div{display:flex;align-items:center;justify-content:center;font:600 10px 'Martian Mono',monospace;font-stretch:87.5%;overflow:hidden;white-space:nowrap}
.repartition>div+div{border-left:1px solid ${C.graphite}}
.legende{display:flex;gap:18px;flex-wrap:wrap;font-size:12.5px;margin:0 0 6px}
.legende span{display:flex;gap:7px;align-items:center}
.cal{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:0;border-top:1.5px solid ${C.graphite};border-left:1px solid ${C.graphite}}
.cal>div{border-right:1px solid ${C.graphite};border-bottom:1px solid ${C.graphite};padding:12px 14px 14px;min-height:150px;display:flex;flex-direction:column}
.cal .q{font-size:9px;color:${C.gris}}
.cal .a{font-family:Archivo,sans-serif;font-weight:800;font-stretch:75%;text-transform:uppercase;font-size:19px;margin:4px 0 6px}
.cal .d{font-size:12.5px;flex:1}
.cal .v{font-family:Archivo,sans-serif;font-weight:800;font-stretch:62%;font-size:40px;line-height:1;margin-top:10px}
.cal .v small{font:500 9px 'Martian Mono',monospace;font-stretch:87.5%;text-transform:uppercase;letter-spacing:.06em;color:${C.gris};margin-left:6px}
.cal .zero .v{color:${C.carter}}
details.r>summary{cursor:pointer;list-style:none}
details.r>summary::-webkit-details-marker{display:none}
details.r>summary .plus{font:600 11px 'Martian Mono',monospace;margin-left:6px;color:${C.gris}}
details.r[open]>summary .plus{transform:rotate(45deg);display:inline-block}
.explication{margin:6px 0 4px;font-size:13px;max-width:760px}
.explication b{font:600 9px 'Martian Mono',monospace;font-stretch:87.5%;text-transform:uppercase;letter-spacing:.06em;margin-right:4px}
.occ{list-style:none;margin:6px 0 2px;font-size:12px}
.occ li{padding:3px 0;border-top:1px dotted rgba(27,30,33,.3);display:grid;grid-template-columns:36px minmax(0,1fr) minmax(0,1.5fr);gap:12px}
.occ .chemin{overflow-wrap:break-word;font-size:11.5px}
.occ code{background:none;padding:0;color:${C.gris}}
.filtres{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:4px 0 10px}
.filtres input,.filtres select{font:13px Atkinson,sans-serif;border:1.5px solid ${C.graphite};background:#fff;padding:7px 9px;border-radius:0;color:${C.graphite}}
.filtres input{flex:1;min-width:220px}
.revue{border-left:4px solid ${C.graphite};padding:4px 0 4px 18px;margin:8px 0 14px;font-size:15.5px}
.revue h3,.revue h4{margin-top:12px}
.encart{border:1.5px dashed ${C.graphite};padding:12px 16px;margin:12px 0 0;font-size:13.5px}
footer{margin-top:24px;display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;font-size:9px;color:${C.gris};border-top:1.5px solid ${C.graphite};padding-top:8px}
@media screen and (max-width:860px){header.tete{grid-template-columns:1fr}header.tete h1{font-size:52px}.kpi{grid-template-columns:1fr 1fr}.kpi>div:nth-child(3){border-left:0}.kpi .v{font-size:42px}.cartouche{min-width:0}.page{padding:0 12px 48px}section.bloc{padding:18px 16px}.occ li{grid-template-columns:40px 1fr}.occ li code{grid-column:1/-1}.masquable{display:none}table{display:block;overflow-x:auto}}
@media print{html,body{background:#fff}header.tete{grid-template-columns:auto minmax(0,1fr)}header.tete h1{font-size:60px}header.tete .cartouche{grid-column:1/-1}nav.sommaire,.filtres,.masquable,.ecran{display:none}section.bloc{break-inside:auto;border:0;padding:0;margin-top:18px}header.tete{margin-top:0}details.r>summary .plus{display:none}tr,.kpi,.kpi>div,.cal>div,.legende,.barre,.revue p,.revue li{break-inside:avoid}h2,h3{break-after:avoid}}
`;

function bulle(d = 112) {
  const sw = Math.max(2, d / 40), r = d / 2 - sw / 2;
  return `<svg width="${d}" height="${d}" viewBox="0 0 ${d} ${d}" role="img" aria-label="Repère DEV-607"><circle cx="${d / 2}" cy="${d / 2}" r="${r}" fill="${C.signal}" stroke="${C.graphite}" stroke-width="${sw}"/><line x1="${sw / 2}" y1="${d / 2}" x2="${d - sw / 2}" y2="${d / 2}" stroke="${C.graphite}" stroke-width="${sw}"/><text x="${d / 2}" y="${d * 0.4}" text-anchor="middle" font-family="Archivo" font-stretch="62%" font-weight="800" font-size="${d * 0.3}" fill="${C.graphite}">DEV</text><text x="${d / 2}" y="${d * 0.76}" text-anchor="middle" font-family="Martian Mono" font-stretch="87.5%" font-weight="500" font-size="${d * 0.2}" fill="${C.graphite}">607</text></svg>`;
}

const sev = (p) => `<span class="sev p${p}" title="${p <= 4 ? e(PRIORITES[p]) : 'Aucun constat'}">${p <= 4 ? `P${p}` : '—'}</span>`;
const cartouche = (lignes) => `<div class="cartouche">${lignes.map((l) => `<div class="l" style="grid-template-columns:${l.map((c) => `${c.f ?? 1}fr`).join(' ')}">${l.map((c) => `<div class="c"><div class="data k">${e(c.k)}</div><div class="data v" title="${e(c.v)}">${e(c.v)}</div></div>`).join('')}</div>`).join('')}</div>`;

/**
 * @param inv inventaire (inventaire.json)
 * @param syn synthese(inv)
 * @param opts { pro: { estimation, dossiers, doublons, libreoffice, comparaison }, revue: markdown, organisation, limite }
 */
const abrege = (u) => { const s = u.replace(/^https:\/\//, '').replace(/\/$/, ''); return s.length > 72 ? `${s.slice(0, 40)}…${s.slice(-28)}` : s; };

export function rapportHtml(inv, syn, opts = {} ) {
  const pro = opts.pro ?? null;
  const revue = opts.revue ? sections(opts.revue) : null;
  const org = opts.organisation || (inv.anonymise ? 'Parc analysé' : inv.racines.map((r) => r.split(/[\\/]/).filter(Boolean).pop() ?? r).join(', '));
  const nb = (regles) => {
    const fichiers = new Set();
    let occ = 0;
    for (const f of inv.fichiers) for (const c of f.constats) if (regles.includes(c.regle)) { fichiers.add(f.chemin); occ++; }
    return { fichiers: fichiers.size, occ };
  };
  const fichiersMacros = inv.fichiers.filter((f) => f.macros || f.constats.length);
  const totalOffice = syn.fichiersAnalyses + syn.fichiersSansMacroComptes + syn.nonAnalyses;
  const blocs = [];
  const sommaire = [];
  let n = 0;
  const bloc = (id, titre, contenu) => { sommaire.push([id, titre]); blocs.push(`<section class="bloc" id="${id}"><h2><span class="n">${++n}</span>${e(titre)}</h2>${contenu}</section>`); };

  // 1. Synthèse
  const rep = [1, 2, 3, 4, 5].map((p) => ({ p, v: fichiersMacros.filter((f) => syn.prioriteFichier[f.chemin] === p).length }));
  const totalRep = rep.reduce((a, b) => a + b.v, 0) || 1;
  const styleRep = { 1: `background:${C.graphite};color:${C.releve}`, 2: `background:${C.gris};color:${C.releve}`, 3: `background:${C.carter}`, 4: `background:${C.process}`, 5: `background:${C.releve};color:${C.gris}` };
  const libRep = { 1: 'P1 casse', 2: 'P2 avant 2027', 3: 'P3 à planifier', 4: 'P4 information', 5: 'Aucun constat' };
  bloc('synthese', 'Synthèse', `
    ${revue?.['synthèse'] ? `<div class="revue">${versHtml(revue['synthèse'])}</div>` : `<p style="font-size:17px">${nf(syn.fichiersAvecMacros)} fichiers contiennent du code VBA, soit ${nf(syn.lignesCode)} lignes dans ${nf(syn.modules)} modules. <strong>${nf(syn.fichiersP1)} fichier${syn.fichiersP1 > 1 ? 's cassent' : ' casse'}</strong> à une échéance connue (P1) et ${nf(syn.fichiersP2)} demande${syn.fichiersP2 > 1 ? 'nt' : ''} une action avant 2027 (P2).</p>`}
    <div class="kpi">
      <div><div class="titre v">${nf(syn.fichiersAvecMacros)}</div><div class="data k">Fichiers avec macros</div></div>
      <div><div class="titre v">${nf(syn.lignesCode)}</div><div class="data k">Lignes de code VBA</div></div>
      <div><div class="titre v">${nf(syn.fichiersP1)}</div><div class="data k">Fichiers P1 : cassent</div></div>
      <div><div class="titre v">${nf(syn.fichiersP2)}</div><div class="data k">Fichiers P2 : avant 2027</div></div>
    </div>
    ${syn.vb6.projets || syn.courriel.fichiers ? `<p>${syn.vb6.projets ? `S'y ajoute${syn.vb6.projets > 1 ? 'nt' : ''} ${nf(syn.vb6.projets)} application${syn.vb6.projets > 1 ? 's' : ''} Visual Basic 6 (${nf(syn.vb6.lignes)} lignes). ` : ''}${syn.courriel.fichiers ? `${nf(syn.courriel.fichiers)} fichier${syn.courriel.fichiers > 1 ? 's envoient' : ' envoie'} des courriels par Outlook ou par SMTP : voir <a href="#courriel">Envoi de courriels</a>.` : ''}</p>` : ''}
    <p class="note">${nf(totalOffice)} fichiers trouvés : ${nf(syn.fichiersAnalyses)} analysés, ${nf(syn.fichiersSansMacroComptes)} dans un format qui ne peut pas contenir de macros (.xlsx, .docx…)${syn.nonAnalyses ? `, ${nf(syn.nonAnalyses)} non analysés (limite de l'édition gratuite)` : ''}.</p>
    <h3>Fichiers par priorité la plus haute</h3>
    <div class="repartition" role="img" aria-label="Répartition des fichiers par priorité">${rep.filter((r) => r.v).map((r) => `<div style="flex:${r.v};${styleRep[r.p]}" title="${libRep[r.p]} : ${r.v}">${r.v / totalRep > 0.06 ? nf(r.v) : ''}</div>`).join('')}</div>
    <div class="legende">${rep.map((r) => `<span>${sev(r.p)} ${libRep[r.p]} : <b>${nf(r.v)}</b></span>`).join('')}</div>
    ${opts.limite && syn.nonAnalyses ? `<div class="encart"><strong>Édition gratuite :</strong> l'inventaire s'est arrêté à ${nf(opts.limite)} fichiers ; ${nf(syn.nonAnalyses)} autres sont listés en fin de rapport sans être analysés. La licence de parc lève la limite et ajoute l'export Excel, l'estimation de migration, la répartition par dossier, les copies d'un même projet et la comparaison entre deux inventaires.</div>` : ''}`);

  // 2. Calendrier
  bloc('echeances', 'Ce qui casse, et quand', `
    <p>Chaque case compte les fichiers concernés par une échéance publiée par Microsoft. Les dates sont celles annoncées au ${e(date(inv.date))} ; les sources sont dans le catalogue des règles.</p>
    <div class="cal">${ECHEANCES.map((x) => { const c = nb(x.regles); return `<div class="${c.fichiers ? '' : 'zero'}"><div class="data q">${e(x.quand)}</div><div class="a">${e(x.axe)}</div><div class="d">${e(x.detail)}</div><div class="v">${nf(c.fichiers)}<small>fichier${c.fichiers > 1 ? 's' : ''} · ${nf(c.occ)} occurrence${c.occ > 1 ? 's' : ''}</small></div></div>`; }).join('')}</div>`);

  // 3. Constats par règle
  const lignesRegles = Object.entries(syn.parRegle).map(([id, v]) => {
    const r = REGLES[id];
    const ex = inv.fichiers.flatMap((f) => f.constats.filter((c) => c.regle === id).map((c) => ({ f, c }))).slice(0, 12);
    const remarque = revue?.['remarques par constat'] ? extraireRemarque(revue['remarques par constat'], id) : '';
    return `<tr><td>${sev(r.p)}</td><td><details class="r"><summary><span class="ids">${id}</span> <strong>${e(r.titre)}</strong><span class="plus">+</span></summary>
      <div class="explication"><p><b>Pourquoi</b>${e(r.pourquoi)}</p><p><b>Quoi faire</b>${e(r.quoi)}</p>${remarque ? `<p><b>Avis de l'expert</b>${enLigne(remarque)}</p>` : ''}${r.source ? `<p class="note"><b>Source</b><a href="${e(r.source)}">${e(abrege(r.source))}</a></p>` : ''}</div>
      <ul class="occ">${ex.map(({ f, c }) => `<li><span class="mono">${c.ligne || '—'}</span><span class="chemin" title="${e(f.chemin)}">${e(court(f.chemin))}${c.module && !c.module.startsWith('(') ? ` · ${e(c.module)}` : ''}</span><code>${e(c.extrait)}</code></li>`).join('')}${v.occurrences > ex.length ? `<li><span></span><span class="note">… et ${nf(v.occurrences - ex.length)} autres</span></li>` : ''}</ul></details></td>
      <td class="data masquable" style="font-size:9px">${e(r.axe)}</td><td class="num">${nf(v.fichiers)}</td><td class="num">${nf(v.occurrences)}</td></tr>`;
  }).join('');
  bloc('regles', 'Constats par règle', `
    <p>Une ligne par règle déclenchée, de la plus urgente à la simple information. Ouvrez une ligne pour lire pourquoi c'est un risque, quoi faire, et les premiers emplacements (fichier, module, numéro de ligne dans l'éditeur VBA).</p>
    ${lignesRegles ? `<table><thead><tr><th style="width:40px">Prio.</th><th>Règle</th><th class="masquable" style="width:120px">Axe</th><th class="num" style="width:70px">Fichiers</th><th class="num" style="width:90px">Occurrences</th></tr></thead><tbody>${lignesRegles}</tbody></table>` : '<p>Aucun constat : aucune règle ne s\'est déclenchée sur ce parc.</p>'}`);

  // 4. Signatures et références
  const certs = new Map();
  for (const f of inv.fichiers) for (const c of f.signature.certificats.filter((x) => !x.autorite)) {
    const k = c.empreinte;
    if (!certs.has(k)) certs.set(k, { ...c, fichiers: 0 });
    certs.get(k).fichiers++;
  }
  const refs = new Map();
  for (const f of inv.fichiers) for (const p of f.projets) for (const r of p.references) {
    if (/^(stdole|Office|Excel|Word|PowerPoint|Access|VBA)$/i.test(r.nom)) continue;
    const k = `${r.nom}|${r.description || r.chemin}`;
    if (!refs.has(k)) refs.set(k, { ...r, fichiers: new Set() });
    refs.get(k).fichiers.add(f.chemin);
  }
  const com = new Map();
  for (const f of inv.fichiers) for (const x of f.progIds ?? []) {
    const k = x.toLowerCase();
    if (!com.has(k)) com.set(k, { nom: x, fichiers: 0 });
    com.get(k).fichiers++;
  }
  const s = syn.signatures;
  bloc('signatures', 'Signatures et références', `
    <h3>Signatures des projets</h3>
    <table style="max-width:640px"><tbody>
      <tr><td>Signés en V3</td><td class="num">${nf(s.V3)}</td></tr>
      <tr><td>Signés sans V3 (agile ou ancienne) : <strong>bloqués dès que la stratégie V3 est activée</strong></td><td class="num">${nf(s.agile + s.ancienne)}</td></tr>
      <tr><td>Non signés : dépendent des emplacements approuvés et de la marque du Web</td><td class="num">${nf(s.aucune)}</td></tr>
    </tbody></table>
    ${certs.size ? `<h3>Certificats</h3><table><thead><tr><th>Signataire</th><th>Émetteur</th><th>Fin de validité</th><th class="num">Fichiers</th></tr></thead><tbody>${[...certs.values()].map((c) => `<tr><td>${e(c.sujet)}</td><td>${e(c.autosigne ? 'autosigné' : c.emetteur)}</td><td class="mono">${e(date(c.fin))}${new Date(c.fin) < new Date(inv.date) ? ' <strong>expiré</strong>' : ''}</td><td class="num">${nf(c.fichiers)}</td></tr>`).join('')}</tbody></table>` : ''}
    <h3>Bibliothèques référencées</h3>
    ${refs.size ? `<table><thead><tr><th>Référence</th><th>Description ou chemin</th><th class="num">Fichiers</th></tr></thead><tbody>${[...refs.values()].sort((a, b) => b.fichiers.size - a.fichiers.size).map((r) => `<tr><td>${e(r.nom)}</td><td class="chemin">${e(r.description || r.chemin || r.guid)}</td><td class="num">${nf(r.fichiers.size)}</td></tr>`).join('')}</tbody></table>` : '<p class="note">Aucune référence en dehors des bibliothèques Office standard.</p>'}
    <h3>Objets COM créés par le code</h3>
    <p class="note">ProgID passés à CreateObject ou GetObject : chacun est une dépendance du poste (application, pilote, bibliothèque Windows).</p>
    ${com.size ? `<table style="max-width:640px"><thead><tr><th>ProgID</th><th class="num">Fichiers</th></tr></thead><tbody>${[...com.values()].sort((a, b) => b.fichiers - a.fichiers || a.nom.localeCompare(b.nom)).map((x) => `<tr><td class="mono">${e(x.nom)}</td><td class="num">${nf(x.fichiers)}</td></tr>`).join('')}</tbody></table>` : '<p class="note">Aucun.</p>'}`);

  // Envoi de courriels : chaque point d'envoi, comment il part, et la cible recommandée.
  const prioCourriel = (f) => Math.min(...f.courriel.mecanismes.map((m) => REGLES[m].p));
  const envois = inv.fichiers.filter((f) => f.courriel).sort((a, b) => prioCourriel(a) - prioCourriel(b) || a.chemin.localeCompare(b.chemin, 'fr'));
  if (envois.length) {
    const o = envois.filter((f) => f.courriel.mecanismes.some((m) => /^OUT-/.test(m))).length;
    const m1 = envois.filter((f) => f.courriel.mecanismes.includes('MEL-001')).length;
    const usages = (c) => [[c.envoiDirect, 'envoi direct'], [c.affichage, 'affiché avant envoi'], [c.piecesJointes, 'pièces jointes'], [c.lectureBoite, 'lit la boîte']].filter(([n]) => n).map(([n, t]) => `${t}${n > 1 ? ` ×${n}` : ''}`).join(' · ') || '—';
    bloc('courriel', 'Envoi de courriels', `
      <p>${nf(envois.length)} fichier${envois.length > 1 ? 's envoient' : ' envoie'} des courriels. ${o ? `<strong>${nf(o)}</strong> pilote${o > 1 ? 'nt' : ''} Outlook (COM, MAPI ou SendObject) et cesser${o > 1 ? 'ont' : 'a'} d'envoyer sur un poste passé au nouvel Outlook. ` : ''}${m1 ? `<strong>${nf(m1)}</strong> envoie${m1 > 1 ? 'nt' : ''} en SMTP avec un mot de passe vers Exchange Online, désactivé par défaut fin décembre 2026.` : ''}</p>
      <table><thead><tr><th style="width:40px">Prio.</th><th>Fichier</th><th>Mécanisme</th><th class="masquable">Usage</th><th>Cible recommandée</th></tr></thead><tbody>${envois.map((f) => {
        const p = prioCourriel(f);
        return `<tr><td>${sev(p)}</td><td class="chemin">${e(f.chemin)}</td><td class="ids">${f.courriel.mecanismes.join(' ')}</td><td class="masquable" style="font-size:12.5px">${e(usages(f.courriel))}</td><td style="font-size:12.5px">${e(cibleCourriel(f))}</td></tr>`;
      }).join('')}</tbody></table>
      <div class="encart"><strong>Remédiation :</strong> le dossier <code>remediation</code> fourni avec l'outil contient le module <code>ModGraphMail.bas</code> (VBA et VB6, sans DLL), qui remplace l'envoi par Outlook ou par SMTP par Microsoft Graph, et le guide de l'administrateur Microsoft 365 qui limite l'accès à une seule boîte. La remédiation peut aussi être faite pour vous, à prix fixe et entièrement à distance.</div>`);
  }

  // Applications Visual Basic 6
  const vb6 = inv.fichiers.filter((f) => f.famille === 'vb6');
  if (vb6.length) {
    bloc('vb6', 'Applications Visual Basic 6', `
      <p>${nf(vb6.length)} projet${vb6.length > 1 ? 's' : ''} VB6, ${nf(syn.vb6.lignes)} lignes de code. Leur moteur d'exécution reste pris en charge par Microsoft sur Windows 11 et Server 2025, en 32 bits ; ce qui casse vient de leurs dépendances : VBScript, Outlook, composants hors support.</p>
      <table><thead><tr><th style="width:40px">Prio.</th><th>Projet</th><th class="num">Lignes</th><th class="num masquable">Formulaires</th><th class="masquable">Composants</th><th>Règles</th></tr></thead><tbody>${vb6.map((f) => {
        const p = f.projets[0] ?? {};
        const mods = p.modules ?? [];
        const ids = [...new Set(f.constats.map((c) => c.regle))].filter((x) => x !== 'VB6-001');
        return `<tr><td>${sev(syn.prioriteFichier[f.chemin])}</td><td><div class="chemin">${e(f.chemin)}</div><div class="note">${e([p.vb6?.type, p.vb6?.exe, `${nf(mods.length)} fichiers sources`].filter(Boolean).join(' · '))}${p.vb6?.manquants?.length ? ` · <strong>${nf(p.vb6.manquants.length)} introuvable${p.vb6.manquants.length > 1 ? 's' : ''}</strong>` : ''}</div></td><td class="num">${nf(f.lignesCode)}</td><td class="num masquable">${nf(mods.filter((m) => m.type === 'formulaire').length)}</td><td class="mono masquable" style="font-size:10.5px">${e((p.objets ?? []).map((o) => o.fichier).join(', ') || '—')}</td><td class="ids">${ids.join(' ') || '—'}</td></tr>`;
      }).join('')}</tbody></table>`);
  }

  // 5. Modules Pro
  if (pro?.dossiers) bloc('dossiers', 'Par dossier', pro.dossiers);
  if (pro?.estimation) bloc('estimation', 'Estimation de migration', pro.estimation);
  if (pro?.doublons) bloc('copies', 'Copies d\'un même projet', pro.doublons);
  if (pro?.libreoffice) bloc('libreoffice', 'Compatibilité LibreOffice', pro.libreoffice);
  if (pro?.comparaison) bloc('evolution', 'Évolution depuis l\'inventaire précédent', pro.comparaison);
  if (revue?.["plan d'action"]) bloc('plan', 'Plan d\'action', `<div class="revue">${versHtml(revue["plan d'action"])}</div>`);

  // 6. Fichiers
  const auteur = (f) => f.meta.dernierAuteur || f.meta.auteur || '';
  const lignesF = fichiersMacros.map((f) => {
    const p = syn.prioriteFichier[f.chemin];
    const ids = [...new Set(f.constats.map((c) => c.regle))];
    const mods = f.projets.reduce((t, x) => t + x.modules.filter((m) => m.lignesCode > 0).length, 0);
    const est = pro?.parFichier?.[f.chemin];
    return `<tr data-p="${p}" data-t="${e((f.chemin + ' ' + ids.join(' ') + ' ' + auteur(f)).toLowerCase())}"><td data-v="${p}">${sev(p)}</td>
      <td><details class="r"><summary><span class="chemin">${e(f.chemin)}</span><span class="plus">+</span></summary>
      <ul class="occ">${f.constats.slice(0, 60).map((c) => `<li><span class="ids">${c.regle}</span><span class="note">${c.ligne ? `${e(c.module)} : ${c.ligne}` : e(c.module)}</span><code>${e(c.extrait || REGLES[c.regle].titre)}</code></li>`).join('')}${f.constats.length > 60 ? `<li><span></span><span class="note">… ${f.constats.length - 60} de plus dans l'export</span></li>` : ''}${f.erreurs.length ? `<li><span class="ids">!</span><span class="note">lecture</span><code>${e(f.erreurs.join(' ; '))}</code></li>` : ''}</ul></details></td>
      <td class="ids masquable">${ids.slice(0, 6).join(' ')}${ids.length > 6 ? ' …' : ''}</td>
      <td class="num" data-v="${f.lignesCode}">${nf(f.lignesCode)}</td><td class="num masquable" data-v="${mods}">${nf(mods)}</td>
      <td class="data masquable" style="font-size:9px">${e(f.signature.etat)}</td>
      ${est ? `<td class="num masquable" data-v="${est.corriger}">${nf(est.corriger, 1)}</td>` : ''}
      <td class="mono masquable date" data-v="${e(f.meta.modifie || f.modifieDisque || '')}">${e(date(f.meta.modifie || f.modifieDisque))}</td><td class="masquable" style="font-size:12px">${e(auteur(f))}</td></tr>`;
  }).join('');
  bloc('fichiers', 'Fichiers', `
    <p>${nf(fichiersMacros.length)} fichiers avec du code VBA ou un constat. <span class="ecran">Cliquez un en-tête pour trier, un chemin pour voir ses constats.</span></p>
    <div class="filtres"><input type="search" id="filtre" placeholder="Filtrer : chemin, règle (B64-001), auteur…" aria-label="Filtrer les fichiers"><select id="prio" aria-label="Priorité"><option value="">Toutes priorités</option><option value="1">P1 seulement</option><option value="2">P1 et P2</option><option value="3">P1 à P3</option></select><span class="note" id="compte"></span></div>
    <table id="tf"><thead><tr><th data-tri="n" style="width:40px">Prio.</th><th data-tri="s">Fichier</th><th class="masquable">Règles</th><th class="num" data-tri="n">Lignes</th><th class="num masquable" data-tri="n">Modules</th><th class="masquable">Signature</th>${pro?.parFichier ? '<th class="num masquable" data-tri="n">Correction (h)</th>' : ''}<th class="masquable" data-tri="s">Modifié</th><th class="masquable" data-tri="s">Auteur</th></tr></thead><tbody>${lignesF}</tbody></table>`);

  // 7. Hors inventaire
  const hors = [];
  if (inv.nonAnalyses.length) hors.push(`<h3>Non analysés (édition gratuite) : ${nf(inv.nonAnalyses.length)}</h3><ul class="occ">${inv.nonAnalyses.slice(0, 500).map((c) => `<li><span></span><span class="chemin" style="grid-column:2/-1">${e(c)}</span></li>`).join('')}</ul>${inv.nonAnalyses.length > 500 ? `<p class="note">… et ${nf(inv.nonAnalyses.length - 500)} autres (liste complète dans inventaire.json).</p>` : ''}`);
  if (inv.inaccessibles.length) hors.push(`<h3>Inaccessibles : ${nf(inv.inaccessibles.length)}</h3><p class="note">Droits insuffisants ou fichier verrouillé : relancer avec un compte qui a la lecture sur ces dossiers.</p><ul class="occ">${inv.inaccessibles.slice(0, 200).map((x) => `<li><span class="ids">${e(x.raison)}</span><span class="chemin" style="grid-column:2/-1">${e(x.chemin)}</span></li>`).join('')}</ul>`);
  if (inv.tropGros.length) hors.push(`<h3>Ignorés car trop volumineux : ${nf(inv.tropGros.length)}</h3><ul class="occ">${inv.tropGros.map((x) => `<li><span class="mono">${nf(x.taille / 1048576, 0)} Mo</span><span class="chemin" style="grid-column:2/-1">${e(x.chemin)}</span></li>`).join('')}</ul>`);
  if (hors.length) bloc('hors', 'Hors inventaire', hors.join(''));

  // 8. Méthode
  bloc('methode', 'Méthode et limites', `
    <p>L'outil lit chaque fichier en lecture seule, sur ce poste : il décompresse le projet VBA (format MS-OVBA), lit les références, la signature et les métadonnées, puis applique ${Object.keys(REGLES).length} règles au code. Il n'ouvre pas Office, n'exécute aucune macro et n'envoie rien sur le réseau.</p>
    <ul><li>Les numéros de ligne sont ceux de l'éditeur VBA (en-têtes « Attribute » exclus).</li>
    <li>Un Declare sans PtrSafe placé dans la branche #Else d'un « #If VBA7 » n'est pas compté : c'est le motif recommandé par Microsoft.</li>
    <li>B64-002 est une heuristique sur les noms (hWnd, lpBuffer…) : à confirmer à la lecture.</li>
    <li>Le code des bases Access n'est lisible qu'avec Access : utilisez <code>outils\\exporter-access.ps1</code> puis relancez l'inventaire sur le dossier exporté.</li>
    <li>La validité cryptographique des signatures n'est pas vérifiée : seuls leur format (V3 ou non) et les dates du certificat le sont.</li>
    <li>Les projets Visual Basic 6 sont lus à partir de leur fichier .vbp et de ses sources (.frm, .bas, .cls, .ctl). Un exécutable compilé sans ses sources n'est pas analysé.</li>
    <li>Les sources .NET (.cs, .vb, projets) et les fichiers .config et .ini ne sont lus que pour l'envoi de courriels, et n'apparaissent ici que s'ils déclenchent une règle${inv.annexesLues && Object.keys(inv.annexesLues).length ? ` (${Object.entries(inv.annexesLues).map(([k, v]) => `${nf(v)} ${k === 'dotnet' ? 'sources .NET' : 'configurations'}`).join(', ')} lus)` : ''}.</li></ul>
    <h3>Catalogue des règles</h3>
    <table><thead><tr><th style="width:40px">Prio.</th><th style="width:72px">Règle</th><th>Titre et source</th></tr></thead><tbody>${Object.entries(REGLES).map(([id, r]) => `<tr><td>${sev(r.p)}</td><td class="ids">${id}</td><td>${e(r.titre)}${r.source ? ` · <a class="note" href="${e(r.source)}">source</a>` : ''}</td></tr>`).join('')}</tbody></table>`);

  const titre = `Inventaire des macros VBA — ${org}`;
  return `<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${e(titre)}</title><style>${CSS}</style></head>
<body><div class="page">
<header class="tete planche">
  ${bulle(112)}
  <div><div class="data note">Inventaire des macros · ${e(date(inv.date))}</div><h1 class="titre">Macros VBA<br>du parc</h1><p class="sous">${e(org)} : ce qui casse avec Office 64 bits, les signatures V3, le nouvel Outlook, la fin de SMTP par mot de passe et celle de VBScript, fichier par fichier.</p></div>
  ${cartouche([[{ k: 'Désignation', v: 'Inventaire VBA' }, { k: 'Repère', v: 'DEV-607' }], [{ k: 'Périmètre', v: org }, { k: 'Date', v: date(inv.date) }], [{ k: 'Outil', v: `inventaire-vba ${inv.outil.version}` }, { k: 'Édition', v: inv.outil.edition === 'parc' ? 'parc' : 'gratuite' }]])}
</header>
<nav class="sommaire" aria-label="Sommaire">${sommaire.map(([id, t]) => `<a href="#${id}">${e(t)}</a>`).join('')}</nav>
${blocs.join('\n')}
<footer class="data"><span>inventaire-vba ${e(inv.outil.version)} · édition ${inv.outil.edition === 'parc' ? 'parc' : 'gratuite'}${inv.anonymise ? ' · données anonymisées' : ''}</span><span>Analyse locale, en lecture seule : aucun fichier n'a quitté ce poste</span><span>DEV-607 · Consigne</span></footer>
</div>
<script>
(function(){
  var t=document.getElementById('tf'); if(!t) return;
  var corps=t.tBodies[0], lignes=[].slice.call(corps.rows), f=document.getElementById('filtre'), p=document.getElementById('prio'), compte=document.getElementById('compte');
  function filtrer(){ var q=f.value.trim().toLowerCase(), m=+p.value||9, n=0;
    lignes.forEach(function(r){ var ok=(!q||r.dataset.t.indexOf(q)>=0)&&(+r.dataset.p<=m); r.style.display=ok?'':'none'; if(ok)n++; });
    compte.textContent=n+' / '+lignes.length+' fichiers'; }
  f.addEventListener('input',filtrer); p.addEventListener('change',filtrer); filtrer();
  [].forEach.call(t.tHead.rows[0].cells,function(th,i){ if(!th.dataset.tri) return; var sens=1;
    th.addEventListener('click',function(){ sens=-sens; var num=th.dataset.tri==='n';
      lignes.sort(function(a,b){ var x=a.cells[i].dataset.v||a.cells[i].textContent, y=b.cells[i].dataset.v||b.cells[i].textContent;
        return (num?(parseFloat(x)||0)-(parseFloat(y)||0):x.localeCompare(y,'fr'))*sens; });
      lignes.forEach(function(r){ corps.appendChild(r); }); }); });
})();
</script>
</body></html>`;
}

export function cibleCourriel(f) {
  const c = f.courriel;
  if (c.mecanismes.length === 1 && c.mecanismes[0] === 'MEL-002') return 'À vérifier : serveur SMTP et mode d\'authentification';
  if (c.lectureBoite) return 'Graph avec lecture de la boîte (Mail.Read) : sur devis';
  if (f.famille === 'dotnet' || f.famille === 'config') return 'Graph depuis .NET (SDK Microsoft Graph)';
  if (c.affichage) return 'Graph : brouillon ouvert dans Outlook sur le web (CreerBrouillon)';
  return 'Graph : EnvoyerCourriel (ModGraphMail)';
}

const court = (c) => (c.length > 64 ? `…${c.slice(-62)}` : c);

// Dans la section « Remarques par constat » de la revue : une ligne « - B64-001 : texte ».
function extraireRemarque(md, id) {
  const l = md.split(/\r?\n/).find((x) => new RegExp(`^\\s*[-*]?\\s*\\**${id}\\**\\s*[:：]`).test(x));
  return l ? l.replace(new RegExp(`^\\s*[-*]?\\s*\\**${id}\\**\\s*[:：]\\s*`), '') : '';
}
