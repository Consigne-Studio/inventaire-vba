// Parcours d'un dossier ou d'un partage, analyse de chaque fichier, et assemblage de l'inventaire.
import { readdir, readFile, stat } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { analyserOffice, familleDe, SANS_MACRO, ANNEXES, lireVbp, decoderTexte } from './format/office.mjs';
import { analyserModule, constatsProjet, constatsFichier, analyserTexteAnnexe, REGLES } from './regles.mjs';
import { VERSION } from './version.mjs';

export const LIMITE_GRATUITE = 200;
const IGNORES = new Set(['$recycle.bin', 'system volume information', 'node_modules', '.git', '.svn', '$windows.~bt', 'windows', 'program files', 'program files (x86)', 'appdata', 'bin', 'obj', 'packages', '.vs']);
// Règles qui décrivent un point d'envoi de courriel.
export const REGLES_COURRIEL = ['MEL-001', 'OUT-001', 'OUT-002', 'OUT-003', 'MEL-002'];

/** Liste les fichiers à analyser et compte les formats sans macros. */
export async function lister(racines, { maxTaille = 200 * 1024 * 1024, exclure = [] } = {}) {
  const fichiers = [];
  const sansMacro = {};
  const inaccessibles = [];
  const tropGros = [];
  const exclusions = exclure.map((e) => e.toLowerCase());
  const parcourir = async (dossier, racine) => {
    let entrees;
    try { entrees = await readdir(dossier, { withFileTypes: true }); } catch (e) { inaccessibles.push({ chemin: dossier, raison: e.code ?? e.message }); return; }
    entrees.sort((a, b) => a.name.localeCompare(b.name, 'fr'));
    for (const e of entrees) {
      const plein = path.join(dossier, e.name);
      const bas = e.name.toLowerCase();
      if (exclusions.some((x) => plein.toLowerCase().includes(x))) continue;
      if (e.isDirectory()) { if (!IGNORES.has(bas)) await parcourir(plein, racine); continue; }
      if (!e.isFile() || bas.startsWith('~$')) continue;
      const ext = path.extname(bas).slice(1);
      if (SANS_MACRO.includes(ext)) { sansMacro[ext] = (sansMacro[ext] ?? 0) + 1; continue; }
      if (!familleDe(bas)) continue;
      let st;
      try { st = await stat(plein); } catch (err) { inaccessibles.push({ chemin: plein, raison: err.code ?? err.message }); continue; }
      if (st.size > maxTaille) { tropGros.push({ chemin: plein, taille: st.size }); continue; }
      fichiers.push({ plein, racine, taille: st.size, modifie: st.mtime.toISOString() });
    }
  };
  for (const r of racines) {
    const abs = path.resolve(r);
    const st = await stat(abs);
    if (st.isFile()) fichiers.push({ plein: abs, racine: path.dirname(abs), taille: st.size, modifie: st.mtime.toISOString() });
    else await parcourir(abs, abs);
  }
  return { fichiers, sansMacro, inaccessibles, tropGros };
}

const empreinte = (s) => createHash('sha256').update(s).digest('hex').slice(0, 16);
// Code normalisé pour reconnaître les copies : sans attributs, commentaires ni espaces de fin.
const normaliser = (code) => code.split(/\r\n|\r|\n/).filter((l) => !/^\s*(Attribute VB_|'|$)/i.test(l)).map((l) => l.trimEnd()).join('\n');

/** Analyse un fichier lu en mémoire. Ne lève jamais : les erreurs deviennent un constat FMT-003. */
export function analyserFichier(rel, buf, { maintenant = new Date(), annexes } = {}) {
  let f;
  try {
    f = analyserOffice(rel, buf, { annexes });
  } catch (e) {
    f = { famille: familleDe(rel), format: 'inconnu', projets: [], signature: { legacy: false, agile: false, v3: false, certificats: [] }, meta: {}, activeX: [], erreurs: [e.message] };
  }
  const constats = [];
  const projets = [];
  let auto = false;
  const progIds = new Set();
  const courriel = { envoiDirect: 0, affichage: 0, piecesJointes: 0, lectureBoite: 0 };
  const cumuler = (c) => { for (const k of Object.keys(courriel)) courriel[k] += c[k]; };
  if (f.texte != null) {
    const a = analyserTexteAnnexe(f.texte, f.famille);
    constats.push(...a.constats);
    cumuler(a.courriel);
  }
  for (const p of f.projets) {
    const { constats: cp, refVbs } = constatsProjet(p);
    constats.push(...cp);
    const modules = [];
    for (const m of p.modules) {
      const a = analyserModule(m, { langage: p.langage });
      cumuler(a.courriel);
      auto ||= a.auto;
      for (const x of a.progIds) progIds.add(x);
      for (const c of a.constats) constats.push({ ...c, module: m.nom });
      if (a.regexNatif && !refVbs) constats.push({ regle: 'VBS-004', module: m.nom, ligne: a.regexNatif.ligne, extrait: a.regexNatif.brut.slice(0, 160) });
      modules.push({ nom: m.nom, type: m.type, lignes: a.lignes, lignesCode: a.lignesCode, procedures: a.procedures, declares: a.declares, empreinte: empreinte(normaliser(m.code)) });
    }
    projets.push({ nom: p.nom, ...(p.vb6 ? { langage: 'vb6', vb6: p.vb6, objets: p.objets } : {}), protege: p.protege, references: p.references.map(({ nom, guid, version, chemin, description, type }) => ({ nom, guid, version, chemin, description, type })), modules, erreurs: p.erreurs });
    f.erreurs.push(...p.erreurs);
  }
  constats.push(...constatsFichier(f, maintenant));
  const regles = new Set(constats.map((c) => c.regle));
  if (auto && regles.has('SEC-001') && regles.has('SEC-002')) constats.push({ regle: 'SEC-010', module: '(projet)', ligne: 0, extrait: 'ouverture + téléchargement + commande' });
  const modulesCode = projets.flatMap((p) => p.modules).filter((m) => m.lignesCode > 0);
  if (f.erreurs.length && !f.projets.length && f.famille !== 'access') constats.push({ regle: 'FMT-003', module: '(fichier)', ligne: 0, extrait: f.erreurs[0].slice(0, 160) });
  const s = f.signature;
  const mecanismes = REGLES_COURRIEL.filter((id) => constats.some((c) => c.regle === id));
  return {
    chemin: rel,
    famille: f.famille,
    extension: path.extname(rel).slice(1).toLowerCase(),
    format: f.format,
    ...(mecanismes.length ? { courriel: { mecanismes, ...courriel } } : {}),
    macros: modulesCode.length > 0 || f.xlm || (f.famille === 'access' && f.accessVbaProbable) || f.famille === 'script',
    meta: f.meta ?? {},
    signature: { etat: s.v3 ? 'V3' : s.agile ? 'agile' : s.legacy ? 'ancienne' : 'aucune', certificats: s.certificats },
    projets,
    progIds: [...progIds].sort((a, b) => a.localeCompare(b)),
    lignesCode: modulesCode.reduce((t, m) => t + m.lignesCode, 0),
    empreinteProjet: modulesCode.length ? empreinte(modulesCode.map((m) => m.empreinte).sort().join()) : null,
    constats: constats.sort((a, b) => REGLES[a.regle].p - REGLES[b.regle].p || a.regle.localeCompare(b.regle) || a.ligne - b.ligne),
    erreurs: f.erreurs,
  };
}

/**
 * Inventaire complet. options : { limite, maintenant, progression(n, total), exclure, maxTaille, edition }
 */
export async function inventorier(racines, options = {}) {
  const maintenant = options.maintenant ?? new Date();
  const liste = await lister(racines, options);
  // Projets VB6 : leurs fichiers sources sont lus avec le .vbp, et non comme modules isolés.
  const membres = new Set();
  const annexesVbp = new Map();
  for (const x of liste.fichiers.filter((y) => familleDe(y.plein) === 'vb6')) {
    let vbp;
    try { vbp = lireVbp(decoderTexte(await readFile(x.plein))); } catch { continue; }
    const contenus = new Map();
    for (const m of vbp.membres) {
      const plein = await trouver(path.dirname(x.plein), m.fichier);
      if (!plein) continue;
      membres.add(plein.toLowerCase());
      try { contenus.set(m.fichier.toLowerCase(), await readFile(plein)); } catch (e) { liste.inaccessibles.push({ chemin: plein, raison: e.code ?? e.message }); }
    }
    annexesVbp.set(x.plein, contenus);
  }
  const candidats = liste.fichiers.filter((x) => !membres.has(x.plein.toLowerCase()));
  // La limite de l'édition gratuite porte sur les fichiers Office, VB6 et scripts, pas sur les sources .NET ni les configurations.
  const limite = options.limite ?? Infinity;
  const principaux = candidats.filter((x) => !ANNEXES.includes(familleDe(x.plein)));
  const aAnalyser = [...principaux.slice(0, limite), ...candidats.filter((x) => ANNEXES.includes(familleDe(x.plein)))];
  const nonAnalyses = principaux.slice(limite).map((x) => relatif(x));
  const fichiers = [];
  const annexesLues = {};
  let n = 0;
  for (const x of aAnalyser) {
    let buf;
    try { buf = await readFile(x.plein); } catch (e) { liste.inaccessibles.push({ chemin: x.plein, raison: e.code ?? e.message }); continue; }
    const r = analyserFichier(relatif(x), buf, { maintenant, annexes: annexesVbp.get(x.plein) });
    options.progression?.(++n, aAnalyser.length);
    if (ANNEXES.includes(r.famille)) {
      annexesLues[r.famille] = (annexesLues[r.famille] ?? 0) + 1;
      if (!r.constats.length) continue;
    }
    r.taille = x.taille;
    r.modifieDisque = x.modifie;
    fichiers.push(r);
  }
  return {
    schema: 'inventaire-vba/1',
    outil: { nom: 'inventaire-vba', version: VERSION, edition: options.edition ?? 'gratuite' },
    date: maintenant.toISOString(),
    racines: racines.map((r) => path.resolve(r)),
    fichiers,
    nonAnalyses,
    sansMacro: liste.sansMacro,
    annexesLues,
    inaccessibles: liste.inaccessibles.map((x) => ({ chemin: x.chemin, raison: x.raison })),
    tropGros: liste.tropGros,
  };
}

// Fichier d'un projet VB6 (chemin relatif au .vbp, séparateurs Windows) ; insensible à la casse comme Windows.
async function trouver(dossier, relatif) {
  let courant = dossier;
  for (const seg of relatif.split(/[\\/]/).filter(Boolean)) {
    if (seg === '..' || seg === '.') { courant = path.resolve(courant, seg); continue; }
    let noms;
    try { noms = await readdir(courant); } catch { return null; }
    const trouve = noms.find((x) => x === seg) ?? noms.find((x) => x.toLowerCase() === seg.toLowerCase());
    if (!trouve) return null;
    courant = path.join(courant, trouve);
  }
  return courant;
}

// Chemin relatif à la racine analysée, préfixé du nom de la racine (« Finance/Budget/outil.xlsm »).
function relatif(x) {
  const base = path.basename(x.racine) || x.racine;
  return [base, ...path.relative(x.racine, x.plein).split(path.sep)].filter(Boolean).join('/');
}

/** Chiffres de synthèse, recalculés à partir des fichiers (aussi utilisés avec --depuis). */
export function synthese(inv) {
  const f = inv.fichiers;
  const vba = (x) => !['vb6', 'dotnet', 'config'].includes(x.famille);
  const avecMacros = f.filter((x) => x.macros && vba(x));
  const vb6 = f.filter((x) => x.famille === 'vb6');
  const parRegle = {};
  for (const x of f) for (const c of x.constats) {
    const r = (parRegle[c.regle] ??= { occurrences: 0, fichiers: new Set() });
    r.occurrences++;
    r.fichiers.add(x.chemin);
  }
  const prio = (x) => Math.min(5, ...x.constats.map((c) => REGLES[c.regle].p));
  const parPriorite = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  for (const x of f) parPriorite[prio(x)]++;
  const sig = { V3: 0, agile: 0, ancienne: 0, aucune: 0 };
  for (const x of avecMacros) sig[x.signature.etat]++;
  return {
    fichiersAnalyses: f.length,
    fichiersAvecMacros: avecMacros.length,
    fichiersSansMacroComptes: Object.values(inv.sansMacro ?? {}).reduce((a, b) => a + b, 0),
    nonAnalyses: inv.nonAnalyses?.length ?? 0,
    lignesCode: f.filter(vba).reduce((t, x) => t + x.lignesCode, 0),
    modules: f.filter(vba).reduce((t, x) => t + x.projets.reduce((u, p) => u + p.modules.filter((m) => m.lignesCode > 0).length, 0), 0),
    vb6: { projets: vb6.length, lignes: vb6.reduce((t, x) => t + x.lignesCode, 0), formulaires: vb6.reduce((t, x) => t + x.projets.reduce((u, p) => u + p.modules.filter((m) => m.type === 'formulaire').length, 0), 0) },
    courriel: { fichiers: f.filter((x) => x.courriel).length, smtp: f.filter((x) => x.constats.some((c) => c.regle === 'MEL-001')).length, outlook: f.filter((x) => x.constats.some((c) => /^OUT-/.test(c.regle))).length },
    parPriorite,
    fichiersP1: f.filter((x) => prio(x) === 1).length,
    fichiersP2: f.filter((x) => prio(x) === 2).length,
    signatures: sig,
    parFamille: Object.fromEntries(['excel', 'word', 'powerpoint', 'access', 'module', 'script', 'vb6', 'dotnet', 'config'].map((k) => [k, f.filter((x) => (x.macros || x.constats.length) && x.famille === k).length])),
    parRegle: Object.fromEntries(Object.entries(parRegle).sort(([a], [b]) => REGLES[a].p - REGLES[b].p || a.localeCompare(b)).map(([k, v]) => [k, { occurrences: v.occurrences, fichiers: v.fichiers.size }])),
    prioriteFichier: Object.fromEntries(f.map((x) => [x.chemin, prio(x)])),
  };
}

/**
 * Version partageable : sans extraits de code, sans auteurs ; avec `sansChemins`, noms de fichiers et dossiers remplacés.
 * C'est le fichier envoyé pour un diagnostic à distance.
 */
export function anonymiser(inv, { sansChemins = false } = {}) {
  const segments = new Map();
  const alias = (seg, prefixe) => {
    if (!segments.has(seg)) segments.set(seg, `${prefixe}-${String(segments.size + 1).padStart(3, '0')}`);
    return segments.get(seg);
  };
  const chemin = (c) => {
    if (!sansChemins) return c;
    const parts = c.split('/');
    const nom = parts.pop();
    return [...parts.map((p) => alias(p, 'dossier')), `${alias(nom, 'fichier')}.${path.extname(nom).slice(1)}`].join('/');
  };
  return {
    ...inv,
    anonymise: { code: true, auteurs: true, chemins: sansChemins },
    racines: inv.racines.map(() => '(masqué)'),
    fichiers: inv.fichiers.map((x) => ({
      ...x,
      chemin: chemin(x.chemin),
      meta: { modifie: x.meta.modifie },
      signature: { ...x.signature, certificats: x.signature.certificats.map((c) => ({ ...c, sujet: '(masqué)', emetteur: c.autosigne ? '(autosigné)' : c.emetteur })) },
      projets: x.projets.map((p) => ({ ...p, ...(p.vb6 ? { nom: sansChemins ? '(masqué)' : p.nom, vb6: { ...p.vb6, nom: sansChemins ? '(masqué)' : p.vb6.nom, exe: sansChemins ? '' : p.vb6.exe, manquants: p.vb6.manquants.map(() => '(masqué)') } } : {}), modules: p.modules.map((m) => ({ ...m, nom: sansChemins ? '(masqué)' : m.nom })), references: p.references.map((r) => ({ ...r, chemin: r.chemin && !/^[a-z]:\\(windows|program files)/i.test(r.chemin) ? '(masqué)' : r.chemin })) })),
      constats: x.constats.map((c) => ({ ...c, module: sansChemins && !c.module.startsWith('(') ? '(masqué)' : c.module, extrait: '' })),
      erreurs: x.erreurs.map(() => 'erreur de lecture'),
    })),
    nonAnalyses: inv.nonAnalyses.map(chemin),
    inaccessibles: inv.inaccessibles.map((x) => ({ chemin: sansChemins ? '(masqué)' : x.chemin, raison: x.raison })),
    tropGros: inv.tropGros.map((x) => ({ chemin: chemin(x.chemin), taille: x.taille })),
  };
}
