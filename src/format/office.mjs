// Ouvre un fichier du parc et en extrait ce qui compte pour l'inventaire : projets VBA, signatures,
// métadonnées, contrôles ActiveX, macros Excel 4.0, liaisons DDE. Lecture seule, rien n'est exécuté.
import { X509Certificate } from 'node:crypto';
import path from 'node:path';
import { lireCfb, estCfb, ErreurFormat } from './cfb.mjs';
import { lireZip, estZip } from './zip.mjs';
import { lireProjetVba, decodeur } from './ovba.mjs';

export const EXTENSIONS = {
  excel: ['xls', 'xlsm', 'xlsb', 'xlt', 'xltm', 'xla', 'xlam'],
  word: ['doc', 'docm', 'dot', 'dotm'],
  powerpoint: ['ppt', 'pptm', 'pot', 'potm', 'ppa', 'ppam', 'pps', 'ppsm'],
  access: ['mdb', 'accdb', 'mde', 'accde', 'accda', 'mda'],
  module: ['bas', 'cls', 'frm'],
  script: ['vbs'],
  vb6: ['vbp'],
  // Sources .NET et fichiers de configuration : lus pour l'envoi de courriels (Outlook par COM, SMTP par
  // mot de passe). Gardés dans l'inventaire seulement s'ils déclenchent une règle.
  dotnet: ['cs', 'vb', 'csproj', 'vbproj'],
  config: ['config', 'ini'],
};
// Familles « annexes » : ni macros ni projet ; elles ne comptent pas dans la limite de l'édition gratuite.
export const ANNEXES = ['dotnet', 'config'];
// Formats sans macros possibles : seulement comptés (sans ouverture) pour donner la part du parc.
export const SANS_MACRO = ['xlsx', 'xltx', 'docx', 'dotx', 'pptx', 'potx', 'ppsx'];
const FAMILLE = Object.fromEntries(Object.entries(EXTENSIONS).flatMap(([f, l]) => l.map((e) => [e, f])));
export const familleDe = (fichier) => FAMILLE[path.extname(fichier).slice(1).toLowerCase()] ?? null;

const vide = (famille, format) => ({ famille, format, projets: [], signature: { legacy: false, agile: false, v3: false, certificats: [] }, meta: {}, activeX: [], xlm: false, dde: false, chiffre: false, accessVbaProbable: false, erreurs: [] });

export function analyserOffice(chemin, buf, { annexes } = {}) {
  const famille = familleDe(chemin);
  if (famille === 'vb6') return analyserVbp(chemin, buf, annexes ?? new Map());
  if (famille === 'dotnet' || famille === 'config') { const r = vide(famille, 'texte'); r.texte = decoderTexte(buf); return r; }
  if (famille === 'module' || famille === 'script') return analyserTexte(chemin, buf, famille);
  if (famille === 'access') return analyserAccess(buf);
  if (estZip(buf)) return analyserOoxml(buf, famille);
  if (estCfb(buf)) return analyserOle(buf, famille);
  const r = vide(famille, 'inconnu');
  r.erreurs.push('format non reconnu (ni OLE, ni ZIP) : fichier renommé ou corrompu');
  return r;
}

// ---------- Office Open XML (.xlsm, .xlsb, .docm, .pptm…) ----------
function analyserOoxml(buf, famille) {
  const r = vide(famille, 'ooxml');
  const z = lireZip(buf);
  const core = z.noms.includes('docProps/core.xml') ? z.texte('docProps/core.xml') : null;
  if (core) {
    const v = (balise) => xmlTexte(new RegExp(`<${balise}[^>]*>([^<]*)</${balise}>`).exec(core)?.[1]);
    r.meta = { auteur: v('dc:creator'), dernierAuteur: v('cp:lastModifiedBy'), modifie: v('dcterms:modified') };
  }
  for (const nom of z.noms.filter((n) => /(^|\/)vbaProject\.bin$/i.test(n))) {
    try {
      r.projets.push(lireProjetVba(lireCfb(z.lire(nom)).racine));
    } catch (e) {
      r.erreurs.push(`${nom} : ${e.message}`);
    }
    // Signatures : relations du projet si présentes, sinon noms des parties.
    const dossier = nom.replace(/vbaProject\.bin$/i, '');
    const rels = z.texte(`${dossier}_rels/vbaProject.bin.rels`) ?? '';
    const parties = z.noms.filter((n) => n.startsWith(dossier) && /vbaProjectSignature/i.test(n));
    const types = rels + parties.join(' ');
    r.signature.v3 ||= /vbaProjectSignatureV3/i.test(types);
    r.signature.agile ||= /vbaProjectSignatureAgile/i.test(types);
    r.signature.legacy ||= /vbaProjectSignature(\.bin|"|\s|$)/i.test(types) || parties.some((p) => /vbaProjectSignature\.bin$/i.test(p));
    for (const p of parties) r.signature.certificats.push(...certificats(z.lire(p)));
  }
  r.activeX = z.noms.filter((n) => /\/activeX\/activeX\d+\.xml$/i.test(n)).map((n) => /classid="\{?([0-9A-Fa-f-]{36})\}?"/.exec(z.texte(n) ?? '')?.[1]?.toUpperCase()).filter(Boolean);
  r.xlm = z.noms.some((n) => /^xl\/macrosheets\//i.test(n));
  r.dde = z.noms.filter((n) => /^xl\/externalLinks\/.*\.xml$/i.test(n)).some((n) => /<ddeLink/i.test(z.texte(n) ?? ''));
  r.signature.certificats = uniques(r.signature.certificats);
  return r;
}

// ---------- Formats binaires 97-2003 (.xls, .doc, .ppt) ----------
function analyserOle(buf, famille) {
  const r = vide(famille, 'ole');
  const cfb = lireCfb(buf);
  const racine = cfb.racine;
  if (racine.enfant('EncryptedPackage')) {
    r.format = 'ooxml';
    r.chiffre = true;
    return r;
  }
  const fluxSig = { legacy: '\u0005DigitalSignature', agile: '\u0005DigitalSignatureEx', v3: '\u0005DigitalSignatureExt' };
  for (const [k, nom] of Object.entries(fluxSig)) {
    const f = racine.enfant(nom);
    if (f) { r.signature[k] = true; r.signature.certificats.push(...certificats(f.lire())); }
  }
  r.signature.certificats = uniques(r.signature.certificats);
  const si = racine.enfant('\u0005SummaryInformation');
  if (si) { try { r.meta = resume(si.lire()); } catch { /* métadonnées facultatives */ } }
  // Tout stockage qui contient PROJECT et VBA/dir est un projet (Excel : _VBA_PROJECT_CUR, Word : Macros).
  const visiter = (n, profondeur) => {
    if (profondeur > 4) return;
    for (const e of n.enfants().filter((x) => x.estStockage)) {
      if (e.enfant('VBA')?.enfant('dir')) {
        try { r.projets.push(lireProjetVba(e)); } catch (err) { r.erreurs.push(`${e.nom} : ${err.message}`); }
      } else visiter(e, profondeur + 1);
    }
  };
  visiter(racine, 0);
  // PowerPoint 97-2003 range son projet VBA compressé dans le flux du document : non analysé (rare).
  if (famille === 'powerpoint' && !r.projets.length) r.pptNonAnalyse = true;
  return r;
}

// Jeu de propriétés \005SummaryInformation ([MS-OLEPS]) : auteur (4), dernier auteur (8), dernière sauvegarde (13).
function resume(b) {
  const offSection = b.readUInt32LE(44);
  const nb = b.readUInt32LE(offSection + 4);
  const meta = {};
  let page = 1252;
  const props = [];
  for (let k = 0; k < nb; k++) props.push([b.readUInt32LE(offSection + 8 + k * 8), offSection + b.readUInt32LE(offSection + 12 + k * 8)]);
  for (const [pid, o] of props) if (pid === 1 && b.readUInt32LE(o) === 2) page = b.readUInt16LE(o + 4);
  for (const [pid, o] of props) {
    const type = b.readUInt32LE(o);
    if (type === 0x1e && (pid === 4 || pid === 8)) {
      const lg = b.readUInt32LE(o + 4);
      const v = decodeur(page).decode(b.subarray(o + 8, o + 8 + lg)).replace(/\0+$/, '');
      meta[pid === 4 ? 'auteur' : 'dernierAuteur'] = v;
    }
    if (type === 0x40 && pid === 13) {
      const ft = b.readBigUInt64LE(o + 4);
      if (ft > 0n) meta.modifie = new Date(Number(ft / 10000n - 11644473600000n)).toISOString();
    }
  }
  return meta;
}

// ---------- Access : le code n'est pas lisible sans Access ; on signale sa présence probable ----------
function analyserAccess(buf) {
  const r = vide('access', 'access');
  // Le projet VBA d'une base est un fichier OLE rangé dans la table MSysAccessStorage :
  // ses noms de flux (« _VBA_PROJECT », « dir ») y apparaissent en UTF-16.
  r.accessVbaProbable = buf.includes(Buffer.from('_VBA_PROJECT', 'utf16le')) || buf.includes(Buffer.from('PROJECTwm', 'utf16le'));
  return r;
}

// ---------- Modules exportés (.bas, .cls, .frm) et scripts .vbs ----------
export function decoderTexte(buf) {
  if (buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) return buf.subarray(3).toString('utf8');
  if (buf[0] === 0xff && buf[1] === 0xfe) return buf.subarray(2).toString('utf16le');
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch { return new TextDecoder('windows-1252').decode(buf); }
}

function analyserTexte(chemin, buf, famille) {
  let texte = decoderTexte(buf);
  // Un .frm exporté de Visual Basic 6 (« Begin VB.Form ») n'est pas un UserForm VBA.
  if (/^Begin\s+VB\.(Form|MDIForm|UserControl)\b/im.test(texte)) {
    const nom = /^Attribute VB_Name = "([^"]+)"/im.exec(texte)?.[1] ?? path.basename(chemin, path.extname(chemin));
    const r = vide('vb6', 'texte');
    const m = moduleVb6(nom, path.extname(chemin).slice(1).toLowerCase(), texte);
    r.projets.push({ nom, langage: 'vb6', vb6: { type: 'source isolée', nom, exe: '', membres: 1, manquants: [] }, codePage: 1252, references: [], objets: m.objets, modules: [m.module], protege: false, erreurs: [] });
    return r;
  }
  const r = vide(famille, 'texte');
  const nom = /^Attribute VB_Name = "([^"]+)"/im.exec(texte)?.[1] ?? path.basename(chemin, path.extname(chemin));
  const ext = path.extname(chemin).slice(1).toLowerCase();
  const type = famille === 'script' ? 'script' : ext === 'frm' ? 'formulaire' : ext === 'cls' ? (/^Attribute VB_Base = "/im.test(texte) ? 'document' : 'classe') : 'module';
  // .frm : l'en-tête de description du formulaire précède le code (« VERSION 5.00 … End »).
  if (ext === 'frm') texte = texte.replace(/^VERSION[\s\S]*?^End\s*$/im, '');
  r.projets.push({ nom, codePage: 1252, references: [], modules: [{ nom, type, code: texte }], protege: false, erreurs: [] });
  return r;
}

// ---------- Projets Visual Basic 6 (.vbp et ses fichiers sources) ----------
// Le .vbp est un fichier texte « Clé=Valeur » : membres (Form, Module, Class, UserControl…), références
// (Reference=*\G{guid}#ver#0#chemin#description) et contrôles OCX (Object={guid}#ver#0; fichier.ocx).
export function lireVbp(texte) {
  const p = { type: 'Exe', nom: '', exe: '', membres: [], references: [], objets: [] };
  for (const brute of texte.split(/\r\n|\r|\n/)) {
    const m = /^([A-Za-z0-9]+)=(.*)$/.exec(brute.trim());
    if (!m) continue;
    const [, cle, val] = m;
    const k = cle.toLowerCase();
    if (k === 'type') p.type = val.trim();
    else if (k === 'name') p.nom = val.replace(/^"|"$/g, '');
    else if (k === 'exename32') p.exe = val.replace(/^"|"$/g, '');
    else if (['form', 'module', 'class', 'usercontrol', 'userdocument', 'propertypage', 'designer'].includes(k)) {
      // « Module=modMain; modMain.bas » ou « Form=frmMain.frm »
      const fichier = val.includes(';') ? val.split(';').pop().trim() : val.trim();
      p.membres.push({ type: k, fichier: fichier.replace(/\\/g, '/') });
    } else if (k === 'reference') {
      const r = /^\*\\G\{([0-9A-F-]{36})\}#([^#]*)#[^#]*#([^#]*)#(.*)$/i.exec(val);
      if (r) p.references.push({ nom: r[4] || path.win32.basename(r[3]), guid: r[1].toUpperCase(), version: r[2], chemin: r[3], description: r[4], type: 'registered' });
    } else if (k === 'object') {
      const o = /^\{([0-9A-F-]{36})\}#([^;]*);\s*(.+)$/i.exec(val);
      if (o) p.objets.push({ guid: o[1].toUpperCase(), version: o[2].split('#')[0], fichier: o[3].trim() });
    }
  }
  return p;
}

const TYPES_VB6 = { form: 'formulaire', module: 'module', class: 'classe', usercontrol: 'contrôle', userdocument: 'document', propertypage: 'formulaire', designer: 'concepteur', frm: 'formulaire', bas: 'module', cls: 'classe', ctl: 'contrôle' };

// Un fichier source VB6 : l'en-tête (VERSION, Object = …, Begin VB.Form … End) est retiré, ses OCX gardés.
function moduleVb6(nomParDefaut, type, texte) {
  const objets = [...texte.matchAll(/^Object\s*=\s*"\{([0-9A-F-]{36})\}#([^"]*)";\s*"([^"]+)"/gim)].map((m) => ({ guid: m[1].toUpperCase(), version: m[2].split('#')[0], fichier: m[3] }));
  const nom = /^Attribute VB_Name = "([^"]+)"/im.exec(texte)?.[1] ?? nomParDefaut;
  let code = texte;
  if (/^VERSION\s/im.test(code)) {
    // Retire tout jusqu'à la fermeture du dernier bloc « Begin … End » de premier niveau.
    const lignes = code.split(/\r\n|\r|\n/);
    let profondeur = 0, fin = -1, vu = false;
    for (let i = 0; i < lignes.length; i++) {
      const l = lignes[i].trim();
      if (/^Begin(Property)?(\s|$)/i.test(l)) { profondeur++; vu = true; }
      else if (/^End(Property)?$/i.test(l) && profondeur) { profondeur--; if (!profondeur && vu) fin = i; }
      else if (/^Attribute\s+VB_/i.test(l) && !profondeur) break;
    }
    code = lignes.slice(fin + 1).join('\r\n').replace(/^(?:VERSION\s.*|Object\s*=.*)\r?\n/gim, '');
  }
  return { objets, module: { nom, type: TYPES_VB6[type] ?? 'module', code } };
}

function analyserVbp(chemin, buf, annexes) {
  const r = vide('vb6', 'vb6');
  const vbp = lireVbp(decoderTexte(buf));
  const nom = vbp.nom || path.basename(chemin, path.extname(chemin));
  const modules = [];
  const objets = [...vbp.objets];
  const manquants = [];
  for (const m of vbp.membres) {
    const contenu = annexes.get(m.fichier.toLowerCase());
    if (!contenu) { manquants.push(m.fichier); continue; }
    const x = moduleVb6(path.basename(m.fichier, path.extname(m.fichier)), m.type, decoderTexte(contenu));
    modules.push(x.module);
    for (const o of x.objets) if (!objets.some((y) => y.guid === o.guid)) objets.push(o);
  }
  if (manquants.length) r.erreurs.push(`${manquants.length} fichier(s) du projet introuvable(s) : ${manquants.slice(0, 5).join(', ')}`);
  r.projets.push({ nom, langage: 'vb6', vb6: { type: vbp.type, nom, exe: vbp.exe, membres: vbp.membres.length, manquants }, codePage: 1252, references: vbp.references, objets, modules, protege: false, erreurs: [] });
  return r;
}

// ---------- Certificats de signature ----------
// La signature est une structure PKCS#7 : on y cherche les certificats X.509 (SEQUENCE DER de longueur 2 octets).
export function certificats(blob) {
  const out = [];
  for (let i = 0; i + 8 < blob.length; i++) {
    if (blob[i] !== 0x30 || blob[i + 1] !== 0x82 || blob[i + 4] !== 0x30 || blob[i + 5] !== 0x82) continue;
    const lg = blob.readUInt16BE(i + 2) + 4;
    if (i + lg > blob.length || lg < 300) continue;
    try {
      const c = new X509Certificate(blob.subarray(i, i + lg));
      out.push({ sujet: cn(c.subject), emetteur: cn(c.issuer), debut: new Date(c.validFrom).toISOString(), fin: new Date(c.validTo).toISOString(), autorite: c.ca, autosigne: c.subject === c.issuer, empreinte: c.fingerprint256.replace(/:/g, '').slice(0, 16) });
      i += lg - 1;
    } catch { /* pas un certificat */ }
  }
  return out;
}
const cn = (dn) => /(?:^|\n)CN=([^\n]+)/.exec(dn)?.[1] ?? dn.split('\n').pop();
const uniques = (l) => [...new Map(l.map((c) => [c.empreinte, c])).values()];
const xmlTexte = (s) => s?.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

export { ErreurFormat };
