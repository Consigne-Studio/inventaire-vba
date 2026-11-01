// Fabrique de fichiers Office fictifs pour les tests : fichier composé OLE, projet VBA compressé,
// classeurs .xlsm / .xls / .docm. Aucun fichier réel de tiers : tout est construit ici, selon [MS-CFB] et [MS-OVBA].
import { ecrireZip } from '../../src/format/zip.mjs';

// ---------- Compression VBA (MS-OVBA 2.4.1.3) ----------
export function compresser(src) {
  const out = [0x01];
  for (let debut = 0; debut < src.length || (debut === 0 && src.length === 0); debut += 4096) {
    const fin = Math.min(debut + 4096, src.length);
    const morceau = [];
    let p = debut;
    while (p < fin) {
      const iDrapeau = morceau.length;
      morceau.push(0);
      for (let bit = 0; bit < 8 && p < fin; bit++) {
        const dif = p - debut;
        let meilleur = 0, decal = 0;
        if (dif > 0) {
          const nbBits = Math.max(Math.ceil(Math.log2(dif)), 4);
          const maxLg = (0xffff >> nbBits) + 3;
          for (let c = p - 1; c >= debut; c--) {
            let lg = 0;
            while (lg < maxLg && p + lg < fin && src[c + lg] === src[p + lg]) lg++;
            if (lg > meilleur) { meilleur = lg; decal = p - c; }
            if (meilleur === maxLg) break;
          }
          if (meilleur >= 3) {
            const jeton = ((decal - 1) << (16 - nbBits)) | (meilleur - 3);
            morceau.push(jeton & 0xff, jeton >> 8);
            morceau[iDrapeau] |= 1 << bit;
            p += meilleur;
            continue;
          }
        }
        morceau.push(src[p++]);
      }
    }
    if (morceau.length > 4096) {
      const entete = 0x3000 | 0x0fff;
      out.push(entete & 0xff, entete >> 8);
      for (let k = 0; k < 4096; k++) out.push(src[debut + k] ?? 0);
    } else {
      const entete = 0xb000 | (morceau.length - 1);
      out.push(entete & 0xff, entete >> 8, ...morceau);
    }
    if (src.length === 0) break;
  }
  return Buffer.from(out);
}

// ---------- Fichier composé OLE, version 3 (secteurs de 512 octets) ----------
// arbre : { enfants: [ { nom, donnees } | { nom, enfants } ] }
export function ecrireCfb(arbre) {
  const FIN = 0xfffffffe, LIBRE = 0xffffffff, FATSECT = 0xfffffffd;
  const entrees = [{ nom: 'Root Entry', type: 5, noeud: arbre }];
  const indexer = (n, e) => {
    const enf = [...(n.enfants ?? [])].sort((a, b) => a.nom.length - b.nom.length || a.nom.toUpperCase().localeCompare(b.nom.toUpperCase()));
    e.enfantsIdx = enf.map((c) => {
      const x = { nom: c.nom, type: c.enfants ? 1 : 2, noeud: c, donnees: c.donnees ? Buffer.from(c.donnees) : null };
      entrees.push(x);
      return x;
    });
    for (const x of e.enfantsIdx) if (x.type === 1) indexer(x.noeud, x);
  };
  indexer(arbre, entrees[0]);
  entrees.forEach((e, i) => { e.i = i; });
  const flux = entrees.filter((e) => e.type === 2);
  const petits = flux.filter((e) => e.donnees.length < 4096);
  const gros = flux.filter((e) => e.donnees.length >= 4096);
  // mini-flux
  const miniParts = [];
  const miniFat = [];
  for (const e of petits) {
    if (!e.donnees.length) { e.debut = FIN; continue; }
    const n = Math.ceil(e.donnees.length / 64);
    e.debut = miniFat.length;
    for (let k = 0; k < n; k++) miniFat.push(k === n - 1 ? FIN : miniFat.length + 1);
    const pad = Buffer.alloc(n * 64);
    e.donnees.copy(pad);
    miniParts.push(pad);
  }
  const miniFlux = Buffer.concat(miniParts);
  const secteurs = (lg) => Math.ceil(lg / 512);
  const nbDir = secteurs(entrees.length * 128);
  const nbMiniFat = secteurs(miniFat.length * 4);
  const nbMini = secteurs(miniFlux.length);
  const nbGros = gros.reduce((s, e) => s + secteurs(e.donnees.length), 0);
  let nbFat = 1;
  while (nbFat * 128 < nbFat + nbDir + nbMiniFat + nbMini + nbGros) nbFat++;
  if (nbFat > 109) throw new Error('fabrique : fichier trop grand');
  const fat = [];
  const chaine = (n) => {
    if (!n) return FIN;
    const d = fat.length;
    for (let k = 0; k < n; k++) fat.push(k === n - 1 ? FIN : fat.length + 1);
    return d;
  };
  for (let k = 0; k < nbFat; k++) fat.push(FATSECT);
  const debutDir = chaine(nbDir);
  const debutMiniFat = chaine(nbMiniFat);
  const debutMini = chaine(nbMini);
  for (const e of gros) e.debut = chaine(secteurs(e.donnees.length));
  while (fat.length % 128) fat.push(LIBRE);
  while (fat.length < nbFat * 128) fat.push(LIBRE);

  const entete = Buffer.alloc(512);
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(entete, 0);
  entete.writeUInt16LE(0x3e, 0x18); entete.writeUInt16LE(3, 0x1a); entete.writeUInt16LE(0xfffe, 0x1c);
  entete.writeUInt16LE(9, 0x1e); entete.writeUInt16LE(6, 0x20);
  entete.writeUInt32LE(nbFat, 0x2c); entete.writeUInt32LE(debutDir, 0x30); entete.writeUInt32LE(4096, 0x38);
  entete.writeUInt32LE(nbMiniFat ? debutMiniFat : FIN, 0x3c); entete.writeUInt32LE(nbMiniFat, 0x40);
  entete.writeUInt32LE(FIN, 0x44); entete.writeUInt32LE(0, 0x48);
  for (let k = 0; k < 109; k++) entete.writeUInt32LE(k < nbFat ? k : LIBRE, 0x4c + k * 4);

  const dir = Buffer.alloc(nbDir * 512);
  for (let o = 0; o < dir.length; o += 128) { dir.writeUInt32LE(LIBRE, o + 68); dir.writeUInt32LE(LIBRE, o + 72); dir.writeUInt32LE(LIBRE, o + 76); }
  for (const e of entrees) {
    const o = e.i * 128;
    const nom = Buffer.from(e.nom + '\0', 'utf16le');
    nom.copy(dir, o, 0, Math.min(64, nom.length));
    dir.writeUInt16LE(Math.min(64, nom.length), o + 64);
    dir[o + 66] = e.type; dir[o + 67] = 1;
    if (e.enfantsIdx?.length) {
      dir.writeUInt32LE(e.enfantsIdx[0].i, o + 76);
      e.enfantsIdx.forEach((c, k) => { if (k + 1 < e.enfantsIdx.length) dir.writeUInt32LE(e.enfantsIdx[k + 1].i, c.i * 128 + 72); });
    }
    if (e.type === 5) { dir.writeUInt32LE(nbMini ? debutMini : FIN, o + 116); dir.writeUInt32LE(miniFlux.length, o + 120); }
    if (e.type === 2) { dir.writeUInt32LE(e.donnees.length ? e.debut : FIN, o + 116); dir.writeUInt32LE(e.donnees.length, o + 120); }
  }
  const fatBuf = Buffer.alloc(nbFat * 512);
  fat.forEach((v, k) => fatBuf.writeUInt32LE(v >>> 0, k * 4));
  const pad = (b) => { const x = Buffer.alloc(secteurs(b.length) * 512); b.copy(x); return x; };
  const mf = Buffer.alloc(nbMiniFat * 512, 0xff);
  miniFat.forEach((v, k) => mf.writeUInt32LE(v >>> 0, k * 4));
  return Buffer.concat([entete, fatBuf, dir, mf, pad(miniFlux), ...gros.map((e) => pad(e.donnees))]);
}

// ---------- Chiffrement « Data Encryption » (MS-OVBA 2.4.3.2) ----------
export function chiffrer(donnees, idProjet, graine = 0x8e) {
  const cle = [...idProjet].reduce((s, c) => (s + c.charCodeAt(0)) & 0xff, 0);
  const out = [graine, graine ^ 2, graine ^ cle];
  let nonChiffre1 = cle, chiffre1 = graine ^ cle, chiffre2 = graine ^ 2;
  const pousser = (x) => {
    const c = x ^ ((chiffre2 + nonChiffre1) & 0xff);
    out.push(c);
    chiffre2 = chiffre1; chiffre1 = c; nonChiffre1 = x;
  };
  for (let k = 0; k < (graine & 6) / 2; k++) pousser(0x07);
  const lg = donnees.length;
  [lg & 0xff, (lg >> 8) & 0xff, (lg >> 16) & 0xff, (lg >>> 24) & 0xff].forEach(pousser);
  for (const x of donnees) pousser(x);
  return Buffer.from(out).toString('hex').toUpperCase();
}

// ---------- Projet VBA ----------
const ENTETES = {
  classeur: (n) => `Attribute VB_Name = "${n}"\r\nAttribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n`,
  feuille: (n) => `Attribute VB_Name = "${n}"\r\nAttribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n`,
  document: (n) => `Attribute VB_Name = "${n}"\r\nAttribute VB_Base = "1Normal.ThisDocument"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\nAttribute VB_TemplateDerived = True\r\nAttribute VB_Customizable = True\r\n`,
  module: (n) => `Attribute VB_Name = "${n}"\r\n`,
  classe: (n) => `Attribute VB_Name = "${n}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\n`,
  formulaire: (n) => `Attribute VB_Name = "${n}"\r\nAttribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}{8F1A2C3B-1111-4A2B-9C3D-0123456789AB}"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = False\r\nAttribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = False\r\n`,
};

export const REF_STDOLE = { nom: 'stdole', libid: '*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation' };
export const REF_OFFICE = { nom: 'Office', libid: '*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library' };

const rec = (id, donnees = Buffer.alloc(0)) => {
  const h = Buffer.alloc(6);
  h.writeUInt16LE(id, 0); h.writeUInt32LE(donnees.length, 2);
  return Buffer.concat([h, donnees]);
};
const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16LE(v); return b; };
const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); return b; };
const ansi = (s) => Buffer.from(s, 'latin1');
const avecTaille = (s) => Buffer.concat([u32(ansi(s).length), ansi(s)]);

/**
 * projet : { nom, modules: [{ nom, type: 'module'|'classe'|'classeur'|'feuille'|'document'|'formulaire', code }],
 *            references: [{ nom, libid, type?: 'registered'|'project'|'control', relatif? }], motDePasse?, win64? }
 * Renvoie l'arbre CFB du stockage (enfants : PROJECT, PROJECTwm, VBA).
 */
export function arbreProjet(projet) {
  const id = '{9E394C0B-697E-4AEE-9FA6-446F51FB30F3}';
  const refs = [REF_STDOLE, REF_OFFICE, ...(projet.references ?? [])];
  const parts = [
    rec(0x0001, u32(projet.win64 ? 3 : 1)), rec(0x0002, u32(0x40c)), rec(0x0014, u32(0x40c)), rec(0x0003, u16(1252)),
    rec(0x0004, ansi(projet.nom ?? 'VBAProject')), rec(0x0005), rec(0x0040), rec(0x0006), rec(0x003d), rec(0x0007, u32(0)), rec(0x0008, u32(0)),
    Buffer.concat([u16(0x0009), u32(4), u32(1), u16(0)]), rec(0x000c), rec(0x003c),
  ];
  for (const r of refs) {
    parts.push(rec(0x0016, ansi(r.nom)), rec(0x003e, Buffer.from(r.nom, 'utf16le')));
    if (r.type === 'project') {
      parts.push(rec(0x000e, Buffer.concat([avecTaille(r.libid), avecTaille(r.relatif ?? r.libid), u32(1), u16(0)])));
    } else if (r.type === 'control') {
      parts.push(rec(0x0033, ansi(r.libid)));
      parts.push(rec(0x002f, Buffer.concat([avecTaille(r.libid), u32(0), u16(0)])));
      parts.push(rec(0x0030, Buffer.concat([avecTaille(r.libid), u32(0), u16(0), Buffer.alloc(16), u32(0)])));
    } else {
      parts.push(rec(0x000d, Buffer.concat([avecTaille(r.libid), u32(0), u16(0)])));
    }
  }
  parts.push(rec(0x000f, u16(projet.modules.length)), rec(0x0013, u16(0xffff)));
  const fluxModules = [];
  for (const m of projet.modules) {
    const type = m.type ?? 'module';
    const code = ENTETES[type](m.nom) + m.code.replace(/\r?\n/g, '\r\n');
    parts.push(
      rec(0x0019, ansi(m.nom)), rec(0x0047, Buffer.from(m.nom, 'utf16le')), rec(0x001a, ansi(m.nom)), rec(0x0032, Buffer.from(m.nom, 'utf16le')),
      rec(0x001c), rec(0x0048), rec(0x0031, u32(0)), rec(0x001e, u32(0)), rec(0x002c, u16(0xffff)),
      rec(type === 'module' ? 0x0021 : 0x0022), rec(0x002b),
    );
    fluxModules.push({ nom: m.nom, donnees: compresser(Buffer.from(code, 'latin1')) });
  }
  parts.push(Buffer.concat([u16(0x0010), u32(0)]));
  const dir = compresser(Buffer.concat(parts));

  const lignes = [`ID="${id}"`];
  for (const m of projet.modules) {
    const t = m.type ?? 'module';
    if (t === 'module') lignes.push(`Module=${m.nom}`);
    else if (t === 'classe') lignes.push(`Class=${m.nom}`);
    else if (t === 'formulaire') lignes.push(`BaseClass=${m.nom}`);
    else lignes.push(`Document=${m.nom}/&H00000000`);
  }
  lignes.push(`Name="${projet.nom ?? 'VBAProject'}"`, 'HelpContextID="0"', 'VersionCompatible32="393222000"');
  lignes.push(`CMG="${chiffrer(u32(projet.motDePasse ? 1 : 0), id, 0x1c)}"`);
  lignes.push(`DPB="${chiffrer(projet.motDePasse ? Buffer.from('FFFF1234AABBCCDDEEFF00112233445566778899AABBCC00', 'hex') : Buffer.from([0]), id, 0x3a)}"`);
  lignes.push(`GC="${chiffrer(Buffer.from([projet.motDePasse ? 0 : 0xff]), id, 0x5e)}"`);
  lignes.push('', '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '');

  return [
    { nom: 'PROJECT', donnees: Buffer.from(lignes.join('\r\n'), 'latin1') },
    { nom: 'PROJECTwm', donnees: Buffer.concat(projet.modules.map((m) => Buffer.concat([ansi(m.nom), Buffer.from([0]), Buffer.from(m.nom, 'utf16le'), Buffer.from([0, 0])])).concat([Buffer.from([0])])) },
    { nom: 'VBA', enfants: [{ nom: '_VBA_PROJECT', donnees: Buffer.from('cc61ffff000000', 'hex') }, { nom: 'dir', donnees: dir }, ...fluxModules] },
  ];
}

export const vbaProjectBin = (projet) => ecrireCfb({ enfants: arbreProjet(projet) });

// ---------- Fichiers Office ----------
const CORE = ({ auteur = 'Service Achats', dernier = auteur, modifie = '2024-03-12T09:30:00Z' } = {}) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${auteur}</dc:creator><cp:lastModifiedBy>${dernier}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2019-01-07T08:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${modifie}</dcterms:modified></cp:coreProperties>`;

/**
 * Classeur .xlsm (ou .xlam / .xltm selon `type`). options : projet, signatures: ['legacy','agile','v3'],
 * activeX: [classid], macroFeuille: bool, dde: bool, meta: { auteur, dernier, modifie }, signatureBlob
 */
export function xlsm({ projet, signatures = [], activeX = [], macroFeuille = false, dde = false, meta, signatureBlob, type = 'sheet' } = {}) {
  const ct = type === 'addin' ? 'application/vnd.ms-excel.addin.macroEnabled.main+xml' : type === 'template' ? 'application/vnd.ms-excel.template.macroEnabled.main+xml' : 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
  const fichiers = [];
  const overrides = [`<Override PartName="/xl/workbook.xml" ContentType="${ct}"/>`, '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'];
  const relsWb = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'];
  const feuilles = ['<sheet name="Feuil1" sheetId="1" r:id="rId1"/>'];
  if (projet) {
    fichiers.push({ nom: 'xl/vbaProject.bin', donnees: vbaProjectBin(projet) });
    relsWb.push('<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>');
    const typesSig = { legacy: ['vbaProjectSignature.bin', 'http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature', 'application/vnd.ms-office.vbaProjectSignature'], agile: ['vbaProjectSignatureAgile.bin', 'http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile', 'application/vnd.ms-office.vbaProjectSignatureAgile'], v3: ['vbaProjectSignatureV3.bin', 'http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3', 'application/vnd.ms-office.vbaProjectSignatureV3'] };
    const rels = [];
    signatures.forEach((s, k) => {
      const [nom, typeRel, ctSig] = typesSig[s];
      fichiers.push({ nom: `xl/${nom}`, donnees: signatureBlob ?? Buffer.from('signature fictive') });
      overrides.push(`<Override PartName="/xl/${nom}" ContentType="${ctSig}"/>`);
      rels.push(`<Relationship Id="rId${k + 1}" Type="${typeRel}" Target="${nom}"/>`);
    });
    if (rels.length) fichiers.push({ nom: 'xl/_rels/vbaProject.bin.rels', donnees: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>` });
  }
  activeX.forEach((cls, k) => {
    fichiers.push({ nom: `xl/activeX/activeX${k + 1}.xml`, donnees: `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ax:ocx ax:classid="{${cls}}" ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` });
    overrides.push(`<Override PartName="/xl/activeX/activeX${k + 1}.xml" ContentType="application/vnd.ms-office.activeX+xml"/>`);
  });
  if (macroFeuille) {
    fichiers.push({ nom: 'xl/macrosheets/sheet2.xml', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xm:macrosheet xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData/></xm:macrosheet>' });
    overrides.push('<Override PartName="/xl/macrosheets/sheet2.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/>');
    relsWb.push('<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet2.xml"/>');
    feuilles.push('<sheet name="Macro1" sheetId="2" r:id="rId2"/>');
  }
  if (dde) {
    fichiers.push({ nom: 'xl/externalLinks/externalLink1.xml', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="WINWORD" ddeTopic="C:\\Rapports\\mensuel.docx"/></externalLink>' });
    overrides.push('<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>');
  }
  return ecrireZip([
    { nom: '[Content_Types].xml', donnees: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>${overrides.join('')}</Types>` },
    { nom: '_rels/.rels', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>' },
    { nom: 'docProps/core.xml', donnees: CORE(meta) },
    { nom: 'xl/workbook.xml', donnees: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${feuilles.join('')}</sheets></workbook>` },
    { nom: 'xl/_rels/workbook.xml.rels', donnees: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relsWb.join('')}</Relationships>` },
    { nom: 'xl/worksheets/sheet1.xml', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Données fictives</t></is></c></row></sheetData></worksheet>' },
    ...fichiers,
  ]);
}

/** Document Word .docm : projet dans word/vbaProject.bin. */
export function docm({ projet, signatures = [], meta } = {}) {
  const sig = { legacy: 'vbaProjectSignature.bin', agile: 'vbaProjectSignatureAgile.bin', v3: 'vbaProjectSignatureV3.bin' };
  return ecrireZip([
    { nom: '[Content_Types].xml', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/document.xml" ContentType="application/vnd.ms-word.document.macroEnabled.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>' },
    { nom: '_rels/.rels', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>' },
    { nom: 'docProps/core.xml', donnees: CORE(meta) },
    { nom: 'word/document.xml', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Modèle de courrier fictif</w:t></w:r></w:p></w:body></w:document>' },
    { nom: 'word/_rels/document.xml.rels', donnees: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>' },
    ...(projet ? [{ nom: 'word/vbaProject.bin', donnees: vbaProjectBin(projet) }] : []),
    ...signatures.map((s) => ({ nom: `word/${sig[s]}`, donnees: Buffer.from('signature fictive') })),
  ]);
}

/** Classeur Excel 97-2003 (.xls) : fichier OLE avec _VBA_PROJECT_CUR. Le flux Workbook est un BIFF8 minimal. */
export function xls({ projet, signatures = [], auteur = 'Comptabilité', signatureBlob } = {}) {
  const noms = { legacy: '\u0005DigitalSignature', agile: '\u0005DigitalSignatureEx', v3: '\u0005DigitalSignatureExt' };
  const enfants = [
    { nom: 'Workbook', donnees: biffMinimal() },
    { nom: '\u0005SummaryInformation', donnees: resume(auteur) },
    ...signatures.map((s) => ({ nom: noms[s], donnees: signatureBlob ?? Buffer.from('signature fictive') })),
  ];
  if (projet) enfants.push({ nom: '_VBA_PROJECT_CUR', enfants: arbreProjet(projet) });
  return ecrireCfb({ enfants });
}

/** Document Word 97-2003 (.doc) : stockage Macros. */
export function doc({ projet, auteur = 'Secrétariat' } = {}) {
  const enfants = [{ nom: 'WordDocument', donnees: Buffer.alloc(4096, 0) }, { nom: '1Table', donnees: Buffer.alloc(512, 0) }, { nom: '\u0005SummaryInformation', donnees: resume(auteur) }];
  if (projet) enfants.push({ nom: 'Macros', enfants: arbreProjet(projet) });
  return ecrireCfb({ enfants });
}

/** Fichier chiffré à l'ouverture (OOXML protégé par mot de passe) : un fichier OLE EncryptionInfo + EncryptedPackage. */
export const ooxmlChiffre = () => ecrireCfb({ enfants: [{ nom: 'EncryptionInfo', donnees: Buffer.from('0400040040000000', 'hex') }, { nom: 'EncryptedPackage', donnees: Buffer.alloc(5000, 0x5a) }] });

function biffMinimal() {
  const r = (type, d = Buffer.alloc(0)) => { const h = Buffer.alloc(4); h.writeUInt16LE(type, 0); h.writeUInt16LE(d.length, 2); return Buffer.concat([h, d]); };
  const bof = Buffer.alloc(16); bof.writeUInt16LE(0x0600, 0); bof.writeUInt16LE(0x0005, 2);
  const b = Buffer.concat([r(0x0809, bof), r(0x000a)]);
  const out = Buffer.alloc(4096); b.copy(out); return out; // ≥ 4 096 octets : flux du FAT principal
}

// Flux \005SummaryInformation : un jeu de propriétés avec PIDSI_AUTHOR (4), PIDSI_LASTAUTHOR (8), PIDSI_LASTSAVE_DTM (13).
function resume(auteur) {
  const FMTID = Buffer.from('e0859ff2f94f6810ab9108002b27b3d9', 'hex');
  const chaine = (s) => { const t = Buffer.from(s + '\0', 'latin1'); const pad = Buffer.alloc((4 - (t.length % 4)) % 4); return Buffer.concat([u32(0x1e), u32(t.length), t, pad]); };
  const filetime = () => { const b = Buffer.alloc(12); b.writeUInt32LE(0x40, 0); b.writeBigUInt64LE((BigInt(Date.UTC(2023, 5, 15, 8, 0)) + 11644473600000n) * 10000n, 4); return b; };
  const props = [[1, Buffer.concat([u32(2), u16(1252), u16(0)])], [4, chaine(auteur)], [8, chaine(auteur)], [13, filetime()]];
  let decal = 8 + props.length * 8;
  const table = [];
  const valeurs = [];
  for (const [pid, v] of props) { table.push(u32(pid), u32(decal)); valeurs.push(v); decal += v.length; }
  const section = Buffer.concat([u32(decal), u32(props.length), ...table, ...valeurs]);
  const entete = Buffer.concat([u16(0xfffe), u16(0), u32(0x00020a00), Buffer.alloc(16), u32(1), FMTID, u32(48)]);
  return Buffer.concat([entete, section]);
}
