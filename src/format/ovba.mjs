// Projet VBA ([MS-OVBA]) : décompression, flux « dir », flux PROJECT, code des modules.
import { ErreurFormat } from './cfb.mjs';

// 2.4.1 : conteneur compressé (octet 0x01, puis morceaux de 4 096 octets au plus une fois décompressés).
export function decompresser(buf, debut = 0) {
  if (buf[debut] !== 0x01) throw new ErreurFormat('conteneur VBA compressé invalide');
  const out = [];
  let i = debut + 1;
  while (i + 2 <= buf.length) {
    const entete = buf.readUInt16LE(i);
    const fin = Math.min(i + (entete & 0x0fff) + 3, buf.length);
    const compresse = (entete & 0x8000) !== 0;
    i += 2;
    const debutMorceau = out.length;
    if (!compresse) {
      for (let k = 0; k < 4096 && i < buf.length; k++) out.push(buf[i++]);
      continue;
    }
    while (i < fin) {
      const drapeaux = buf[i++];
      for (let bit = 0; bit < 8 && i < fin; bit++) {
        if (!((drapeaux >> bit) & 1)) {
          out.push(buf[i++]);
          continue;
        }
        if (i + 2 > fin) throw new ErreurFormat('jeton de copie tronqué');
        const jeton = buf.readUInt16LE(i);
        i += 2;
        const dif = out.length - debutMorceau;
        const nbBits = Math.max(Math.ceil(Math.log2(dif)), 4);
        const masqueLg = 0xffff >> nbBits;
        const longueur = (jeton & masqueLg) + 3;
        const decalage = (jeton >> (16 - nbBits)) + 1;
        const source = out.length - decalage;
        if (source < debutMorceau) throw new ErreurFormat('jeton de copie hors du morceau');
        for (let k = 0; k < longueur; k++) out.push(out[source + k]);
      }
    }
    i = fin;
  }
  return Buffer.from(out);
}

const ETIQUETTES = { 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1200: 'utf-16le', 65001: 'utf-8', 10000: 'macintosh' };
export function decodeur(codePage) {
  const nom = codePage >= 1250 && codePage <= 1258 ? `windows-${codePage}` : ETIQUETTES[codePage] ?? 'windows-1252';
  try { return new TextDecoder(nom); } catch { return new TextDecoder('windows-1252'); }
}

// 2.3.4.2 : flux « dir » décompressé. Enregistrements (id u16, taille u32, données),
// sauf PROJECTVERSION (0x0009) dont la taille annoncée (4) ne compte pas les 2 derniers octets.
export function lireDir(dir) {
  const projet = { codePage: 1252, nom: '', references: [], modules: [] };
  let i = 0;
  let ref = null;
  let mod = null;
  const txt = (b) => decodeur(projet.codePage).decode(b);
  const libid = (b, o) => {
    const n = b.readUInt32LE(o);
    return { valeur: txt(b.subarray(o + 4, o + 4 + n)), fin: o + 4 + n };
  };
  while (i + 6 <= dir.length) {
    const id = dir.readUInt16LE(i);
    let taille = dir.readUInt32LE(i + 2);
    if (id === 0x0009) taille = 6;
    const d = dir.subarray(i + 6, i + 6 + taille);
    i += 6 + taille;
    switch (id) {
      case 0x0003: projet.codePage = d.readUInt16LE(0); break;
      case 0x0004: projet.nom = txt(d); break;
      case 0x0016: ref = { nom: txt(d), type: '', libid: '' }; projet.references.push(ref); break;
      case 0x003e: if (ref && !ref.nomU) ref.nomU = d.toString('utf16le'); break;
      case 0x000d: // REFERENCEREGISTERED
        if (!ref) { ref = { nom: '' }; projet.references.push(ref); }
        ref.type = 'registered'; ref.libid = libid(d, 0).valeur; ref = null; break;
      case 0x000e: { // REFERENCEPROJECT
        if (!ref) { ref = { nom: '' }; projet.references.push(ref); }
        const a = libid(d, 0);
        ref.type = 'project'; ref.libid = a.valeur; ref.libidRelatif = libid(d, a.fin).valeur; ref = null; break;
      }
      case 0x0033: // REFERENCEORIGINAL : précède un REFERENCECONTROL
        if (!ref) { ref = { nom: '' }; projet.references.push(ref); }
        ref.libidOriginal = txt(d); break;
      case 0x002f: // REFERENCECONTROL (partie « twiddled »)
        if (!ref) { ref = { nom: '' }; projet.references.push(ref); }
        ref.type = 'control'; ref.libid = libid(d, 0).valeur; break;
      case 0x0030: // REFERENCECONTROL (partie étendue) : fin de la référence
        if (ref) { ref.libidEtendu = libid(d, 0).valeur; ref = null; } break;
      case 0x0019: mod = { nom: txt(d), flux: '', decalage: 0, type: 'standard' }; projet.modules.push(mod); ref = null; break;
      case 0x0047: if (mod) mod.nomU = d.toString('utf16le'); break;
      case 0x001a: if (mod) mod.flux = txt(d); break;
      case 0x0032: if (mod) mod.fluxU = d.toString('utf16le'); break;
      case 0x0031: if (mod) mod.decalage = d.readUInt32LE(0); break;
      case 0x0022: if (mod) mod.type = 'document'; break; // document ou classe : précisé par l'attribut VB_Base
      case 0x0021: if (mod) mod.type = 'standard'; break;
      case 0x0010: return projet; // fin du flux dir
      default: break;
    }
  }
  return projet;
}

// 2.4.3 : « Data Encryption » des valeurs CMG / DPB / GC du flux PROJECT.
export function dechiffrer(hex) {
  const b = Buffer.from(hex.replace(/[^0-9a-f]/gi, ''), 'hex');
  if (b.length < 3) return null;
  const graine = b[0];
  const version = graine ^ b[1];
  const cle = graine ^ b[2];
  if (version !== 2) return null;
  let nonChiffre1 = cle, chiffre1 = b[2], chiffre2 = b[1];
  let i = 3;
  const suivant = () => {
    if (i >= b.length) throw new ErreurFormat('donnée chiffrée tronquée');
    const c = b[i++];
    const x = c ^ ((chiffre2 + nonChiffre1) & 0xff);
    chiffre2 = chiffre1; chiffre1 = c; nonChiffre1 = x;
    return x;
  };
  try {
    for (let k = 0; k < (graine & 6) / 2; k++) suivant();
    const lg = suivant() | (suivant() << 8) | (suivant() << 16) | (suivant() << 24);
    if (lg < 0 || lg > 4096) return null;
    const out = Buffer.alloc(lg);
    for (let k = 0; k < lg; k++) out[k] = suivant();
    return out;
  } catch { return null; }
}

// Flux PROJECT (texte) : modules déclarés, protection et mot de passe.
export function lireProject(buf, codePage = 1252) {
  const texte = decodeur(codePage).decode(buf);
  const valeurs = {};
  const formulaires = [];
  for (const ligne of texte.split(/\r?\n/)) {
    if (ligne.startsWith('[')) break;
    const m = /^([A-Za-z0-9]+)=(.*)$/.exec(ligne);
    if (!m) continue;
    if (m[1] === 'BaseClass' || m[1] === 'Package') formulaires.push(m[2]);
    valeurs[m[1]] = m[2].replace(/^"|"$/g, '');
  }
  const etat = valeurs.CMG ? dechiffrer(valeurs.CMG) : null;
  const motDePasse = valeurs.DPB ? dechiffrer(valeurs.DPB) : null;
  const visibilite = valeurs.GC ? dechiffrer(valeurs.GC) : null;
  const etatProtection = etat && etat.length >= 4 ? etat.readUInt32LE(0) : 0;
  return {
    protege: Boolean((motDePasse && !(motDePasse.length === 1 && motDePasse[0] === 0)) || (etatProtection & 1) || (visibilite && visibilite.length === 1 && visibilite[0] === 0)),
    formulaires: texte.split(/\r?\n/).filter((l) => /^BaseClass=/i.test(l)).map((l) => l.slice(10)),
  };
}

// Projet complet depuis un stockage qui contient PROJECT et VBA/dir.
export function lireProjetVba(stockage) {
  const vba = stockage.enfant('VBA');
  const dirFlux = vba?.enfant('dir');
  if (!dirFlux) throw new ErreurFormat('flux VBA/dir absent');
  const projet = lireDir(decompresser(dirFlux.lire()));
  const pj = stockage.enfant('PROJECT');
  const infos = pj ? lireProject(pj.lire(), projet.codePage) : { protege: false, formulaires: [] };
  const dec = decodeur(projet.codePage);
  const modules = [];
  const erreurs = [];
  for (const m of projet.modules) {
    const nomFlux = m.fluxU || m.flux || m.nom;
    const flux = vba.enfant(nomFlux);
    if (!flux) { erreurs.push(`module ${m.nom} : flux ${nomFlux} absent`); continue; }
    try {
      const code = dec.decode(decompresser(flux.lire(), m.decalage));
      modules.push({ nom: m.nomU || m.nom, type: typeModule(m.type, code), code });
    } catch (e) {
      erreurs.push(`module ${m.nom} : ${e.message}`);
    }
  }
  return { nom: projet.nom, codePage: projet.codePage, references: projet.references.map(nettoyerRef), modules, protege: infos.protege, erreurs };
}

// Classe, formulaire, document (feuille, classeur) ou module standard, d'après les attributs du code.
function typeModule(type, code) {
  if (type === 'standard') return 'module';
  const base = /^Attribute VB_Base = "(.*)"/im.exec(code);
  if (!base) return 'classe';
  return /C62A69F0-16DC-11CE-9E98-00AA00574A4F/i.test(base[1]) ? 'formulaire' : 'document';
}

function nettoyerRef(r) {
  const libid = r.libidEtendu || r.libid || '';
  // *\G{GUID}#version#lcid#chemin#description
  const m = /^\*\\[GH]\{?([0-9A-Fa-f-]{36})\}?#([^#]*)#([^#]*)#([^#]*)#(.*)$/.exec(libid);
  if (m) return { nom: r.nomU || r.nom, type: r.type, guid: m[1].toUpperCase(), version: m[2], chemin: m[4], description: m[5] };
  // référence de projet : *\C<chemin> ou *\D<chemin>
  const p = /^\*\\[CD](.*)$/.exec(libid);
  return { nom: r.nomU || r.nom, type: r.type, guid: '', version: '', chemin: p ? p[1] : libid, description: '' };
}
