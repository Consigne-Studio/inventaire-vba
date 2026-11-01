// Archives ZIP (formats Office Open XML) : lecture des entrées et écriture du classeur Excel du rapport.
import zlib from 'node:zlib';
import { ErreurFormat } from './cfb.mjs';

export const estZip = (buf) => buf.length >= 22 && buf.readUInt32LE(0) === 0x04034b50;

export function lireZip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new ErreurFormat('fin d\'archive ZIP introuvable');
  const nb = buf.readUInt16LE(eocd + 10);
  let o = buf.readUInt32LE(eocd + 16);
  const entrees = new Map();
  for (let k = 0; k < nb; k++) {
    if (o + 46 > buf.length || buf.readUInt32LE(o) !== 0x02014b50) throw new ErreurFormat('répertoire ZIP corrompu');
    const drapeaux = buf.readUInt16LE(o + 8);
    const methode = buf.readUInt16LE(o + 10);
    const tailleC = buf.readUInt32LE(o + 20);
    const taille = buf.readUInt32LE(o + 24);
    const ln = buf.readUInt16LE(o + 28), le = buf.readUInt16LE(o + 30), lc = buf.readUInt16LE(o + 32);
    const local = buf.readUInt32LE(o + 42);
    const nom = buf.toString(drapeaux & 0x800 ? 'utf8' : 'latin1', o + 46, o + 46 + ln);
    entrees.set(nom, { nom, methode, tailleC, taille, local, chiffre: Boolean(drapeaux & 1) });
    o += 46 + ln + le + lc;
  }
  const lire = (nom) => {
    const e = typeof nom === 'string' ? entrees.get(nom) : nom;
    if (!e) return null;
    if (e.chiffre) throw new ErreurFormat(`entrée chiffrée : ${e.nom}`);
    if (e.taille > 512 * 1024 * 1024) throw new ErreurFormat(`entrée trop grande : ${e.nom}`);
    if (buf.readUInt32LE(e.local) !== 0x04034b50) throw new ErreurFormat(`en-tête local corrompu : ${e.nom}`);
    const debut = e.local + 30 + buf.readUInt16LE(e.local + 26) + buf.readUInt16LE(e.local + 28);
    const brut = buf.subarray(debut, debut + e.tailleC);
    if (e.methode === 0) return brut;
    if (e.methode === 8) return zlib.inflateRawSync(brut, { maxOutputLength: Math.max(e.taille, 1) * 2 + 1024 });
    throw new ErreurFormat(`compression ZIP non prise en charge (${e.methode})`);
  };
  return { noms: [...entrees.keys()], lire, texte: (nom) => lire(nom)?.toString('utf8') ?? null };
}

const TABLE = new Int32Array(256).map((_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c;
});
export function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

// Écrit une archive à partir de [{ nom, donnees }]. Date fixe : sortie reproductible.
export function ecrireZip(fichiers) {
  const locaux = [];
  const central = [];
  let decalage = 0;
  const DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;
  for (const f of fichiers) {
    const donnees = Buffer.isBuffer(f.donnees) ? f.donnees : Buffer.from(f.donnees, 'utf8');
    const comp = zlib.deflateRawSync(donnees, { level: 9 });
    const nom = Buffer.from(f.nom, 'utf8');
    const crc = crc32(donnees);
    const l = Buffer.alloc(30);
    l.writeUInt32LE(0x04034b50, 0); l.writeUInt16LE(20, 4); l.writeUInt16LE(0x800, 6); l.writeUInt16LE(8, 8);
    l.writeUInt16LE(0, 10); l.writeUInt16LE(DATE, 12); l.writeUInt32LE(crc, 14);
    l.writeUInt32LE(comp.length, 18); l.writeUInt32LE(donnees.length, 22); l.writeUInt16LE(nom.length, 26);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(0x800, 8); c.writeUInt16LE(8, 10);
    c.writeUInt16LE(0, 12); c.writeUInt16LE(DATE, 14); c.writeUInt32LE(crc, 16); c.writeUInt32LE(comp.length, 20);
    c.writeUInt32LE(donnees.length, 24); c.writeUInt16LE(nom.length, 28); c.writeUInt32LE(decalage, 42);
    locaux.push(l, nom, comp);
    central.push(c, nom);
    decalage += 30 + nom.length + comp.length;
  }
  const cd = Buffer.concat(central);
  const fin = Buffer.alloc(22);
  fin.writeUInt32LE(0x06054b50, 0); fin.writeUInt16LE(fichiers.length, 8); fin.writeUInt16LE(fichiers.length, 10);
  fin.writeUInt32LE(cd.length, 12); fin.writeUInt32LE(decalage, 16);
  return Buffer.concat([...locaux, cd, fin]);
}
