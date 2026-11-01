// Lecture d'un fichier composé OLE (Compound File Binary, [MS-CFB]) : .xls, .doc, vbaProject.bin…
// Lecture seule, en mémoire, sans dépendance. Les erreurs de structure lèvent une ErreurFormat.

export class ErreurFormat extends Error {}

const SIGNATURE = Buffer.from('d0cf11e0a1b11ae1', 'hex');
const FIN_CHAINE = 0xfffffffe;
const LIBRE = 0xffffffff;
const MAX_SECTEURS = 0xfffffffa;

export const estCfb = (buf) => buf.length >= 512 && buf.subarray(0, 8).equals(SIGNATURE);

export function lireCfb(buf) {
  if (!estCfb(buf)) throw new ErreurFormat('signature OLE absente');
  const major = buf.readUInt16LE(0x1a);
  const tailleSecteur = 1 << buf.readUInt16LE(0x1e);
  const tailleMini = 1 << buf.readUInt16LE(0x20);
  if (![512, 4096].includes(tailleSecteur) || tailleMini !== 64) throw new ErreurFormat('taille de secteur invalide');
  const nbFat = buf.readUInt32LE(0x2c);
  const premierDir = buf.readUInt32LE(0x30);
  const coupure = buf.readUInt32LE(0x38);
  const premierMiniFat = buf.readUInt32LE(0x3c);
  const premierDifat = buf.readUInt32LE(0x44);
  const nbDifat = buf.readUInt32LE(0x48);
  const nbSecteurs = Math.floor((buf.length - tailleSecteur) / tailleSecteur) + 1;

  const secteur = (n) => {
    const debut = (n + 1) * tailleSecteur;
    if (n >= MAX_SECTEURS || debut >= buf.length) throw new ErreurFormat(`secteur ${n} hors du fichier`);
    return buf.subarray(debut, Math.min(debut + tailleSecteur, buf.length));
  };

  // Table DIFAT : 109 entrées dans l'en-tête, puis des secteurs chaînés.
  const difat = [];
  for (let i = 0; i < 109; i++) difat.push(buf.readUInt32LE(0x4c + i * 4));
  let d = premierDifat;
  for (let i = 0; i < nbDifat && d < MAX_SECTEURS; i++) {
    const s = secteur(d);
    for (let j = 0; j < tailleSecteur / 4 - 1; j++) difat.push(s.readUInt32LE(j * 4));
    d = s.readUInt32LE(tailleSecteur - 4);
  }
  const fat = [];
  for (const n of difat.slice(0, nbFat)) {
    if (n >= MAX_SECTEURS) continue;
    const s = secteur(n);
    for (let j = 0; j < s.length / 4; j++) fat.push(s.readUInt32LE(j * 4));
  }

  const chaine = (debut, table, limite) => {
    const out = [];
    const vus = new Set();
    for (let n = debut; n !== FIN_CHAINE && n !== LIBRE; n = table[n]) {
      if (n === undefined || n >= table.length || vus.has(n) || out.length > limite) throw new ErreurFormat('chaîne de secteurs corrompue');
      vus.add(n);
      out.push(n);
    }
    return out;
  };
  const lireChaine = (debut) => Buffer.concat(chaine(debut, fat, nbSecteurs).map(secteur));

  // Répertoire : entrées de 128 octets.
  const dir = lireChaine(premierDir);
  const entrees = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const lg = dir.readUInt16LE(o + 64);
    const type = dir[o + 66];
    entrees.push({
      nom: lg >= 2 ? dir.toString('utf16le', o, o + Math.min(lg, 64) - 2) : '',
      type, // 0 vide, 1 stockage, 2 flux, 5 racine
      gauche: dir.readUInt32LE(o + 68),
      droite: dir.readUInt32LE(o + 72),
      enfant: dir.readUInt32LE(o + 76),
      debut: dir.readUInt32LE(o + 116),
      taille: major === 3 ? dir.readUInt32LE(o + 120) : Number(dir.readBigUInt64LE(o + 120)),
    });
  }
  const racine = entrees[0];
  if (!racine || racine.type !== 5) throw new ErreurFormat('entrée racine absente');

  // Mini-flux (petits flux de moins de `coupure` octets).
  let miniFat = [];
  let miniFlux = Buffer.alloc(0);
  if (premierMiniFat < MAX_SECTEURS) {
    const mf = lireChaine(premierMiniFat);
    for (let j = 0; j < mf.length / 4; j++) miniFat.push(mf.readUInt32LE(j * 4));
    miniFlux = racine.debut < MAX_SECTEURS ? lireChaine(racine.debut) : miniFlux;
  }

  const contenu = (e) => {
    if (e.taille === 0) return Buffer.alloc(0);
    if (e.taille < coupure) {
      const parts = chaine(e.debut, miniFat, miniFat.length).map((n) => {
        if ((n + 1) * tailleMini > miniFlux.length) throw new ErreurFormat('mini-secteur hors du mini-flux');
        return miniFlux.subarray(n * tailleMini, (n + 1) * tailleMini);
      });
      return Buffer.concat(parts).subarray(0, e.taille);
    }
    return lireChaine(e.debut).subarray(0, e.taille);
  };

  // Arbre : les enfants d'un stockage forment un arbre binaire (gauche / droite).
  const enfants = (e) => {
    const out = [];
    const pile = [e.enfant];
    const vus = new Set();
    while (pile.length) {
      const i = pile.pop();
      if (i === LIBRE || i >= entrees.length || vus.has(i)) continue;
      vus.add(i);
      const x = entrees[i];
      out.push(x);
      pile.push(x.gauche, x.droite);
    }
    return out;
  };

  const noeud = (e) => ({
    nom: e.nom,
    estStockage: e.type === 1 || e.type === 5,
    taille: e.taille,
    enfants: () => enfants(e).filter((x) => x.type).map(noeud),
    enfant(nom) {
      const x = enfants(e).find((c) => c.type && c.nom.toLowerCase() === nom.toLowerCase());
      return x ? noeud(x) : null;
    },
    lire: () => contenu(e),
  });

  return { racine: noeud(racine) };
}

// Chemin « A/B/C » depuis la racine, insensible à la casse.
export function trouver(cfb, chemin) {
  let n = cfb.racine;
  for (const p of chemin.split('/').filter(Boolean)) {
    n = n?.enfant(p);
    if (!n) return null;
  }
  return n;
}
