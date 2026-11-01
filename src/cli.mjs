// Ligne de commande : inventaire-vba <dossier…> [options]
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { createInterface } from 'node:readline/promises';
import { spawn } from 'node:child_process';
import path from 'node:path';
import { inventorier, synthese, anonymiser, LIMITE_GRATUITE } from './scan.mjs';
import { rapportHtml } from './rapport/html.mjs';
import { VERSION } from './version.mjs';

const AIDE = `inventaire-vba ${VERSION} — inventaire des macros VBA d'un parc Office, hors ligne.
Lit aussi les projets Visual Basic 6 (.vbp) et repère l'envoi de courriels (Outlook, SMTP) dans les sources .NET.

Usage :
  inventaire-vba <dossier ou partage> [autres dossiers] [options]
  inventaire-vba                        (sans argument : demande le dossier)

Exemples :
  inventaire-vba \\\\serveur\\partages\\Finance -o C:\\Inventaires\\finance
  inventaire-vba D:\\Commun --exclure Archives --organisation "Ville de Exemple"
  inventaire-vba D:\\Commun --diagnostic          (fichier à envoyer pour un diagnostic à distance)

Options :
  -o, --sortie <dossier>     dossier des résultats (défaut : inventaire-vba-AAAA-MM-JJ)
  --organisation <nom>       nom affiché sur le rapport
  --exclure <texte>          ignore les chemins qui contiennent ce texte (répétable)
  --max-taille <Mo>          ignore les fichiers plus gros (défaut : 200)
  --diagnostic               écrit aussi diagnostic.json : sans code, sans auteurs, à envoyer
  --sans-chemins             avec --diagnostic : remplace aussi noms de fichiers, dossiers et modules
  --depuis <inventaire.json> refait les rapports sans relire les fichiers
  --comparer <inventaire.json> compare à un inventaire précédent (édition parc)
  --revue <revue.md>         intègre la revue d'un expert (diagnostic)
  --date <AAAA-MM-JJ>        date de référence (expiration des certificats)
  --pdf                      imprime aussi rapport.pdf (Chromium via Playwright requis)
  --ouvrir                   ouvre le rapport à la fin
  -q, --silencieux           sans progression
  -h, --aide                 cette aide
  -v, --version              version

Rien n'est envoyé sur le réseau, rien n'est exécuté : les fichiers sont lus, jamais modifiés.
Résultats : rapport.html, inventaire.json${'\n'}  + édition parc : inventaire.xlsx, fichiers.csv, constats.csv
Documentation et licence de parc : https://consigne-studio.github.io/inventaire-vba/`;

export function lireArgs(argv) {
  const o = { dossiers: [], exclure: [], maxTaille: 200 };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const val = () => { if (i + 1 >= argv.length) throw new Error(`${a} attend une valeur`); return argv[++i]; };
    switch (a) {
      case '-o': case '--sortie': o.sortie = val(); break;
      case '--organisation': o.organisation = val(); break;
      case '--exclure': o.exclure.push(val()); break;
      case '--max-taille': o.maxTaille = Number(val()); if (!(o.maxTaille > 0)) throw new Error('--max-taille attend un nombre de Mo'); break;
      case '--diagnostic': o.diagnostic = true; break;
      case '--sans-chemins': o.sansChemins = true; break;
      case '--depuis': o.depuis = val(); break;
      case '--comparer': o.comparer = val(); break;
      case '--revue': o.revue = val(); break;
      case '--date': o.date = val(); if (!/^\d{4}-\d{2}-\d{2}$/.test(o.date)) throw new Error('--date attend AAAA-MM-JJ'); break;
      case '--pdf': o.pdf = true; break;
      case '--ouvrir': o.ouvrir = true; break;
      case '-q': case '--silencieux': o.silencieux = true; break;
      case '-h': case '--aide': case '--help': case '/?': o.aide = true; break;
      case '-v': case '--version': o.version = true; break;
      default:
        if (a.startsWith('-') && a.length > 1 && !/^[a-z]:/i.test(a)) throw new Error(`option inconnue : ${a} (voir --aide)`);
        o.dossiers.push(a);
    }
  }
  if (o.sansChemins && !o.diagnostic) throw new Error('--sans-chemins s\'utilise avec --diagnostic');
  return o;
}

// Console interactive (double-clic sur l'exécutable). L'accès à stdin peut échouer sur certains hôtes : alors non.
function estInteractif() {
  try { return Boolean(process.stdin.isTTY && process.stdout.isTTY); } catch { return false; }
}

async function chargerPro() {
  try { return await import('./pro/index.mjs'); } catch (e) { if (e.code === 'ERR_MODULE_NOT_FOUND' || /pro\/index/.test(e.message)) return null; throw e; }
}

export async function principal(argv, { sortieStd = process.stdout, sortieErr = process.stderr, interactif = estInteractif() } = {}) {
  const o = lireArgs(argv);
  const pro = await chargerPro();
  const edition = pro ? 'parc' : 'gratuite';
  if (o.version) { sortieStd.write(`inventaire-vba ${VERSION} (édition ${edition})\n`); return 0; }
  if (o.aide) { sortieStd.write(AIDE + '\n'); return 0; }
  const dire = (s) => { if (!o.silencieux) sortieErr.write(s + '\n'); };
  let pause = false;
  if (!o.dossiers.length && !o.depuis) {
    if (!interactif) { sortieErr.write('Indiquez un dossier à analyser (voir --aide).\n'); return 1; }
    const rl = createInterface({ input: process.stdin, output: process.stdout });
    sortieStd.write(`inventaire-vba ${VERSION} — édition ${edition}\nAucun fichier n'est modifié ni envoyé.\n\n`);
    const rep = (await rl.question('Dossier ou partage à analyser (glissez-le ici, puis Entrée) : ')).trim().replace(/^"(.*)"$/, '$1');
    rl.close();
    if (!rep) return 1;
    o.dossiers.push(rep);
    o.ouvrir = true;
    pause = true;
  }
  if (o.comparer && !pro) throw new Error('--comparer fait partie de l\'édition parc');
  for (const d of o.dossiers) if (!existsSync(d)) throw new Error(`introuvable : ${d}`);

  const maintenant = o.date ? new Date(`${o.date}T12:00:00Z`) : new Date();
  let inv;
  if (o.depuis) {
    inv = JSON.parse(await readFile(o.depuis, 'utf8'));
    if (inv.schema !== 'inventaire-vba/1') throw new Error(`${o.depuis} n'est pas un inventaire-vba`);
  } else {
    const limite = pro || o.diagnostic ? Infinity : LIMITE_GRATUITE;
    let dernier = 0;
    inv = await inventorier(o.dossiers, {
      maintenant, exclure: o.exclure, maxTaille: o.maxTaille * 1024 * 1024, limite, edition,
      progression: (n, total) => {
        if (o.silencieux || (n !== total && Date.now() - dernier < 250)) return;
        dernier = Date.now();
        sortieErr.write(`\rAnalyse : ${n} / ${total} fichiers${n === total ? '\n' : ''}`);
      },
    });
  }
  const sortie = path.resolve(o.sortie ?? `inventaire-vba-${maintenant.toISOString().slice(0, 10)}`);
  await mkdir(sortie, { recursive: true });
  const syn = synthese(inv);
  const revue = o.revue ? await readFile(o.revue, 'utf8') : null;
  const precedent = o.comparer ? JSON.parse(await readFile(o.comparer, 'utf8')) : null;
  const extra = pro ? pro.enrichir(inv, syn, { precedent }) : null;
  const ecrits = [];
  const ecrire = async (nom, donnees) => { await writeFile(path.join(sortie, nom), donnees); ecrits.push(nom); };

  await ecrire('rapport.html', rapportHtml(inv, syn, { pro: extra && { ...extra.html, parFichier: extra.parFichier }, revue, organisation: o.organisation, limite: pro || o.diagnostic ? null : LIMITE_GRATUITE }));
  if (!o.depuis) await ecrire('inventaire.json', JSON.stringify(inv, null, 1));
  if (pro) {
    await ecrire('inventaire.xlsx', pro.classeur(inv, syn, extra));
    const v = pro.vues(inv, syn, extra);
    await ecrire('fichiers.csv', pro.csv(v.fichiers));
    await ecrire('constats.csv', pro.csv(v.constats));
  }
  if (o.diagnostic) {
    const anon = anonymiser(inv, { sansChemins: o.sansChemins });
    await ecrire('diagnostic.json', JSON.stringify(anon, null, 1));
    // Table de correspondance alias → chemin réel : reste chez le client, pour relire le rapport livré.
    if (o.sansChemins) await ecrire('diagnostic-correspondance.csv', '\ufeffalias;chemin\r\n' + inv.fichiers.map((f, k) => `${anon.fichiers[k].chemin};${f.chemin}`).join('\r\n') + '\r\n');
    await ecrire('diagnostic-apercu.html', rapportHtml(anon, synthese(anon), { organisation: 'Aperçu du fichier de diagnostic' }));
  }
  if (o.pdf) {
    const { imprimerPdf } = await import('./rapport/pdf.mjs');
    await imprimerPdf(path.join(sortie, 'rapport.html'), path.join(sortie, 'rapport.pdf'));
    ecrits.push('rapport.pdf');
  }

  dire(`\n${syn.fichiersAvecMacros} fichiers avec macros sur ${syn.fichiersAnalyses} analysés · ${syn.lignesCode} lignes de VBA`);
  dire(`P1 (cassent) : ${syn.fichiersP1} fichiers · P2 (avant 2027) : ${syn.fichiersP2} fichiers`);
  if (syn.vb6.projets) dire(`Applications VB6 : ${syn.vb6.projets} projet(s) · ${syn.vb6.lignes} lignes`);
  if (syn.courriel.fichiers) dire(`Envoi de courriels : ${syn.courriel.fichiers} fichier(s), dont ${syn.courriel.smtp} en SMTP par mot de passe (fin 2026) et ${syn.courriel.outlook} par Outlook (nouvel Outlook) — kit dans remediation\\`);
  if (syn.nonAnalyses) dire(`Édition gratuite : ${syn.nonAnalyses} fichiers au-delà de ${LIMITE_GRATUITE} n'ont pas été analysés (liste dans le rapport).`);
  if (inv.inaccessibles.length) dire(`${inv.inaccessibles.length} chemins inaccessibles (droits insuffisants), listés dans le rapport.`);
  dire(`\nRésultats dans ${sortie}\n  ${ecrits.join('\n  ')}`);
  if (o.diagnostic) dire(`\nÀ envoyer pour le diagnostic : diagnostic.json (sans code ni auteurs${o.sansChemins ? ', ni noms de fichiers' : ''}). Ouvrez diagnostic-apercu.html pour voir exactement ce qu'il contient.${o.sansChemins ? ' Gardez diagnostic-correspondance.csv : il relie chaque alias à son fichier et ne doit pas être envoyé.' : ''}`);
  if (o.ouvrir) ouvrir(path.join(sortie, 'rapport.html'));
  if (pause && interactif) {
    const rl = createInterface({ input: process.stdin, output: process.stdout });
    await rl.question('\nAppuyez sur Entrée pour fermer.');
    rl.close();
  }
  return 0;
}

function ouvrir(fichier) {
  const cmd = process.platform === 'win32' ? ['cmd', ['/c', 'start', '""', fichier]] : process.platform === 'darwin' ? ['open', [fichier]] : ['xdg-open', [fichier]];
  try { spawn(cmd[0], cmd[1], { detached: true, stdio: 'ignore', windowsHide: true }).on('error', () => {}).unref(); } catch { /* pas de navigateur : le chemin est affiché */ }
}
