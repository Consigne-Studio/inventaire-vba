// Régénère le rapport d'exemple publié : parc fictif « Exemple Industrie », date fixe (reproductible).
// Usage : node scripts/exemple.mjs   → exemples/exemple-industrie/
import { rm, mkdir, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { genererParc } from '../test/outils/parc-exemple.mjs';
import { principal } from '../src/cli.mjs';

const RACINE = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const parc = path.join(tmpdir(), 'inventaire-vba-exemple', 'Exemple-Industrie');
const sortie = path.join(RACINE, 'exemples', 'exemple-industrie');
await rm(path.dirname(parc), { recursive: true, force: true });
await genererParc(parc);
await rm(sortie, { recursive: true, force: true });
await mkdir(sortie, { recursive: true });
const code = await principal([parc, '-o', sortie, '--date', '2026-10-23', '--organisation', 'Exemple Industrie (organisation fictive)', '--diagnostic', '-q'], { interactif: false });
// Chemins absolus de la machine de génération : remplacés pour un fichier publiable.
for (const f of ['inventaire.json', 'diagnostic.json']) {
  const p = path.join(sortie, f);
  const j = JSON.parse(await readFile(p, 'utf8'));
  j.racines = j.racines.map(() => 'D:\\Partages\\Exemple-Industrie');
  await writeFile(p, JSON.stringify(j, null, 1));
}
await rm(path.dirname(parc), { recursive: true, force: true });
// Livrable d'un diagnostic à distance : rapport refait depuis diagnostic.json, avec la revue de l'expert, en PDF.
const livre = path.join(sortie, 'diagnostic-livre');
await principal(['--depuis', path.join(sortie, 'diagnostic.json'), '--revue', path.join(RACINE, 'exemples', 'revue-exemple-industrie.md'), '-o', livre, '--date', '2026-10-23', '--organisation', 'Exemple Industrie (organisation fictive)', '--pdf', '-q'], { interactif: false });
console.log(`Rapport d'exemple : ${path.relative(RACINE, sortie)}/rapport.html`);
process.exitCode = code;
