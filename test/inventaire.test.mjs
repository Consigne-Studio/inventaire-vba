import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { execFileSync, spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { decompresser } from '../src/format/ovba.mjs';
import { lireCfb } from '../src/format/cfb.mjs';
import { lireZip } from '../src/format/zip.mjs';
import { lireProjetVba } from '../src/format/ovba.mjs';
import { analyserModule, REGLES } from '../src/regles.mjs';
import { inventorier, synthese, anonymiser, analyserFichier } from '../src/scan.mjs';
import { principal } from '../src/cli.mjs';
import * as F from './outils/fabrique.mjs';
import { genererParc } from './outils/parc-exemple.mjs';

const ICI = path.dirname(fileURLToPath(import.meta.url));
const DATE = new Date('2026-10-23T12:00:00Z');
const tmp = () => mkdtemp(path.join(tmpdir(), 'inventaire-vba-'));
const regles = (f) => [...new Set(f.constats.map((c) => c.regle))].sort();
const muet = { write() {} };
const PRO = existsSync(path.join(ICI, '..', 'src', 'pro', 'index.mjs'));

test('compression VBA : aller-retour sur texte, octets aléatoires et morceaux multiples', () => {
  const texte = Buffer.from('Sub Bonjour()\r\n    MsgBox "Bonjour"\r\nEnd Sub\r\n'.repeat(300), 'latin1');
  const alea = Buffer.from(Array.from({ length: 9000 }, (_, i) => (i * 7919 + (i >> 3)) & 0xff));
  for (const b of [texte, alea, Buffer.alloc(0), Buffer.from('a'), Buffer.alloc(4096, 0x41), Buffer.alloc(4097, 0x42)]) {
    assert.deepEqual(decompresser(F.compresser(b)), b);
  }
  assert.ok(F.compresser(texte).length < texte.length / 5, 'le texte répétitif est réellement compressé');
});

test('fichier composé OLE : flux du mini-flux et du FAT principal', () => {
  const gros = Buffer.alloc(10000, 7);
  const cfb = lireCfb(F.ecrireCfb({ enfants: [{ nom: 'Petit', donnees: Buffer.from('abc') }, { nom: 'Dossier', enfants: [{ nom: 'Gros', donnees: gros }, { nom: 'Vide', donnees: Buffer.alloc(0) }] }] }));
  assert.equal(cfb.racine.enfant('petit').lire().toString(), 'abc');
  assert.deepEqual(cfb.racine.enfant('Dossier').enfant('Gros').lire(), gros);
  assert.equal(cfb.racine.enfant('Dossier').enfant('Vide').lire().length, 0);
});

test('projet VBA réel enregistré par Excel (exemple de XlsxWriter, BSD)', async () => {
  const p = lireProjetVba(lireCfb(await readFile(path.join(ICI, 'reels', 'xlsxwriter-vbaProject.bin'))).racine);
  assert.equal(p.nom, 'VBAProject');
  assert.equal(p.protege, false);
  assert.deepEqual(p.modules.map((m) => `${m.nom}:${m.type}`), ['ThisWorkbook:document', 'Sheet1:document', 'Module1:module', 'ThisWorkbook1:document', 'Sheet2:document']);
  assert.match(p.modules[2].code, /Sub say_hello\(\)\r\n {4}MsgBox \("Hello from Python!"\)/);
  assert.deepEqual(p.references.map((r) => r.nom), ['stdole', 'Office']);
});

test('contre-vérification indépendante avec olevba (oletools), si installé', { skip: spawnSync('olevba', ['--version']).error ? 'olevba absent' : false }, async () => {
  const d = await tmp();
  const f = path.join(d, 'x.xlsm');
  await writeFile(f, F.xlsm({ projet: { modules: [{ nom: 'Module1', code: 'Sub A()\n  Set re = CreateObject("VBScript.RegExp")\nEnd Sub\n'.repeat(200) }] } }));
  const sortie = execFileSync('olevba', ['-c', f], { encoding: 'utf8' });
  assert.equal((sortie.match(/Sub A\(\)/g) ?? []).length, 200);
  await rm(d, { recursive: true });
});

test('lecture du code : commentaires, Rem, suites de ligne, #If VBA7, numéros de ligne de l\'éditeur', () => {
  const code = `Attribute VB_Name = "M"
Option Explicit
' Set r = CreateObject("VBScript.RegExp")  commentaire : ignoré
Rem Declare Function X Lib "y" () As Long
#If VBA7 Then
  Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
  Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If
#If Not VBA7 Then
  Private Declare Function A1 Lib "a" () As Long
#End If
Private Declare Function SansGarde Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal n As Long) As Long
Sub T()
  s = "texte avec ' apostrophe" : Set o = CreateObject("Outlook.Application") ' commentaire
End Sub
`;
  const a = analyserModule({ nom: 'M', code });
  assert.deepEqual(a.constats.map((c) => `${c.regle}@${c.ligne}`), ['B64-001@12', 'B64-002@14', 'OUT-001@16']);
  assert.equal(a.declares, 5);
  assert.equal(a.procedures, 1);
  assert.deepEqual(a.progIds, ['Outlook.Application']);
});

test('règles : faux amis et cas limites', () => {
  const r = (code) => analyserModule({ nom: 'M', code }).constats.map((c) => c.regle);
  assert.deepEqual(r('Dim Shell As Object\nSet Shell = CreateObject("Shell.Application")\n'), []);
  assert.deepEqual(r('Shell "notepad.exe"\n'), ['SEC-001']);
  assert.deepEqual(r('x = "Le fichier rapport.vbsx"\n'), []);
  assert.deepEqual(r('Shell "wscript.exe C:\\outils\\a.vbs"\n').sort(), ['CHM-001', 'SEC-001', 'VBS-002']);
  assert.deepEqual(r('p = "C:\\Program Files\\App\\a.exe"\n'), []);
  assert.deepEqual(r('p = "\\\\srv\\partage\\a.xlsx"\n'), ['CHM-001']);
  assert.deepEqual(r('cn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=x.mdb"\n').sort(), ['B64-004', 'CNX-001']);
  assert.deepEqual(r('Dim re As New RegExp\n'), []); // VBS-004 est décidé au niveau du projet (référence absente)
});

test('parc d\'exemple : chaque fichier déclenche exactement les règles attendues', async () => {
  const d = await tmp();
  await genererParc(path.join(d, 'Parc'));
  const inv = await inventorier([path.join(d, 'Parc')], { maintenant: DATE });
  const par = Object.fromEntries(inv.fichiers.map((f) => [f.chemin.replace(/^Parc\//, ''), regles(f)]));
  const attendu = {
    'Achats/2025/Commandes fournisseurs v2.xlsm': ['IE-001', 'SEC-002'],
    'Achats/Commandes fournisseurs - Copie.xlsm': ['IE-001', 'SEC-002'],
    'Achats/Commandes fournisseurs.xlsm': ['IE-001', 'SEC-002'],
    'Archives/Ancien budget protégé.xlsm': ['FMT-002'],
    'Commun/Calculatrice TVA.xlsm': [],
    'Commun/Macros/ExportPDF.bas': ['B64-002'],
    'Commun/Macros/OutilsCommuns.xlam': ['VBS-004'],
    'Commun/Macros/Rapport qualité.xlsm': ['REF-001', 'SIG-001', 'SIG-003'],
    'Comptabilité/Relances clients.xlsm': ['MEL-001'],
    'Direction/Indicateurs mensuels.xlsm': ['B64-001', 'CHM-001', 'DDE-001', 'OUT-001', 'PRJ-001', 'SEC-003', 'XLM-001'],
    'Finance/Anciens/Tableau de bord 2017.xls': ['B64-001', 'CHM-001', 'FMT-001'],
    'Finance/Budget/Suivi budgétaire 2024.xlsm': ['B64-001', 'CHM-001', 'OUT-001', 'SEC-003', 'SIG-001', 'SIG-002'],
    'Finance/Calcul des primes.xlsm': ['B64-001', 'VBS-003'],
    'Finance/Rapprochement bancaire.xlsm': ['CNX-001', 'SEC-003', 'VBS-001'],
    'Informatique/Outils/Relances/App.config': ['MEL-001'],
    'Informatique/Outils/Relances/Envoi.cs': ['MEL-001'],
    'Informatique/Outils/Relances/Relances.csproj': ['OUT-001'],
    'Production/Applications/BancTest/BancTest.vbp': ['CHM-001', 'OUT-001', 'VB6-001', 'VB6-002', 'VB6-004', 'VB6-010'],
    'Production/Arrêts machines.xlsm': ['B64-003'],
    'Production/Copie abîmée.xlsm': ['FMT-003'],
    'Production/Relevés capteurs ligne 2.xls': ['DDE-001', 'FMT-001', 'XLM-002'],
    'Production/Suivi des ordres de fabrication.xlsm': ['ACX-001', 'B64-004', 'CHM-001', 'CNX-001', 'SEC-001', 'SEC-003', 'VBS-002'],
    'Qualité/Annuaire fournisseurs.accdb': [],
    'Qualité/Base des non-conformités.accdb': ['ACC-001'],
    'Qualité/Export Access/modAlertes.bas': ['OUT-002'],
    'RH/Modèles/Attestation employeur.doc': ['FMT-001', 'SEC-003'],
    'RH/Modèles/Courrier type.docm': ['CHM-001', 'OUT-001', 'SEC-003'],
    'RH/Planning des congés.xlsm': ['B64-003', 'CHM-001', 'PRJ-001'],
    'Scripts/connexion.vbs': ['CHM-001', 'VBS-010'],
    'Scripts/sauvegarde_nuit.vbs': ['CHM-001', 'VBS-010'],
    'Téléchargements/Facture_2024_0172.docm': ['SEC-001', 'SEC-002', 'SEC-003', 'SEC-010'],
  };
  assert.deepEqual(par, attendu);
  const s = synthese(inv);
  assert.equal(s.fichiersSansMacroComptes, 10, 'les .xlsx/.docx/.pptx sont comptés sans être ouverts ; ~$ ignoré');
  assert.equal(s.fichiersAnalyses, 31, 'Config.cs, sans constat, n\'est pas gardé ; les sources du .vbp ne sont pas comptées à part');
  assert.deepEqual(inv.annexesLues, { dotnet: 3, config: 1 });
  assert.deepEqual(s.signatures, { V3: 1, agile: 1, ancienne: 1, aucune: 21 });
  assert.deepEqual(s.vb6, { projets: 1, lignes: inv.fichiers.find((f) => f.famille === 'vb6').lignesCode, formulaires: 1 });
  assert.deepEqual(s.courriel, { fichiers: 9, smtp: 3, outlook: 6 });
  const suivi = inv.fichiers.find((f) => f.chemin.endsWith('Suivi budgétaire 2024.xlsm'));
  assert.equal(suivi.meta.auteur, 'C. Martin');
  assert.equal(suivi.signature.certificats[0].sujet, 'Service Finance - Signature macros');
  assert.equal(inv.fichiers.find((f) => f.chemin.endsWith('Tableau de bord 2017.xls')).meta.auteur, 'J. Petit');
  const planning = inv.fichiers.find((f) => f.chemin.endsWith('Planning des congés.xlsm'));
  assert.deepEqual(planning.projets[0].modules.map((m) => m.type), ['document', 'document', 'formulaire', 'module', 'module']);
  const b64 = inv.fichiers.find((f) => f.chemin.endsWith('Calcul des primes.xlsm')).constats.filter((c) => c.regle === 'B64-001');
  assert.deepEqual(b64.map((c) => `${c.module}:${c.ligne}`), ['modApi:2', 'modApi:3']);
  await rm(d, { recursive: true });
});

test('fichier illisible, renommé ou vide : constat FMT-003, jamais d\'exception', () => {
  for (const buf of [Buffer.alloc(0), Buffer.from('pas un fichier Office'), Buffer.concat([Buffer.from('d0cf11e0a1b11ae1', 'hex'), Buffer.alloc(600, 0xff)])]) {
    const r = analyserFichier('x/y.xlsm', buf);
    assert.deepEqual(regles(r), ['FMT-003']);
  }
});

test('anonymisation pour le diagnostic : ni code, ni auteurs ; avec sansChemins, ni noms', async () => {
  const d = await tmp();
  await genererParc(path.join(d, 'Parc'));
  const inv = await inventorier([path.join(d, 'Parc')], { maintenant: DATE });
  const a = JSON.stringify(anonymiser(inv));
  assert.doesNotMatch(a, /CreateObject|Declare|C\. Martin|L\. Bernard|Service Finance/);
  assert.match(a, /Suivi budgétaire/);
  const b = JSON.stringify(anonymiser(inv, { sansChemins: true }));
  assert.doesNotMatch(b, /Suivi budgétaire|Finance|modApi|srv-fichiers|OutilsCommuns\.xlam/);
  assert.equal(synthese(anonymiser(inv, { sansChemins: true })).fichiersP1, synthese(inv).fichiersP1, 'les chiffres restent identiques');
  await rm(d, { recursive: true });
});

test('ligne de commande : rapports, Excel, diagnostic, reprise et comparaison', async () => {
  const d = await tmp();
  await genererParc(path.join(d, 'Parc'));
  const out = path.join(d, 'out');
  assert.equal(await principal([path.join(d, 'Parc'), '-o', out, '--date', '2026-10-23', '--diagnostic', '--sans-chemins', '-q'], { sortieErr: muet, sortieStd: muet, interactif: false }), 0);
  for (const f of ['rapport.html', 'inventaire.json', 'diagnostic.json', 'diagnostic-apercu.html', ...(PRO ? ['inventaire.xlsx', 'fichiers.csv', 'constats.csv'] : [])]) assert.ok(existsSync(path.join(out, f)), f);
  if (!PRO) assert.ok(!existsSync(path.join(out, 'inventaire.xlsx')));
  const corr = await readFile(path.join(out, 'diagnostic-correspondance.csv'), 'utf8');
  assert.match(corr, /dossier-\d{3}\/.*fichier-\d{3}\.xlsm;Parc\/Finance\/Budget\/Suivi budgétaire 2024\.xlsm/);
  assert.doesNotMatch(await readFile(path.join(out, 'diagnostic.json'), 'utf8'), /Suivi budgétaire/);
  const html = await readFile(path.join(out, 'rapport.html'), 'utf8');
  assert.doesNotMatch(html, /<(script|link|img)[^>]+(src|href)="https?:/i, 'aucune ressource externe');
  for (const s of ['Ce qui casse, et quand', 'Constats par règle', 'Signatures et références']) assert.ok(html.includes(s), s);
  for (const s of ['Estimation de migration', 'Copies d', 'Compatibilité LibreOffice', 'Par dossier']) assert.equal(html.includes(s), PRO, s);
  const apercu = await readFile(path.join(out, 'diagnostic-apercu.html'), 'utf8');
  assert.doesNotMatch(apercu, /Suivi budgétaire|Set re = |Declare Sub Sleep/);
  // Classeur : huit feuilles lisibles
  if (PRO) {
  const z = lireZip(await readFile(path.join(out, 'inventaire.xlsx')));
  assert.equal(z.noms.filter((n) => n.startsWith('xl/worksheets/')).length, 8);
  assert.match(z.texte('xl/workbook.xml'), /name="Constats"/);
  const csv = await readFile(path.join(out, 'constats.csv'), 'utf8');
  assert.ok(csv.startsWith('\ufeffPriorité;Règle;'));
  // Le fichier malveillant est isolé, jamais chiffré comme une correction
  assert.match(html, /1 fichier à isoler, hors estimation/);
  assert.match(await readFile(path.join(out, 'fichiers.csv'), 'utf8'), /Facture_2024_0172\.docm;.*à isoler \(SEC-010\)/);
  }
  // Reprise depuis inventaire.json avec revue, puis comparaison
  const revue = path.join(d, 'revue.md');
  await writeFile(revue, '## Synthèse\n\nTexte de synthèse **relu**.\n\n## Remarques par constat\n\n- B64-001 : à corriger en premier.\n\n## Plan d\'action\n\n1. Corriger les Declare.\n');
  const out2 = path.join(d, 'out2');
  assert.equal(await principal(['--depuis', path.join(out, 'inventaire.json'), '--revue', revue, '-o', out2, '-q'], { sortieErr: muet, sortieStd: muet, interactif: false }), 0);
  const h2 = await readFile(path.join(out2, 'rapport.html'), 'utf8');
  assert.match(h2, /Texte de synthèse <strong>relu<\/strong>/);
  assert.match(h2, /Avis de l'expert<\/b>à corriger en premier/);
  assert.match(h2, /Plan d'action/);
  if (!PRO) { await assert.rejects(principal([path.join(d, 'Parc'), '--comparer', path.join(out, 'inventaire.json')], { sortieErr: muet, sortieStd: muet, interactif: false }), /édition parc/); await rm(d, { recursive: true }); return; }
  await rm(path.join(d, 'Parc', 'Scripts'), { recursive: true });
  const out3 = path.join(d, 'out3');
  assert.equal(await principal([path.join(d, 'Parc'), '-o', out3, '--comparer', path.join(out, 'inventaire.json'), '--date', '2026-10-23', '-q'], { sortieErr: muet, sortieStd: muet, interactif: false }), 0);
  const h3 = await readFile(path.join(out3, 'rapport.html'), 'utf8');
  assert.match(h3, /0 fichiers sont apparus et 2 ont disparu/);
  await rm(d, { recursive: true });
});

test('édition gratuite : s\'arrête à la limite et liste le reste', async () => {
  const d = await tmp();
  await genererParc(path.join(d, 'Parc'));
  const inv = await inventorier([path.join(d, 'Parc')], { maintenant: DATE, limite: 5 });
  assert.equal(inv.fichiers.filter((f) => !['dotnet', 'config'].includes(f.famille)).length, 5);
  assert.equal(inv.nonAnalyses.length, 23, 'les sources .NET et configurations ne comptent pas dans la limite');
  await rm(d, { recursive: true });
});

test('erreurs d\'usage : messages clairs', async () => {
  await assert.rejects(principal(['--inconnue'], { sortieErr: muet, sortieStd: muet, interactif: false }), /option inconnue/);
  await assert.rejects(principal(['/nexiste/pas'], { sortieErr: muet, sortieStd: muet, interactif: false }), /introuvable/);
  assert.equal(await principal([], { sortieErr: muet, sortieStd: muet, interactif: false }), 1);
  assert.equal(Object.keys(REGLES).length, 39);
});

test('VB6 : projet lu depuis le .vbp, règles 64 bits écartées, VBScript en P1, composants classés', async () => {
  const d = await tmp();
  await genererParc(path.join(d, 'Parc'));
  const inv = await inventorier([path.join(d, 'Parc', 'Production', 'Applications')], { maintenant: DATE });
  assert.equal(inv.fichiers.length, 1, 'les .frm, .bas et .cls du projet ne sont pas inventoriés à part');
  const f = inv.fichiers[0];
  assert.equal(f.famille, 'vb6');
  const p = f.projets[0];
  assert.deepEqual(p.vb6, { type: 'Exe', nom: 'BancTest', exe: 'BancTest.exe', membres: 4, manquants: [] });
  assert.deepEqual(p.modules.map((m) => `${m.nom}:${m.type}`), ['frmBanc:formulaire', 'modMesures:module', 'modAlerte:module', 'clsJournal:classe']);
  assert.deepEqual(p.objets.map((o) => o.fichier), ['MSCOMM32.OCX', 'MSCOMCTL.OCX', 'THREED32.OCX']);
  const ids = f.constats.map((c) => c.regle);
  for (const x of ['B64-001', 'B64-003', 'B64-004', 'VBS-001']) assert.ok(!ids.includes(x), `${x} ne s'applique pas à VB6`);
  assert.equal(REGLES['VB6-010'].p, 1);
  assert.deepEqual(f.constats.filter((c) => c.regle === 'VB6-002').map((c) => c.extrait), ['THREED32.OCX'], 'MSCOMCTL et MSCOMM sont pris en charge par Microsoft');
  const out = f.constats.find((c) => c.regle === 'OUT-001');
  assert.equal(`${out.module}:${out.ligne}`, 'modAlerte:5', 'numéro de ligne de l\'éditeur VB6');
  assert.deepEqual(f.courriel, { mecanismes: ['OUT-001'], envoiDirect: 1, affichage: 0, piecesJointes: 0, lectureBoite: 0 });
  // Membre manquant : signalé, sans exception.
  await rm(path.join(d, 'Parc', 'Production', 'Applications', 'BancTest', 'clsJournal.cls'));
  const inv2 = await inventorier([path.join(d, 'Parc', 'Production', 'Applications')], { maintenant: DATE });
  assert.deepEqual(inv2.fichiers[0].projets[0].vb6.manquants, ['clsJournal.cls']);
  assert.match(inv2.fichiers[0].erreurs[0], /introuvable/);
  // Un .frm VB6 isolé n'est pas pris pour un UserForm VBA.
  const frm = analyserFichier('x/frmBanc.frm', await readFile(path.join(d, 'Parc', 'Production', 'Applications', 'BancTest', 'frmBanc.frm')));
  assert.equal(frm.famille, 'vb6');
  await rm(d, { recursive: true });
});

test('courriel : Outlook, SendObject, MAPI, SMTP par mot de passe et CDO', () => {
  const mod = (code) => regles(analyserFichier('x/m.bas', Buffer.from(`Attribute VB_Name = "m"\r\n${code.replace(/\n/g, '\r\n')}`)));
  assert.deepEqual(mod('DoCmd.SendObject acSendReport, "r", acFormatPDF, "a@b.c", , , "s", "t", True\n'), ['OUT-002']);
  assert.deepEqual(mod('Set s = CreateObject("Redemption.RDOSession")\n'), ['OUT-003']);
  assert.deepEqual(mod('Set s = CreateObject("MAPI.Session")\n'), ['OUT-003']);
  const cdo = 'Set m = CreateObject("CDO.Message")\nc.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"\n';
  assert.deepEqual(mod(cdo + 'c.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "x"\n'), ['MEL-001'], 'Exchange Online + mot de passe');
  assert.deepEqual(mod(cdo), ['MEL-002'], 'sans mot de passe : information');
  assert.deepEqual(mod('Set m = CreateObject("CDO.Message")\nc.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.interne.local"\n'), ['MEL-002']);
  // Usage : affichage avant envoi, et l'envoi d'une requête HTTP n'est pas un courriel.
  const f = analyserFichier('x/m.bas', Buffer.from('Set o = CreateObject("Outlook.Application")\r\nSet m = o.CreateItem(0)\r\nm.Display\r\nSet h = CreateObject(\"MSXML2.XMLHTTP\")\r\nh.Send\r\nSet n = o.GetNamespace("MAPI")\r\n'));
  assert.deepEqual(f.courriel, { mecanismes: ['OUT-001'], envoiDirect: 0, affichage: 1, piecesJointes: 0, lectureBoite: 1 });
  // Configurations : mot de passe SMTP vers Exchange Online seulement.
  const cfg = (t) => regles(analyserFichier('x/app.config', Buffer.from(t)));
  assert.deepEqual(cfg('<add key="Serveur" value="smtp.office365.com" />\n<add key="Password" value="x" />'), ['MEL-001']);
  assert.deepEqual(cfg('<add key="Serveur" value="smtp.office365.com" />'), []);
  assert.deepEqual(cfg('[smtp]\nserveur=mail.interne.local\npassword=x\n'), []);
});

const PWSH = ['pwsh', '/opt/pwsh/pwsh'].find((p) => !spawnSync(p, ['-NoProfile', '-c', '1']).error);
test('export Access (PowerShell, faux Access) puis inventaire du code exporté', { skip: PWSH ? false : 'pwsh absent' }, async () => {
  const r = spawnSync(PWSH, ['-NoProfile', '-File', path.join(ICI, 'access.test.ps1')], { encoding: 'utf8' });
  assert.match(r.stdout, /OK\s*$/, r.stdout + r.stderr);
  const sortie = /SORTIE=(.*)/.exec(r.stdout)[1].trim();
  const inv = await inventorier([sortie], { maintenant: DATE });
  const par = Object.fromEntries(inv.fichiers.map((f) => [f.chemin.split('/').slice(1).join('/'), regles(f)]));
  assert.deepEqual(par['Base NC.accdb/modExport.bas'], ['B64-001']);
  assert.deepEqual(par['Base NC.accdb/Form_Saisie.cls'], ['OUT-001']);
  await rm(path.dirname(sortie), { recursive: true });
});
