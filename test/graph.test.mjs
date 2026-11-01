// ModGraphMail contre un faux Microsoft Graph : le module VBA est transcrit en VBScript (mêmes instructions,
// sans les types) et exécuté par cscript sous Wine. Seuls la lecture binaire des fichiers et le coffre Windows
// (bloc « binaire ») sont remplacés par des équivalents de test.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn, spawnSync } from 'node:child_process';
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyserFichier } from '../src/scan.mjs';
import { decoderTexte } from '../src/format/office.mjs';

const ICI = path.dirname(fileURLToPath(import.meta.url));
const MODULE = path.join(ICI, '..', 'remediation', 'ModGraphMail.bas');
const WINE = ['/usr/lib/wine/wine64', 'wine64', 'wine'].find((p) => !spawnSync(p, ['--version']).error);

import { versVbscript, STUBS, UNICODE, oem, fauxGraph } from './outils/graph-vbscript.mjs';

const lancer = (args, env) => new Promise((ok) => {
  const p = spawn(WINE, args, { env: { ...process.env, ...env } });
  let sortie = '';
  p.stdout.on('data', (d) => { sortie += oem(d); });
  p.stderr.on('data', () => {});
  p.on('close', (code) => ok({ code, sortie }));
});

test('ModGraphMail.bas est au format d\'un export VBA : Windows-1252 et fins de ligne CRLF', async () => {
  const buf = await readFile(MODULE);
  assert.throws(() => new TextDecoder('utf-8', { fatal: true }).decode(buf), 'pas en UTF-8 : l\'éditeur VBA importe en ANSI');
  assert.equal(buf.toString('latin1').split('\n').length - 1, buf.toString('latin1').split('\r\n').length - 1, 'toutes les lignes en CRLF');
  assert.match(decoderTexte(buf), /désactivé par défaut fin décembre 2026/);
});

test('transcription : le VBScript produit garde toute la logique du module', async () => {
  const vbs = versVbscript(decoderTexte(await readFile(MODULE)));
  assert.doesNotMatch(vbs, /\bAs (String|Long|Double|Boolean|Variant|Object|Date|Integer)\b/);
  assert.doesNotMatch(vbs, /Optional|Declare|#If|CredReadW/);
  for (const f of ['EnvoyerCourriel', 'CreerBrouillon', 'Televerser', 'Jeton', 'Echapper', 'ValeurJson', 'EncoderUrl', 'Echec']) assert.match(vbs, new RegExp(`Function ${f}\\(|Sub ${f}\\(`), f);
});

test('le module de remédiation passe son propre inventaire : aucune règle P1 ou P2', async () => {
  const r = analyserFichier('remediation/ModGraphMail.bas', await readFile(MODULE));
  const ids = [...new Set(r.constats.map((c) => c.regle))].sort();
  assert.deepEqual(ids, ['SEC-001', 'SEC-002'], 'appel réseau et ouverture du navigateur, documentés ; Declare protégés par #If VBA7');
  assert.equal(r.courriel, undefined, 'il n\'utilise ni Outlook, ni MAPI, ni SMTP');
});

test('ModGraphMail sous VBScript (Wine) contre un faux Graph', { skip: WINE ? false : 'wine absent', timeout: 240000 }, async () => {
  const d = await mkdtemp(path.join(tmpdir(), 'graphmail-'));
  const g = await fauxGraph();
  try {
    const petit = 'Facture 2024-118\r\nMontant : 1 250,00 EUR\r\n'.repeat(20);
    const moyen = 'ligne de donnees;12;34\r\n'.repeat(150); // 3,6 Ko : pièce directe dans le brouillon
    const gros = Array.from({ length: 3400000 }, (_, i) => String.fromCharCode(65 + (i % 26))).join(''); // 3,4 Mo : téléversé
    await writeFile(path.join(d, 'facture.txt'), petit);
    await writeFile(path.join(d, 'vide.txt'), '');
    await writeFile(path.join(d, 'moyen.csv'), moyen);
    await writeFile(path.join(d, 'gros.dat'), gros);
    const w = (f) => `Z:${path.join(d, f).replace(/\//g, '\\')}`;
    const pilote = `Option Explicit
Const GRAPH_LOCATAIRE = "exemple.onmicrosoft.com"
Const GRAPH_EXPEDITEUR = "compta@exemple.invalid"
Const GRAPH_COFFRE = "GraphMail"
${versVbscript(decoderTexte(await readFile(MODULE))).replace(/^Option Explicit\r?$/m, '').replace(/^Const GRAPH_(LOCATAIRE|EXPEDITEUR|COFFRE) = .*\r?$/gm, '')}
${STUBS}
${UNICODE}
Sub Etape(nom)
    If Err.Number <> 0 Then WScript.Echo "ERREUR " & nom & " : " & Err.Description Else WScript.Echo "OK " & nom
    Err.Clear
End Sub
On Error Resume Next
GraphParametresAvances "http://127.0.0.1:${g.port}/connexion/", "http://127.0.0.1:${g.port}/v1.0/", "MSXML2.ServerXMLHTTP.6.0"
GraphIdentifiants "11111111-2222-3333-4444-555555555555", "faux"
EnvoyerCourriel "a@exemple.invalid", "x", "y", "", "", "", ""
Etape "secret refusé"
GraphIdentifiants "11111111-2222-3333-4444-555555555555", "s3cr~t.x_y+/="
EnvoyerCourriel "client@exemple.invalid; Service achats <achats@exemple.invalid>", "Relance n" & TChrW(176) & " 12 " & TChrW(8212) & " " & TChrW(233) & "t" & TChrW(233) & " ""urgent""", "<p>Bonjour,</p>" & vbCrLf & "<p>Ci-joint.</p>", "${w('facture.txt')}|${w('vide.txt')}", "copie@exemple.invalid", "", ""
Etape "envoi direct"
EnvoyerCourriel "client@exemple.invalid", "Texte", "Ligne 1" & vbCrLf & "Ligne 2 avec \\ et tabulation" & vbTab & "fin", "", "", "archives@exemple.invalid", "autre@exemple.invalid"
Etape "texte et expéditeur"
EnvoyerCourriel "client@exemple.invalid", "Gros", "Voir pièces.", "${w('moyen.csv')}|${w('gros.dat')}|${w('facture.txt')}", "", "", ""
Etape "brouillon et téléversement"
WScript.Echo "LIEN " & CreerBrouillon("relecture@exemple.invalid", "A relire", "Corps", "", "", "", "")
Etape "brouillon"
EnvoyerCourriel "a@exemple.invalid", "x", "y", "", "", "", "interdit@exemple.invalid"
Etape "droit refusé"
EnvoyerCourriel "a@exemple.invalid", "x", "y", "${w('absent.pdf')}", "", "", ""
Etape "pièce absente"
EnvoyerCourriel " ; ", "x", "y", "", "", "", ""
Etape "sans destinataire"
GraphParametresAvances "http://127.0.0.1:1/connexion/", "http://127.0.0.1:1/v1.0/", "MSXML2.ServerXMLHTTP.6.0"
EnvoyerCourriel "a@exemple.invalid", "x", "y", "", "", "", ""
Etape "réseau coupé"
`;
    await writeFile(path.join(d, 'pilote.vbs'), pilote.replace(/\n/g, '\r\n'), 'latin1');
    if (process.env.GRAPHMAIL_TRACE) await writeFile(process.env.GRAPHMAIL_TRACE, pilote.replace(/\n/g, '\r\n'), 'latin1');
    const r = await lancer(['cscript', '//nologo', w('pilote.vbs')], { WINEDEBUG: '-all', LANG: 'C.UTF-8' });
    const lignes = r.sortie.split(/\r?\n/).filter(Boolean);
    const etape = (nom) => lignes.find((l) => l.endsWith(` ${nom}`) || l.includes(` ${nom} : `)) ?? `absente : ${nom}\n${r.sortie}`;

    assert.match(etape('secret refusé'), /^ERREUR .*HTTP 401, invalid_client\) : AADSTS7000215: Invalid client secret provided\. Le secret est faux ou expiré/);
    assert.equal(etape('envoi direct'), 'OK envoi direct');
    assert.equal(etape('texte et expéditeur'), 'OK texte et expéditeur');
    assert.equal(etape('brouillon et téléversement'), 'OK brouillon et téléversement');
    assert.equal(etape('brouillon'), 'OK brouillon');
    assert.match(lignes.find((l) => l.startsWith('LIEN ')), /^LIEN https:\/\/outlook\.office365\.com\/owa\/\?ItemID=AAMk\d&exvsurl=1&viewmodel=ReadMessageItem$/);
    assert.match(etape('droit refusé'), /HTTP 403, ErrorAccessDenied\) : Access is denied.*interdit@exemple\.invalid.*Test-ServicePrincipalAuthorization/);
    assert.match(etape('pièce absente'), /Pièce jointe introuvable : Z:.*absent\.pdf/);
    assert.match(etape('sans destinataire'), /Aucun destinataire/);
    assert.match(etape('réseau coupé'), /Connexion impossible à 127\.0\.0\.1:1 .*proxy/);

    // Jeton : demandé une fois (le refus initial mis à part), formulaire encodé correctement.
    const jetons = g.journal.filter((x) => x.chemin.endsWith('/token'));
    assert.equal(jetons.length, 2);
    const f = new URLSearchParams(jetons[1].corps.toString());
    assert.deepEqual(Object.fromEntries(f), { client_id: '11111111-2222-3333-4444-555555555555', client_secret: 's3cr~t.x_y+/=', scope: `http://127.0.0.1:${g.port}/.default`, grant_type: 'client_credentials' });
    assert.equal(jetons[1].entetes['content-type'], 'application/x-www-form-urlencoded');

    // Envoi direct : JSON valide, accents et guillemets intacts, HTML, destinataires nommés, pièces en Base64.
    const envois = g.journal.filter((x) => x.chemin.endsWith('/sendMail'));
    const m1 = JSON.parse(envois[0].corps.toString());
    assert.equal(envois[0].chemin, '/v1.0/users/compta@exemple.invalid/sendMail');
    assert.ok(/^[\x20-\x7e]*$/.test(envois[0].corps.toString()), 'corps en ASCII pur');
    assert.equal(m1.message.subject, 'Relance n° 12 — été "urgent"');
    assert.deepEqual(m1.message.body, { contentType: 'HTML', content: '<p>Bonjour,</p>\r\n<p>Ci-joint.</p>' });
    assert.deepEqual(m1.message.toRecipients, [{ emailAddress: { address: 'client@exemple.invalid' } }, { emailAddress: { address: 'achats@exemple.invalid', name: 'Service achats' } }]);
    assert.deepEqual(m1.message.ccRecipients, [{ emailAddress: { address: 'copie@exemple.invalid' } }]);
    assert.equal(m1.saveToSentItems, true);
    assert.deepEqual(m1.message.attachments.map((a) => [a['@odata.type'], a.name, Buffer.from(a.contentBytes, 'base64').toString()]), [['#microsoft.graph.fileAttachment', 'facture.txt', petit], ['#microsoft.graph.fileAttachment', 'vide.txt', '']]);
    const m2 = JSON.parse(envois[1].corps.toString());
    assert.equal(envois[1].chemin, '/v1.0/users/autre@exemple.invalid/sendMail');
    assert.deepEqual(m2.message.body, { contentType: 'Text', content: 'Ligne 1\r\nLigne 2 avec \\ et tabulation\tfin' });
    assert.deepEqual(m2.message.bccRecipients, [{ emailAddress: { address: 'archives@exemple.invalid' } }]);
    assert.equal(m2.message.attachments, undefined);

    // Plus de 2,5 Mo : brouillon, petite pièce directe, 3,4 Mo téléversés par blocs de 960 Kio, puis envoi.
    const brouillon = g.journal.findIndex((x) => x.chemin === '/v1.0/users/compta@exemple.invalid/messages');
    const suite = g.journal.slice(brouillon + 1).filter((x) => !x.chemin.endsWith('/token'));
    assert.equal(suite[0].chemin, '/v1.0/users/compta@exemple.invalid/messages/AAMk1+/=/attachments');
    assert.equal(Buffer.from(JSON.parse(suite[0].corps).contentBytes, 'base64').toString(), moyen);
    assert.match(suite[1].chemin, /\/messages\/AAMk1\+\/=\/attachments\/createUploadSession$/);
    assert.deepEqual(JSON.parse(suite[1].corps), { AttachmentItem: { attachmentType: 'file', name: 'gros.dat', size: 3400000 } });
    const blocs = [...g.televersements.values()][0].recu;
    assert.deepEqual(blocs.map((b) => b.plage), ['bytes 0-983039/3400000', 'bytes 983040-1966079/3400000', 'bytes 1966080-2949119/3400000', 'bytes 2949120-3399999/3400000']);
    assert.equal(Buffer.concat(blocs.map((b) => b.corps)).toString(), gros);
    const puts = g.journal.filter((x) => x.methode === 'PUT');
    assert.ok(puts.every((x) => !x.entetes.authorization), 'les blocs partent sans jeton');
    assert.ok(puts.every((x) => x.chemin.startsWith('/televersement/')), 'adresse de téléversement reprise telle quelle');
    const apres = suite.slice(2).filter((x) => x.methode !== 'PUT');
    assert.equal(apres[0].chemin, '/v1.0/users/compta@exemple.invalid/messages/AAMk1+/=/attachments', 'la petite facture suit le gros fichier');
    assert.equal(apres[1].chemin, '/v1.0/users/compta@exemple.invalid/messages/AAMk1+/=/send');
    // Le brouillon seul n'est pas envoyé.
    assert.ok(!g.journal.some((x) => x.chemin.includes('AAMk2') && x.chemin.endsWith('/send')));
  } finally {
    g.serveur.close();
    await rm(d, { recursive: true, force: true });
  }
});
