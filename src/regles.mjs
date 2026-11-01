// Catalogue des règles et détection dans le code VBA.
// Priorités (comme une gestion d'alarmes) : P1 casse à une échéance connue, P2 à traiter avant 2027,
// P3 à planifier, P4 information. Chaque règle a un identifiant stable pour comparer deux inventaires.

const MS_VBSCRIPT = 'https://devblogs.microsoft.com/microsoft365dev/how-to-prepare-vba-projects-for-vbscript-deprecation/';
const MS_V3 = 'https://devblogs.microsoft.com/microsoft365dev/upgrade-signed-office-vba-macro-projects-to-v3-signature/';
const MS_64 = 'https://learn.microsoft.com/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview';
const OUTLOOK = 'https://www.directionsonmicrosoft.com/microsoft-delays-by-a-year-its-new-outlook-rollout/';
const ACTIVEX = 'https://www.techradar.com/pro/activex-is-now-being-blocked-by-default-in-microsoft-365';
const OUTLOOK_FONCTIONS = 'https://support.microsoft.com/help/5038254';
const SMTP_AUTH = 'https://mc.merill.net/message/MC786329';
const VB6 = 'https://learn.microsoft.com/previous-versions/visualstudio/visual-basic-6/visual-basic-6-support-policy';

export const REGLES = {
  'B64-001': { p: 1, axe: '64 bits', titre: 'Declare sans PtrSafe', pourquoi: 'Ne compile pas dans Office 64 bits, installé par défaut depuis Office 2019 : le classeur affiche une erreur de compilation à l\'ouverture.', quoi: 'Ajouter PtrSafe, passer les handles et pointeurs en LongPtr, ou encadrer par #If VBA7 Then … #Else … #End If.', effort: 0.25, source: MS_64 },
  'B64-002': { p: 2, axe: '64 bits', titre: 'Handle ou pointeur déclaré en Long', pourquoi: 'Le Declare compile en 64 bits, mais un handle ou un pointeur tronqué à 32 bits provoque des plantages aléatoires.', quoi: 'Typer en LongPtr les paramètres et retours qui portent un handle (h…, hWnd) ou un pointeur (lp…).', effort: 0.25, source: MS_64 },
  'B64-003': { p: 1, axe: '64 bits', titre: 'Contrôle ou bibliothèque 32 bits uniquement', pourquoi: 'MSCOMCTL (ListView, TreeView), MSCOMCT2 (DTPicker), Calendar, Script Control et leurs semblables n\'existent pas en 64 bits : le formulaire ou la référence est « manquant ».', quoi: 'Remplacer par un contrôle MSForms, une saisie de date maison ou une bibliothèque 64 bits ; retirer la référence.', effort: 4, source: MS_64 },
  'B64-004': { p: 1, axe: '64 bits', titre: 'Jet 4.0, DAO 3.6 ou pilote ODBC 32 bits', pourquoi: 'Microsoft.Jet.OLEDB.4.0, DAO 3.6 et les pilotes ODBC « *.mdb » / « *.xls » n\'existent qu\'en 32 bits.', quoi: 'Passer à Microsoft.ACE.OLEDB.12.0 et à la référence « Microsoft Office Access database engine Object Library ».', effort: 1, source: MS_64 },
  'SIG-001': { p: 1, axe: 'Signatures', titre: 'Signé sans signature V3', pourquoi: 'Dès que la stratégie « Only trust VBA macros that use V3 signatures » est activée, la signature est ignorée et les macros sont bloquées.', quoi: 'Re-signer le projet avec un Office à jour (Outils > Signature numérique), ou en lot avec SignTool et Office SIPs.', effort: 0.25, source: MS_V3 },
  'SIG-002': { p: 2, axe: 'Signatures', titre: 'Certificat de signature expiré', pourquoi: 'Une signature dont le certificat a expiré n\'est plus considérée comme approuvée si elle n\'est pas horodatée.', quoi: 'Renouveler le certificat de signature de code et re-signer (en V3).', effort: 0.25, source: MS_V3 },
  'SIG-003': { p: 3, axe: 'Signatures', titre: 'Certificat qui expire dans moins de 180 jours', pourquoi: 'Le renouvellement doit précéder l\'expiration pour éviter un blocage.', quoi: 'Planifier le renouvellement et la re-signature.', effort: 0.1, source: MS_V3 },
  'VBS-001': { p: 2, axe: 'VBScript', titre: 'VBScript.RegExp', pourquoi: 'Dépend de vbscript.dll. Sur Office LTSC, 2021, 2019 et 2016, le code cassera quand Windows désactivera VBScript (phase 2, « vers 2026 ou 2027 »). Sans effet sur Microsoft 365 2508 et suivants.', quoi: 'Sur Microsoft 365 2508+ : rien d\'urgent, retirer la référence à « Microsoft VBScript Regular Expressions ». Ailleurs : prévoir un moteur sans VBScript ou la mise à jour d\'Office.', effort: 1, source: MS_VBSCRIPT },
  'VBS-002': { p: 1, axe: 'VBScript', titre: 'Lance un script .vbs', pourquoi: 'Les scripts .vbs ne s\'exécutent plus quand VBScript est désactivé, quelle que soit la version d\'Office.', quoi: 'Porter le script en VBA ou en PowerShell.', effort: 2, source: MS_VBSCRIPT },
  'VBS-003': { p: 1, axe: 'VBScript', titre: 'Script Control (MSScriptControl)', pourquoi: 'Utilise le moteur VBScript ou JScript et n\'existe qu\'en 32 bits.', quoi: 'Réécrire l\'évaluation en VBA, ou passer par PowerShell.', effort: 3, source: MS_VBSCRIPT },
  'VBS-004': { p: 2, axe: 'Parc mixte', titre: 'RegExp natif (Office 2508 et plus)', pourquoi: 'New RegExp sans référence ne compile que sur Microsoft 365 2508 et suivants : sur un poste Office LTSC, 2021 ou 2019, le module est en erreur.', quoi: 'Vérifier que tous les utilisateurs du fichier sont sur Microsoft 365 2508+, sinon garder la liaison tardive CreateObject("VBScript.RegExp") en attendant.', effort: 0.5, source: MS_VBSCRIPT },
  'VBS-010': { p: 1, axe: 'VBScript', titre: 'Script VBScript autonome (.vbs)', pourquoi: 'Ne s\'exécutera plus quand Windows désactivera VBScript (phase 2). Scripts de connexion et tâches planifiées compris.', quoi: 'Porter en PowerShell ; lister les tâches planifiées et GPO qui l\'appellent.', effort: 2, source: MS_VBSCRIPT },
  'OUT-001': { p: 2, axe: 'Courriel', titre: 'Automatisation d\'Outlook', pourquoi: 'Le nouvel Outlook n\'a ni modèle objet COM ni VBA. Phase « opt-out » en entreprise : mars 2027 ; Outlook classique pris en charge jusqu\'en 2029 au moins.', quoi: 'Garder Outlook classique sur les postes concernés, ou passer l\'envoi par Microsoft Graph (module ModGraphMail fourni dans remediation\\). Éviter le SMTP par mot de passe : Exchange Online le désactive fin 2026.', effort: 2, source: OUTLOOK },
  'OUT-002': { p: 2, axe: 'Courriel', titre: 'DoCmd.SendObject (Access)', pourquoi: 'SendObject confie le message au client de messagerie par défaut, par MAPI. Le nouvel Outlook ne fournit pas MAPI : l\'envoi échoue ou n\'ouvre rien.', quoi: 'Exporter l\'objet (DoCmd.OutputTo) puis l\'envoyer en pièce jointe par Microsoft Graph (module ModGraphMail).', effort: 1, source: OUTLOOK_FONCTIONS },
  'OUT-003': { p: 2, axe: 'Courriel', titre: 'MAPI, CDO 1.21 ou Redemption', pourquoi: 'Ces bibliothèques s\'appuient sur MAPI, que le nouvel Outlook ne fournit pas : rien à piloter dès que le poste bascule.', quoi: 'Remplacer l\'envoi par Microsoft Graph (module ModGraphMail) ; la lecture de boîte passe par Graph avec l\'autorisation Mail.Read.', effort: 3, source: OUTLOOK_FONCTIONS },
  'MEL-001': { p: 1, axe: 'Courriel', titre: 'SMTP par mot de passe vers Exchange Online', pourquoi: 'Exchange Online désactive par défaut l\'authentification de base SMTP AUTH fin décembre 2026 ; les locataires créés ensuite n\'y ont plus accès. L\'envoi échoue alors avec une erreur d\'authentification. C\'est aussi le contournement le plus souvent conseillé pour sortir d\'Outlook.', quoi: 'Passer l\'envoi par Microsoft Graph (module ModGraphMail, accès limité à une seule boîte). En attendant, l\'administrateur peut réactiver SMTP AUTH jusqu\'au retrait définitif, dont la date sera annoncée au second semestre 2027. Retirer le mot de passe du code ou du fichier.', effort: 1, source: SMTP_AUTH },
  'MEL-002': { p: 4, axe: 'Courriel', titre: 'Envoi SMTP par CDO', pourquoi: 'CDO.Message envoie sans Outlook, vers un serveur SMTP. Il fonctionne avec le nouvel Outlook, mais CDO ne sait pas s\'authentifier en OAuth : à vérifier si le serveur est Exchange Online.', quoi: 'Information : point d\'envoi à garder dans l\'inventaire des courriels ; vérifier le serveur et le mode d\'authentification.', effort: 0, source: '' },
  'VB6-001': { p: 4, axe: 'VB6', titre: 'Application Visual Basic 6', pourquoi: 'Le moteur d\'exécution VB6 reste pris en charge sur Windows 11 et Windows Server 2025, en 32 bits, sans date de fin. L\'environnement de développement ne l\'est plus depuis 2008 : chaque correction dépend d\'un poste et d\'une personne qui savent encore compiler.', quoi: 'Archiver les sources et le poste de compilation ; planifier la migration (C# .NET ou application web) avant le départ de la personne qui maintient l\'application.', effort: 0, source: VB6 },
  'VB6-002': { p: 3, axe: 'VB6', titre: 'Composant hors du support Microsoft', pourquoi: 'Contrôle ou bibliothèque absent de la liste « It Just Works » de Microsoft (threed32.ocx, graph32.ocx, contrôles tiers…) : il peut cesser de s\'installer à une mise à jour de Windows et n\'a pas d\'équivalent direct en .NET.', quoi: 'Garder l\'installateur d\'origine et sa licence ; le remplacer par un contrôle pris en charge, ou prévoir son équivalent dans la migration.', effort: 4, source: VB6 },
  'VB6-004': { p: 4, axe: 'VB6', titre: 'Communication avec du matériel', pourquoi: 'Port série (MSComm), sockets (Winsock) ou entrées-sorties directes : ce sont les parties les plus longues à recetter lors d\'une migration, car elles se testent avec le matériel.', quoi: 'Information : prévoir un banc de test ou des journaux d\'échanges enregistrés avant de migrer.', effort: 0, source: '' },
  'VB6-010': { p: 1, axe: 'VBScript', titre: 'VBScript dans une application VB6', pourquoi: 'VBScript.RegExp ou le Script Control dépendent de vbscript.dll : l\'application cassera quand Windows désactivera VBScript (phase 2, « vers 2026 ou 2027 »). Contrairement à Office, VB6 n\'a pas de RegExp natif de remplacement.', quoi: 'Remplacer les expressions régulières par du code VB6 ou une bibliothèque sans VBScript ; supprimer le Script Control.', effort: 3, source: MS_VBSCRIPT },
  'ACX-001': { p: 2, axe: 'ActiveX', titre: 'Contrôle ActiveX dans le document', pourquoi: 'Les contrôles ActiveX sont désactivés par défaut dans Microsoft 365 (version 2504, avril 2025) et Office 2024 : boutons et listes ActiveX posés sur les feuilles ne répondent plus.', quoi: 'Remplacer par des contrôles de formulaire ou des formes avec macro affectée, ou réactiver par stratégie en connaissance de cause.', effort: 1, source: ACTIVEX },
  'IE-001': { p: 2, axe: 'Dépendances retirées', titre: 'Internet Explorer piloté par macro', pourquoi: 'Internet Explorer est retiré depuis le 15 juin 2022 ; son pilotage COM n\'est plus pris en charge et peut cesser à toute mise à jour.', quoi: 'Passer par une API (MSXML2.XMLHTTP, Power Automate) ou par Edge avec WebDriver.', effort: 4, source: 'https://learn.microsoft.com/lifecycle/announcements/internet-explorer-11-support-end-dates' },
  'DDE-001': { p: 2, axe: 'Dépendances retirées', titre: 'Échange dynamique de données (DDE)', pourquoi: 'DDE est désactivé ou restreint par défaut dans les versions récentes d\'Office pour des raisons de sécurité.', quoi: 'Remplacer par l\'automatisation (CreateObject) ou par des liaisons de fichiers.', effort: 2, source: 'https://msrc.microsoft.com/update-guide/vulnerability/ADV170021' },
  'XLM-001': { p: 2, axe: 'Dépendances retirées', titre: 'Macros Excel 4.0 (XLM)', pourquoi: 'Les feuilles de macros XLM sont bloquées par défaut quand les macros VBA le sont (depuis 2021-2022).', quoi: 'Réécrire en VBA ou en formules.', effort: 4, source: 'https://techcommunity.microsoft.com/blog/excelblog/excel-4-0-xlm-macros-now-restricted-by-default-for-customer-protection/2541209' },
  'XLM-002': { p: 3, axe: 'Dépendances retirées', titre: 'ExecuteExcel4Macro', pourquoi: 'Appelle le moteur XLM depuis VBA ; même restriction que les feuilles XLM.', quoi: 'Remplacer par l\'appel VBA équivalent.', effort: 0.5, source: 'https://techcommunity.microsoft.com/blog/excelblog/excel-4-0-xlm-macros-now-restricted-by-default-for-customer-protection/2541209' },
  'CHM-001': { p: 3, axe: 'Portabilité', titre: 'Chemin de fichier écrit en dur', pourquoi: 'Lecteur réseau mappé, partage UNC ou dossier d\'un utilisateur : la macro casse au moindre déménagement de serveur, migration OneDrive ou changement de poste.', quoi: 'Centraliser les chemins dans une feuille de paramètres ou les calculer (ThisWorkbook.Path, Environ).', effort: 0.25, source: '' },
  'CNX-001': { p: 3, axe: 'Portabilité', titre: 'Connexion à une base de données', pourquoi: 'Dépend d\'un pilote ODBC / OLE DB et d\'un serveur : à vérifier à chaque changement de poste, de version de pilote ou de serveur.', quoi: 'Inventorier serveurs et pilotes ; éviter les mots de passe dans le code.', effort: 1, source: '' },
  'REF-001': { p: 3, axe: 'Portabilité', titre: 'Référence à un fichier hors du système', pourquoi: 'Le projet référence un autre classeur, un complément ou une bibliothèque sur un chemin réseau ou utilisateur : « référence manquante » dès qu\'il bouge.', quoi: 'Installer la dépendance à un emplacement stable ou passer en liaison tardive.', effort: 0.5, source: '' },
  'SEC-001': { p: 3, axe: 'Sécurité', titre: 'Exécute des commandes système', pourquoi: 'Shell, WScript.Shell ou PowerShell : comportement surveillé par les antivirus et souvent bloqué par les règles de réduction de surface d\'attaque (ASR).', quoi: 'Documenter le besoin ; préférer les API Office ou un outil dédié.', effort: 0.5, source: 'https://learn.microsoft.com/defender-endpoint/attack-surface-reduction-rules-reference' },
  'SEC-002': { p: 3, axe: 'Sécurité', titre: 'Télécharge ou appelle le réseau', pourquoi: 'URLDownloadToFile, XMLHTTP, WinHttp : à justifier, et sensible aux proxys et aux règles ASR.', quoi: 'Documenter l\'URL et le besoin ; centraliser les appels.', effort: 0.5, source: 'https://learn.microsoft.com/defender-endpoint/attack-surface-reduction-rules-reference' },
  'SEC-003': { p: 4, axe: 'Sécurité', titre: 'Macro lancée à l\'ouverture', pourquoi: 'Workbook_Open, Auto_Open, Document_Open… : s\'exécute dès que le fichier est ouvert et que les macros sont autorisées.', quoi: 'Information : à connaître pour les tests de migration.', effort: 0, source: '' },
  'SEC-010': { p: 1, axe: 'Sécurité', titre: 'Motif de macro malveillante', pourquoi: 'Exécution à l\'ouverture, téléchargement et lancement de commande dans le même projet : c\'est le schéma des macros de diffusion de logiciels malveillants.', quoi: 'Isoler le fichier et le faire examiner par l\'équipe sécurité avant toute ouverture.', effort: 1, source: '' },
  'PRJ-001': { p: 4, axe: 'Projet', titre: 'Projet protégé par mot de passe', pourquoi: 'Le code reste lisible par l\'outil, mais il faudra le mot de passe pour le corriger dans l\'éditeur VBA.', quoi: 'Retrouver le mot de passe auprès du service propriétaire avant la migration.', effort: 0, source: '' },
  'FMT-001': { p: 4, axe: 'Format', titre: 'Format 97-2003 avec macros', pourquoi: 'Les formats .xls, .doc, .xla sont toujours lus, mais ils limitent la taille des feuilles et ne portent pas les signatures V3 de la même façon que les formats actuels.', quoi: 'Convertir en .xlsm, .docm ou .xlam lors de la correction.', effort: 0.25, source: '' },
  'FMT-002': { p: 3, axe: 'Format', titre: 'Fichier chiffré à l\'ouverture', pourquoi: 'Protégé par mot de passe : les macros ne peuvent pas être analysées.', quoi: 'Obtenir le mot de passe, ou analyser une copie déprotégée.', effort: 0, source: '' },
  'FMT-003': { p: 2, axe: 'Format', titre: 'Fichier illisible', pourquoi: 'Fichier corrompu, tronqué, ou format inattendu sous une extension Office.', quoi: 'Ouvrir et réenregistrer le fichier, ou l\'archiver.', effort: 0.25, source: '' },
  'FMT-004': { p: 4, axe: 'Format', titre: 'Présentation 97-2003 non analysée', pourquoi: 'PowerPoint 97-2003 range son projet VBA dans un format que l\'outil ne lit pas.', quoi: 'Convertir en .pptm puis relancer l\'inventaire.', effort: 0, source: '' },
  'ACC-001': { p: 3, axe: 'Access', titre: 'Base Access avec code VBA', pourquoi: 'Le code d\'une base Access ne se lit qu\'avec Access : il n\'est pas encore dans cet inventaire.', quoi: 'Lancer outils\\exporter-access.ps1 sur un poste avec Access, puis relancer l\'inventaire sur le dossier exporté.', effort: 0, source: '' },
};

export const PRIORITES = { 1: 'P1 casse', 2: 'P2 avant 2027', 3: 'P3 à planifier', 4: 'P4 information' };

// ---------- Références (par nom de fichier, plus robuste que le GUID seul) ----------
const REFS = [
  { re: /vbscript\.dll|3F4DACA7-160D-11D2-A8E9-00104B365C9F/i, regle: 'VBS-001', quoi: 'Microsoft VBScript Regular Expressions' },
  { re: /msscript\.ocx|0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC/i, regle: 'VBS-003', quoi: 'Microsoft Script Control' },
  { re: /mscomctl\.ocx|831FDD16-0C5C-11D2-A9FC-0000F8754DA1/i, regle: 'B64-003', quoi: 'Microsoft Windows Common Controls 6.0 (MSCOMCTL)' },
  { re: /mscomct2\.ocx|86CF1D34-0C5F-11D2-A9FC-0000F8754DA1/i, regle: 'B64-003', quoi: 'Microsoft Windows Common Controls-2 6.0 (DTPicker, MonthView)' },
  { re: /comctl32\.ocx|6B7E6392-850A-101B-AFC0-4210102A8DA7/i, regle: 'B64-003', quoi: 'Microsoft Windows Common Controls 5.0' },
  { re: /mscal\.ocx|8E27C92E-1264-101C-8A2F-040224009C02/i, regle: 'B64-003', quoi: 'Contrôle Calendrier (MSCAL)' },
  { re: /msflxgrd\.ocx|mshflxgd\.ocx|comdlg32\.ocx|mscomm32\.ocx|mswinsck\.ocx|richtx32\.ocx|msinet\.ocx|tabctl32\.ocx|mschrt20\.ocx|msmask32\.ocx/i, regle: 'B64-003', quoi: 'Contrôle OCX de Visual Basic 6 (32 bits)' },
  { re: /dao360\.dll|dao350\.dll|00025E01-0000-0000-C000-000000000046/i, regle: 'B64-004', quoi: 'Microsoft DAO 3.6' },
  { re: /msoutl\.olb|00062FFF-0000-0000-C000-000000000046/i, regle: 'OUT-001', quoi: 'Microsoft Outlook Object Library' },
  { re: /cdo\.dll|3FA7DEA7-6438-101B-ACC1-00AA00423326|redemption(64)?\.dll|2D5E2D34-BED5-4B9F-9793-A31E26E6806E/i, regle: 'OUT-003', quoi: 'CDO 1.21 ou Redemption (MAPI)' },
  { re: /shdocvw\.dll|ieframe\.dll|EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B/i, regle: 'IE-001', quoi: 'Microsoft Internet Controls' },
];
// Composants VB6 (déclaration de support de Microsoft) : livrés avec Windows ou « étendus », pris en charge.
const VB6_PRIS_EN_CHARGE = new Set(['comct232.ocx', 'comct332.ocx', 'comctl32.ocx', 'comdlg32.ocx', 'dbgrid32.ocx', 'dblist32.ocx', 'mci32.ocx', 'msadodc.ocx', 'mschrt20.ocx', 'mscomct2.ocx', 'mscomctl.ocx', 'mscomm32.ocx', 'msdatgrd.ocx', 'msdatlst.ocx', 'msdatrep.ocx', 'msflxgrd.ocx', 'mshflxgd.ocx', 'msinet.ocx', 'msmapi32.ocx', 'msmask32.ocx', 'msrdc20.ocx', 'picclp32.ocx', 'richtx32.ocx', 'sysinfo.ocx', 'tabctl32.ocx', 'hhctrl.ocx']);
const VB6_MATERIEL = /^(mscomm32|mswinsck)\.ocx$/i;
const SYSTEME = /^(?:[a-z]:\\(?:windows|program files(?: \(x86\))?)\\|%|$)/i;

// Contrôles ActiveX posés dans les documents : CLSID connus des contrôles 32 bits.
const CLSID_32 = new Set(['BDD1F04B-858B-11D1-B16A-00C0F0283628', 'C74190B6-8589-11D1-B16A-00C0F0283628', '2C247F23-8591-11D1-B16A-00C0F0283628', '35053A22-8589-11D1-B16A-00C0F0283628', '8E3867A3-8586-11D1-B16A-00C0F0283628', '66833FE6-8583-11D1-B16A-00C0F0283628', '1EFB6596-857C-11D1-B16A-00C0F0283628', 'F08DF954-8592-11D1-B16A-00C0F0283628', '20DD1B9E-87C4-11D1-8BE3-0000F8754DA1', '232E456A-87C3-11D1-8BE3-0000F8754DA1', '8E27C92B-1264-101C-8A2F-040224009C02', '8856F961-340A-11D0-A96B-00C04FD705A2']);

// ---------- Lecture du code ----------
// Découpe en lignes logiques (suites « _ »), retire commentaires et attributs, suit la compilation conditionnelle.
export function lignesLogiques(code) {
  const brutes = code.split(/\r\n|\r|\n/);
  const out = [];
  const pile = [];
  let i = 0;
  while (i < brutes.length) {
    const debut = i;
    let texte = brutes[i];
    while (/(^|\s)_\s*$/.test(texte) && i + 1 < brutes.length) texte = texte.replace(/_\s*$/, ' ') + brutes[++i].trim();
    i++;
    const { code: c, chaines } = separer(texte);
    const t = c.trim();
    if (/^Attribute\s+VB_/i.test(t)) { out.push({ ligne: debut + 1, type: 'attribut' }); continue; }
    const pre = /^#\s*(If|ElseIf|Else|End\s*If)\b(.*)$/i.exec(t);
    if (pre) {
      const mot = pre[1].toLowerCase().replace(/\s+/g, '');
      if (mot === 'if') pile.push({ cond: pre[2], branche: 'if' });
      else if (mot === 'elseif' && pile.length) Object.assign(pile.at(-1), { cond: pre[2], branche: 'elseif' });
      else if (mot === 'else' && pile.length) pile.at(-1).branche = 'else';
      else if (mot === 'endif') pile.pop();
      out.push({ ligne: debut + 1, type: 'directive' });
      continue;
    }
    if (!t) { out.push({ ligne: debut + 1, type: texte.trim() ? 'commentaire' : 'vide' }); continue; }
    out.push({ ligne: debut + 1, type: 'code', code: t, chaines, brut: texte.trim(), contexte: pile.map((f) => ({ ...f })) });
  }
  // Numéros de ligne de l'éditeur VBA, qui n'affiche pas les attributs d'en-tête.
  const entete = out.findIndex((l) => l.type !== 'attribut');
  const decalage = entete < 0 ? 0 : brutesAvant(brutes, entete);
  if (decalage) for (const l of out) l.ligne -= decalage;
  return out;
}

// Nombre de lignes physiques occupées par les `n` premiers attributs (un attribut tient sur une ligne).
const brutesAvant = (brutes, n) => n;

// Sépare le code des chaînes et retire le commentaire (apostrophe hors chaîne, ou Rem).
function separer(ligne) {
  let code = '';
  const chaines = [];
  let k = 0;
  while (k < ligne.length) {
    const ch = ligne[k];
    if (ch === '"') {
      let s = '';
      k++;
      while (k < ligne.length) {
        if (ligne[k] === '"' && ligne[k + 1] === '"') { s += '"'; k += 2; continue; }
        if (ligne[k] === '"') { k++; break; }
        s += ligne[k++];
      }
      chaines.push(s);
      code += `"${'§'.repeat(Math.min(s.length, 1))}"`;
      continue;
    }
    if (ch === "'") break;
    code += ch;
    k++;
  }
  if (/^\s*Rem(\s|$)/i.test(code)) return { code: '', chaines: [] };
  // « : Rem » en milieu de ligne
  code = code.replace(/:\s*Rem\s.*$/i, '');
  return { code, chaines };
}

// Le Declare sans PtrSafe ne gêne pas s'il est dans la branche « ancienne version » d'un #If VBA7 / Win64.
function protege64(contexte) {
  return contexte.some((f) => {
    if (!/\b(VBA7|Win64)\b/i.test(f.cond)) return false;
    const nie = /\bNot\s+(VBA7|Win64)\b/i.test(f.cond);
    return nie ? f.branche === 'if' || f.branche === 'elseif' : f.branche === 'else';
  });
}

const HANDLE = /\b(?:ByVal|ByRef)?\s*(h[A-Z]\w*|hwnd|hWnd\w*|lp[A-Z]\w*|p[A-Z]\w*Ptr)\s+As\s+Long\b(?!Ptr|Long)/;
const RETOUR_HANDLE = /\b(?:FindWindow\w*|GetDC|GetWindowDC|LoadLibrary\w*|GetModuleHandle\w*|CreateFile\w*|GetForegroundWindow|GetActiveWindow|GetDesktopWindow|GetParent|GetWindow|OpenProcess|GlobalAlloc|GlobalLock|SetWindowsHookEx\w*|ShellExecute\w*)\b[^)]*\)\s*As\s+Long\s*$/i;

const MOTIFS = [
  { regle: 'VBS-001', chaine: /^VBScript\.RegExp$/i },
  { regle: 'VBS-002', chaine: /(?:^|[\\/\s"])[wc]script(?:\.exe)?(?:\s|$)|\.vbs\b/i },
  { regle: 'VBS-003', chaine: /^MSScriptControl\.ScriptControl/i },
  { regle: 'OUT-001', chaine: /^Outlook\.Application/i },
  { regle: 'OUT-001', code: /\bAs\s+(?:New\s+)?Outlook\.\w+/i },
  { regle: 'OUT-002', code: /\bSendObject\b/i },
  { regle: 'OUT-003', chaine: /^(MAPI\.Session|Redemption\.\w+|MSMAPI\.MAPI(Session|Messages)(\.1)?)$/i },
  { regle: 'OUT-003', code: /\b(MAPISendMail|MAPILogon|MAPISendDocuments)\b|\bAs\s+(?:New\s+)?(MAPI\.Session|Redemption\.\w+)/i },
  { regle: 'VB6-004', code: /\bAs\s+(?:New\s+)?(MSCommLib\.MSComm|MSWinsockLib\.Winsock)\b|\bOpen\s+"COM\d+:?"/i },
  { regle: 'IE-001', chaine: /^InternetExplorer\.Application/i },
  { regle: 'IE-001', code: /\bAs\s+(?:New\s+)?(?:SHDocVw\.)?InternetExplorer(?:Medium)?\b/i },
  { regle: 'B64-003', chaine: /^(MSComctlLib\.|MSComCtl2\.|MSCAL\.Calendar|MSFlexGridLib\.|MSComDlg\.|MSWinsockLib\.|InetCtls\.)/i },
  { regle: 'B64-004', chaine: /Microsoft\.Jet\.OLEDB\.4\.0|^DAO\.DBEngine\.3[56]$|\{Microsoft Access Driver \(\*\.mdb\)\}|\{Microsoft Excel Driver \(\*\.xls\)\}/i },
  { regle: 'DDE-001', code: /\b(DDEInitiate|DDEExecute|DDEPoke|DDERequest|DDETerminate(All)?)\b/i },
  { regle: 'XLM-002', code: /\bExecuteExcel4Macro\b/i },
  { regle: 'CNX-001', chaine: /(?:^|;)\s*(Provider|DSN|Driver|Data Source)\s*=|^ODBC;/i },
  { regle: 'CNX-001', chaine: /^ADODB\.Connection$/i },
  { regle: 'SEC-001', code: /(?:^|[\s:=(])Shell\s*(?:\(|"|\w)|\bShellExecute\w*\b/i },
  { regle: 'SEC-001', chaine: /^WScript\.Shell$|powershell(\.exe)?\b|cmd(\.exe)?\s+\/[ck]\b/i },
  { regle: 'SEC-002', code: /\bURLDownloadToFile\w*\b/i },
  { regle: 'SEC-002', chaine: /^(MSXML2\.(Server)?XMLHTTP(\.\d\.\d)?|Microsoft\.XMLHTTP|WinHttp\.WinHttpRequest\.5\.1)$/i },
];
const AUTO = /^(?:Private\s+|Public\s+)?Sub\s+(Workbook_Open|Auto_Open|AutoOpen|AutoExec|AutoNew|Document_Open|Document_New|Workbook_Activate|AutoClose|Auto_Close|Document_Close|Workbook_BeforeClose)\s*\(/i;
const CHEMIN = /(?:^|[\s=;(])((?:[A-Za-z]:\\|\\\\[\w.$-]+\\)\S*)/g;

/**
 * Analyse un module. Renvoie { lignes, lignesCode, lignesCommentaire, procedures, declares, constats: [{ regle, ligne, extrait }] }.
 */
// En VB6 (application 32 bits, pas d'Office) : pas de PtrSafe ni de 64 bits, et VBScript n'a pas d'alternative native.
const VB6_SANS = new Set(['B64-001', 'B64-002', 'B64-003', 'B64-004', 'VBS-004', 'SEC-003']);
const VB6_VERS = { 'VBS-001': 'VB6-010', 'VBS-003': 'VB6-010' };
const SMTP_O365 = /^(smtp\.office365\.com|smtp-mail\.outlook\.com|outlook\.office365\.com)$/i;
const SMTP_MDP = /sendpassword|smtpauthenticate/i;

export function analyserModule(module, { langage = 'vba' } = {}) {
  const lignes = lignesLogiques(module.code);
  const constats = [];
  const stats = { lignes: 0, lignesCode: 0, lignesCommentaire: 0, procedures: 0, declares: 0, regexNatif: false, auto: false, progIds: new Set(), courriel: { envoiDirect: 0, affichage: 0, piecesJointes: 0, lectureBoite: 0 } };
  stats.lignes = lignes.filter((l) => l.type !== 'attribut').length;
  const vb6 = langage === 'vb6';
  let smtp = null, cdo = null, motDePasse = false;
  const envoiNon = new Set();
  const ajouter = (regle, l) => {
    if (vb6 && VB6_SANS.has(regle)) return;
    constats.push({ regle: vb6 ? VB6_VERS[regle] ?? regle : regle, ligne: l.ligne, extrait: l.brut.slice(0, 160) });
  };
  for (const l of lignes) {
    if (l.type === 'commentaire') stats.lignesCommentaire++;
    if (l.type !== 'code') continue;
    stats.lignesCode++;
    const c = l.code;
    if (/^(?:(?:Public|Private|Friend|Static)\s+)*(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+\w/i.test(c)) stats.procedures++;
    if (AUTO.test(c)) { stats.auto = true; ajouter('SEC-003', l); }
    if (/^(?:(?:Public|Private)\s+)?Declare\s/i.test(c)) {
      stats.declares++;
      const ptrSafe = /^(?:(?:Public|Private)\s+)?Declare\s+PtrSafe\s/i.test(c);
      if (!ptrSafe && !protege64(l.contexte)) ajouter('B64-001', l);
      if (ptrSafe && (HANDLE.test(c) || RETOUR_HANDLE.test(c))) ajouter('B64-002', l);
    }
    if (/\b(?:CreateObject|GetObject)\s*\(\s*"§"/i.test(c)) for (const x of l.chaines) if (/^[A-Za-z][\w]*\.[\w.]+$/.test(x.trim())) { stats.progIds.add(x.trim()); break; }
    if (/\bNew\s+(?:VBA\.)?RegExp\b|\bAs\s+(?:VBA\.)?RegExp\b/i.test(c)) stats.regexNatif = l;
    for (const m of MOTIFS) {
      const okCode = !m.code || m.code.test(c);
      const okChaine = !m.chaine || l.chaines.some((s) => m.chaine.test(s.trim()));
      if (m.regle === 'SEC-001' && m.code && /\b(?:Dim|Private|Public|Set|As)\s+Shell\b/i.test(c)) continue;
      if (m.code && m.chaine ? okCode && okChaine : m.code ? okCode : okChaine) ajouter(m.regle, l);
    }
    for (const s of l.chaines) {
      if ([...s.matchAll(CHEMIN)].some((m) => !/^[A-Za-z]:\\(Windows|Program)/i.test(m[1]))) { ajouter('CHM-001', l); break; }
    }
    // Courriel : comment le message part (aide à choisir la cible de remédiation).
    // Objets qui ont aussi une méthode Send sans être des courriels (requêtes HTTP, sockets).
    const autre = /^Set\s+(\w+)\s*=\s*(?:New\s+(?:MSXML2|WinHttp|MSWinsockLib)\.|CreateObject\s*\(\s*"§")/i.exec(c);
    if (autre && (/New/i.test(c) || l.chaines.some((x) => /XMLHTTP|WinHttp|Winsock/i.test(x)))) envoiNon.add(autre[1].toLowerCase());
    const dim = /\b(\w+)\s+As\s+(?:New\s+)?(?:MSXML2\.\w*XMLHTTP\w*|WinHttp\.\w+|MSWinsockLib\.\w+)/i.exec(c);
    if (dim) envoiNon.add(dim[1].toLowerCase());
    const envoi = /(\w+)\)?\.Send\s*(?:$|:)/i.exec(c);
    if (envoi && !envoiNon.has(envoi[1].toLowerCase()) && !/http|xml|req|sock/i.test(envoi[1])) stats.courriel.envoiDirect++;
    if (/\.Display\b/i.test(c)) stats.courriel.affichage++;
    if (/\.Attachments\.Add\b|\bAddAttachment\b/i.test(c)) stats.courriel.piecesJointes++;
    if (/\bGetNamespace\s*\(|\bGetDefaultFolder\b|\.Folders\s*\(|\bGetItemFromID\b/i.test(c)) stats.courriel.lectureBoite++;
    if (/\bSendObject\b/i.test(c)) { stats.courriel.envoiDirect++; if (/\bacSend\w+/i.test(c) && /,\s*(True|-1)\s*(,|$)/i.test(c)) stats.courriel.affichage++; }
    if (l.chaines.some((x) => SMTP_O365.test(x.trim()))) smtp ??= l;
    if (l.chaines.some((x) => SMTP_MDP.test(x)) || /\bsendpassword\b/i.test(c)) motDePasse = true;
    if (l.chaines.some((x) => /^CDO\.(Message|Configuration)$/i.test(x.trim()))) cdo ??= l;
  }
  if (smtp && motDePasse) ajouter('MEL-001', smtp);
  else if (cdo || smtp) ajouter('MEL-002', cdo ?? smtp);
  // Une même ligne ne compte qu'une fois par règle.
  const vus = new Set();
  return { ...stats, progIds: [...stats.progIds], courriel: stats.courriel, constats: constats.filter((x) => { const k = `${x.regle}:${x.ligne}`; if (vus.has(k)) return false; vus.add(k); return true; }) };
}

// Règles portant sur le projet entier (références) et sur le fichier (signature, ActiveX, format…).
export function constatsProjet(projet) {
  const out = [];
  let refVbs = false;
  const vb6 = projet.langage === 'vb6';
  for (const r of projet.references) {
    const cible = `${r.chemin} ${r.guid}`;
    for (const x of REFS) {
      if (x.re.test(cible)) {
        if (vb6 && VB6_SANS.has(x.regle)) continue;
        out.push({ regle: vb6 ? VB6_VERS[x.regle] ?? x.regle : x.regle, module: '(références)', ligne: 0, extrait: `${r.nom || x.quoi} — ${r.chemin || r.guid}` });
        if (x.regle === 'VBS-001') refVbs = true;
      }
    }
    if (r.type === 'project' || (r.chemin && !SYSTEME.test(r.chemin) && /^(?:\\\\|[a-z]:\\(?:users|documents and settings)\\|[d-z]:\\)/i.test(r.chemin))) {
      out.push({ regle: 'REF-001', module: '(références)', ligne: 0, extrait: `${r.nom} — ${r.chemin}` });
    }
  }
  if (vb6) {
    out.push({ regle: 'VB6-001', module: '(projet)', ligne: 0, extrait: `${projet.vb6.nom} · ${projet.vb6.type}${projet.vb6.exe ? ` · ${projet.vb6.exe}` : ''}` });
    for (const o of projet.objets ?? []) {
      const f = o.fichier.toLowerCase();
      if (/^msscript\.ocx$/.test(f)) out.push({ regle: 'VB6-010', module: '(composants)', ligne: 0, extrait: o.fichier });
      else if (/^msmapi32\.ocx$/.test(f)) out.push({ regle: 'OUT-003', module: '(composants)', ligne: 0, extrait: `${o.fichier} (contrôles MAPI)` });
      else if (!VB6_PRIS_EN_CHARGE.has(f)) out.push({ regle: 'VB6-002', module: '(composants)', ligne: 0, extrait: o.fichier });
      if (VB6_MATERIEL.test(f)) out.push({ regle: 'VB6-004', module: '(composants)', ligne: 0, extrait: o.fichier });
    }
  }
  return { constats: out, refVbs };
}

// ---------- Sources .NET (.cs, .vb, .csproj, .vbproj) et configurations (.config, .ini) ----------
// Lecture par motifs, ligne à ligne : on ne cherche que l'envoi de courriels.
export function analyserTexteAnnexe(texte, famille) {
  const constats = [];
  const lignes = texte.split(/\r\n|\r|\n/);
  const courriel = { envoiDirect: 0, affichage: 0, piecesJointes: 0, lectureBoite: 0 };
  let smtp = null, motDePasse = false;
  const a = (regle, i) => constats.push({ regle, module: '(source)', ligne: i + 1, extrait: lignes[i].trim().slice(0, 160) });
  lignes.forEach((l, i) => {
    const t = l.trim();
    if (!t || /^(\/\/|'|;|#|<!--)/.test(t)) return;
    if (/Microsoft\.Office\.Interop\.Outlook|"Outlook\.Application"/i.test(t)) a('OUT-001', i);
    if (/\bRedemption\b|"MAPI\.Session"/.test(t)) a('OUT-003', i);
    if (/smtp\.office365\.com|smtp-mail\.outlook\.com/i.test(t)) smtp ??= i;
    if (/NetworkCredential|\.Credentials\s*=|\bAuthenticate\s*\(|(?:pass(?:word)?|pwd|mot_?de_?passe|motdepasse)["']?\s*(?:=|:|value=)/i.test(t)) motDePasse = true;
    if (famille === 'dotnet') {
      if (/\.Send\s*\(|\.SendMailAsync\s*\(|\.SendAsync\s*\(/.test(t)) courriel.envoiDirect++;
      if (/\.Display\s*\(/.test(t)) courriel.affichage++;
      if (/\.Attachments\.Add\s*\(/.test(t)) courriel.piecesJointes++;
      if (/GetNamespace\s*\(|GetDefaultFolder\s*\(/.test(t)) courriel.lectureBoite++;
    }
  });
  if (smtp != null && motDePasse) a('MEL-001', smtp);
  return { constats: [...new Map(constats.map((c) => [`${c.regle}:${c.ligne}`, c])).values()], courriel };
}

export function constatsFichier(f, maintenant = new Date()) {
  const out = [];
  const a = (regle, extrait = '') => out.push({ regle, module: '(fichier)', ligne: 0, extrait });
  const s = f.signature;
  const signe = s.legacy || s.agile || s.v3;
  if (signe && !s.v3) a('SIG-001', s.agile ? 'signature agile sans V3' : 'signature ancienne sans V3');
  for (const c of s.certificats.filter((x) => !x.autorite)) {
    const fin = new Date(c.fin);
    if (fin < maintenant) a('SIG-002', `${c.sujet}, expiré le ${c.fin.slice(0, 10)}`);
    else if (fin - maintenant < 180 * 86400000) a('SIG-003', `${c.sujet}, expire le ${c.fin.slice(0, 10)}`);
  }
  for (const cls of f.activeX) a(CLSID_32.has(cls) ? 'B64-003' : 'ACX-001', `contrôle {${cls}}`);
  if (f.xlm) a('XLM-001', 'feuille de macros Excel 4.0');
  if (f.dde) a('DDE-001', 'liaison DDE dans le classeur');
  if (f.chiffre) a('FMT-002');
  if (f.pptNonAnalyse) a('FMT-004');
  if (f.famille === 'access' && f.accessVbaProbable) a('ACC-001', 'code VBA probable, non extrait');
  if (f.famille === 'script') a('VBS-010');
  if (f.format === 'ole' && f.projets.length) a('FMT-001');
  if (f.projets.some((p) => p.protege)) a('PRJ-001');
  return out;
}
