# Revue du diagnostic — Exemple Industrie (organisation fictive)

## Synthèse

Les partages contiennent 24 fichiers Office avec du VBA, une application Visual Basic 6 (le banc de test de la ligne 2) et un petit outil .NET de relances. **15 fichiers cassent** à une échéance connue. La moitié casse dès qu'un poste passe en Office 64 bits, ce qui est déjà le cas des postes neufs livrés avec Microsoft 365.

L'échéance la plus proche est celle que personne n'attend : **fin décembre 2026**, Exchange Online désactive l'envoi SMTP par mot de passe. Les relances clients de la Comptabilité et l'outil de relances de l'Informatique envoient ainsi, avec un mot de passe écrit en clair. Six autres fichiers envoient par Outlook et cesseront de le faire avec le nouvel Outlook, en mars 2027 : deux classeurs, un modèle Word, l'export Access de la Qualité, le banc de test et l'outil .NET.

Le reste du risque tient en trois familles : les appels à l'API Windows sans `PtrSafe` (4 classeurs de la Direction et de la Finance, qui embarquent le même module `modApi`), les deux classeurs signés avec une signature ancienne, et VBScript, dont dépendent les scripts de nuit et le banc de test.

Budget : environ **15 jours** pour tout corriger sans réécrire, en commençant par les envois de courriels et `modApi`. Un fichier reçu par courriel (`Facture_2024_0172.docm`) présente un motif de macro malveillante : il faut l'isoler dès aujourd'hui.

## Remarques par constat

- MEL-001 : les deux outils de relances passent sur Microsoft Graph avec le module `ModGraphMail`, et une seule boîte d'envoi suffit (`compta@`). L'administrateur prépare l'application Entra ID avec le guide fourni, en 20 minutes. Les deux mots de passe en clair sont à changer le jour même.
- OUT-001 : le banc de test et les classeurs passent sur le même module. Si un utilisateur veut relire le message avant l'envoi, `AfficherBrouillon` remplace `.Display`. Garder Outlook classique reste possible jusqu'en 2029 au moins, mais chaque poste renouvelé recule l'échéance d'autant.
- OUT-002 : dans la base de la Qualité, `DoCmd.SendObject` se remplace par `DoCmd.OutputTo`, puis `EnvoyerCourriel` avec la pièce jointe : une heure.
- VB6-010 : le banc de test valide ses trames avec `VBScript.RegExp`. Remplacez-le par une vérification écrite en VB6, une vingtaine de lignes, avant que Windows désactive VBScript.
- VB6-002 : `THREED32.OCX` n'est plus pris en charge par Microsoft. Gardez l'installateur d'origine. Ce composant sert seulement à un panneau d'état, facile à remplacer par un `Label`.
- B64-001 : les 4 classeurs contiennent le même module `modApi`, avec les mêmes déclarations `Sleep` et `GetUserName` aux mêmes lignes. Corrigez-le une fois, puis réimportez-le partout : 1 jour au lieu de 4.
- SIG-001 : re-signer les deux classeurs avec le certificat interne actuel suffit. Activez ensuite la stratégie « V3 uniquement » sur un groupe pilote avant de la généraliser.
- SEC-010 : faites examiner ce fichier par l'équipe sécurité avant toute ouverture. Il n'est pas chiffré dans l'estimation.
- VBS-010 : les deux scripts de nuit tournent sur le serveur de fichiers. Portez-les en PowerShell, une demi-journée chacun.

## Plan d'action

1. Aujourd'hui : isoler `Facture_2024_0172.docm` et prévenir l'équipe sécurité. Changer les mots de passe des boîtes `compta@` et `relances@`, qui figurent en clair dans le code.
2. Avant décembre 2026 : faire préparer la boîte d'envoi par l'administrateur (guide fourni), puis passer « Relances clients » et l'outil .NET de relances sur Microsoft Graph (MEL-001).
3. Dans le mois : corriger `modApi` (B64-001) et le diffuser dans les 4 classeurs. Passer les handles de `ExportPDF.bas` en `LongPtr` (B64-002). Remplacer Jet 4.0 et DAO 3.6 par ACE dans « Suivi des ordres de fabrication » (B64-004).
4. Avant d'activer la stratégie V3 : re-signer « Rapport qualité » et « Suivi budgétaire 2024 ». Supprimer « Commandes fournisseurs - Copie », copie exacte de l'original.
5. Avant mars 2027 : passer sur le même module les six autres envois par Outlook, dont le banc de test, puis retirer `VBScript.RegExp` du banc (VB6-010).
6. Ce qui peut attendre : les références manquantes et les chemins en dur des archives, qui ne servent plus.
