# Journal des versions

## 1.1.0 — 2026-10-28

* **Envoi de courriels** : nouvelles règles OUT-002 (`DoCmd.SendObject`), OUT-003 (MAPI, CDO 1.21, Redemption), MEL-001 (SMTP par mot de passe vers Exchange Online, désactivé par défaut fin décembre 2026) et MEL-002 (CDO vers un autre serveur). Section « Envoi de courriels » dans le rapport, onglet « Courriels » dans le classeur, avec la façon dont chaque message part et la cible recommandée.
* **Kit de remédiation** (`remediation/`) : module `ModGraphMail.bas` (Microsoft Graph, VBA et VB6, sans DLL, licence MIT) et guide de l'administrateur Microsoft 365 (droit d'envoi limité à une boîte).
* **Visual Basic 6** : lecture des projets `.vbp` et de leurs sources. Règles VB6-001, VB6-002 (composant hors du support Microsoft), VB6-004 (matériel) et VB6-010 (VBScript, en P1). Les règles 64 bits ne s'appliquent pas aux projets VB6. Estimation de migration vers C# .NET en édition parc.
* **Sources .NET et configurations** (`.cs`, `.vb`, `.csproj`, `.vbproj`, `.config`, `.ini`) : lues pour l'envoi de courriels seulement, hors limite de l'édition gratuite.
* 39 règles. Échéance « SMTP AUTH » ajoutée au calendrier du rapport.

## 1.0.0 — 2026-10-24

Première version publique.

* Lecture hors ligne des classeurs, documents et présentations Office : `.xlsm`, `.xlsb`, `.xlam`, `.xls`, `.xla`, `.docm`, `.dotm`, `.doc`, `.dot`, `.pptm`, `.potm`, `.ppam`, modules exportés `.bas`, `.cls`, `.frm`, scripts `.vbs`. Bases Access via l'export PowerShell fourni.
* 31 règles sur six axes : 64 bits, signatures V3, VBScript, nouvel Outlook, ActiveX, Internet Explorer, DDE, macros Excel 4.0, références, sécurité, chemins et connexions.
* Rapport HTML autonome (sans ressource externe, imprimable), `inventaire.json`, fichier de diagnostic anonymisé avec aperçu.
* Édition parc : classeur Excel en 7 feuilles, CSV, estimation de migration sur trois chemins, répartition par dossier, copies et familles, compatibilité LibreOffice, comparaison entre deux inventaires.
* Les fichiers au motif de macro malveillante (SEC-010) sont isolés et sortis de l'estimation.
* Exécutable Windows autonome (sans Node.js) et installeur MSI.
