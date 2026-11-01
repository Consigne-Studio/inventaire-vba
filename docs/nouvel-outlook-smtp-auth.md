# Nouvel Outlook et SMTP AUTH : ce qui casse dans vos macros et vos applications, et quand

Envoyer un courriel depuis Excel, Access ou une petite application maison est l'un des usages les plus courants des macros : relances de factures, alertes qualité, rapports du lundi. Deux changements de Microsoft touchent ces envois en 2026 et 2027. Le plus proche n'est pas celui dont on parle le plus.

## Les deux échéances

| Échéance | Ce qui change | Ce qui casse |
|---|---|---|
| **Fin décembre 2026** | Exchange Online désactive par défaut l'authentification de base de SMTP AUTH. Les locataires créés à partir de 2027 n'y ont plus accès. | Tout envoi vers `smtp.office365.com` avec un identifiant et un mot de passe : CDO dans une macro, `SmtpClient` dans un outil .NET, copieur ou logiciel métier configuré en SMTP. |
| **Mars 2027** | Phase « opt-out » du nouvel Outlook en entreprise : les postes Microsoft 365 passent au nouvel Outlook par défaut. | Tout code qui pilote Outlook : `CreateObject("Outlook.Application")`, référence à la bibliothèque Outlook, `DoCmd.SendObject` dans Access, MAPI, CDO 1.21, Redemption. Le nouvel Outlook n'a ni modèle objet COM, ni VBA, ni MAPI. |

Les deux échéances se repoussent, mais pas indéfiniment. L'administrateur peut réactiver SMTP AUTH jusqu'à son retrait définitif, dont Microsoft annoncera la date au second semestre 2027. Un utilisateur peut revenir à Outlook classique, pris en charge « au moins jusqu'en 2029 ». Ces sursis n'existent que si quelqu'un sait quels outils en dépendent.

## Le piège : passer d'Outlook à SMTP

Le contournement le plus souvent conseillé pour sortir d'Outlook, y compris par certains éditeurs de logiciels, est de passer l'envoi en SMTP. S'il utilise un mot de passe vers Exchange Online, il tombe avec la première échéance, trois mois avant la seconde. Remplacer une dépendance par une autre qui casse plus tôt ne règle rien.

## La sortie durable : Microsoft Graph, limité à une boîte

Microsoft Graph envoie un courriel par une requête HTTPS, sans Outlook et sans mot de passe d'utilisateur. Une application déclarée dans Entra ID reçoit le droit « Application Mail.Send ». Le contrôle d'accès par rôle d'Exchange Online (RBAC pour les applications) limite ce droit à une seule boîte d'envoi, par exemple `envois@votre-domaine.fr`. Sans cette limite, l'application pourrait écrire au nom de n'importe qui.

Côté code, l'appel reste aussi court qu'avec Outlook :

`EnvoyerCourriel "client@exemple.fr", "Relance facture 2024-118", "Veuillez trouver la facture jointe.", "C:\Factures\2024-118.pdf"`

Le module libre `ModGraphMail.bas` fait cela en VBA et en VB6, sans DLL, en Office 32 et 64 bits. Il gère les pièces jointes jusqu'à 150 Mo et le brouillon à relire, qui remplace `.Display`, et il lève des erreurs qui disent quoi vérifier. Le secret de l'application est rangé dans le Gestionnaire d'identification Windows, jamais dans le code. Le guide de l'administrateur décrit la mise en place en 20 minutes.

## Trouver tous les points d'envoi avant qu'ils cassent

Le plus difficile n'est pas de corriger un envoi, c'est de les trouver tous : la macro du contrôle de gestion, la base Access de la qualité, l'application VB6 du banc de test, le petit outil .NET de relances écrit il y a six ans.

L'outil libre inventaire-vba parcourt vos partages, hors ligne et sans rien modifier. Il signale chaque point d'envoi avec son mécanisme, la façon dont le message part (envoi direct, affichage, pièces jointes, lecture de la boîte) et la cible recommandée :

* **OUT-001** : automatisation d'Outlook, dans les macros, les projets VB6 et les projets .NET (interop Outlook) ;
* **OUT-002** : `DoCmd.SendObject` dans Access ;
* **OUT-003** : MAPI, CDO 1.21, Redemption ;
* **MEL-001** : SMTP avec mot de passe vers Exchange Online, dans le code VBA, VB6 et .NET et dans les fichiers `.config` et `.ini` ;
* **MEL-002** : envoi SMTP par CDO vers un autre serveur, à vérifier.

Le rapport les regroupe dans une section « Envoi de courriels » et dans l'onglet « Courriels » du classeur Excel.

## Par où commencer

1. Lancer l'inventaire sur les partages et les dossiers de sources. L'édition gratuite analyse 200 fichiers par inventaire.
2. Traiter d'abord les envois MEL-001, qui cassent fin décembre 2026, puis les envois par Outlook, avant mars 2027.
3. Faire préparer la boîte d'envoi et l'application Entra ID par l'administrateur, à l'aide du guide.
4. Remplacer chaque envoi par `EnvoyerCourriel`, ou confier la remédiation à prix fixe, réalisée entièrement à distance.

Sources : Microsoft, MC786329 (calendrier de SMTP AUTH en authentification de base, mis à jour en janvier 2026) ; Microsoft, comparaison des fonctionnalités d'Outlook classique et du nouvel Outlook (support.microsoft.com/help/5038254) ; Microsoft, « Role Based Access Control for Applications in Exchange Online » (learn.microsoft.com/graph/auth-limit-mailbox-access) ; Directions on Microsoft, report du nouvel Outlook à 2027.
