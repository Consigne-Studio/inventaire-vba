# Remédiation de l'envoi de courriels

Ce dossier remplace les envois de courriels que l'inventaire signale par les règles OUT-001, OUT-002, OUT-003, MEL-001 et MEL-002. Ces envois cassent avec le nouvel Outlook (mars 2027) ou avec la fin du SMTP par mot de passe dans Exchange Online (fin décembre 2026).

* `ModGraphMail.bas` est le module à importer dans le projet VBA (Fichier, Importer) ou dans le projet VB6 (Projet, Ajouter un module). Il fonctionne sans Outlook, sans SMTP, sans DLL ni référence à cocher, en Office 32 et 64 bits.
* `guide-administrateur.pdf` est à transmettre à l'administrateur Microsoft 365 : 20 minutes, une seule boîte d'envoi autorisée.

Le module utilise uniquement MSXML et le Gestionnaire d'identification Windows, présents sur tous les postes Windows depuis Windows 7. Il est testé de bout en bout contre une imitation de Microsoft Graph : jeton, envoi direct, accents, pièces jointes, brouillon, téléversement par blocs au-delà de 3 Mo et messages d'erreur.

## Avant, après

**Excel ou Word qui pilote Outlook (OUT-001)**

Avant :

`Set ol = CreateObject("Outlook.Application")`
`Set msg = ol.CreateItem(0)`
`msg.To = "controle.gestion@exemple.fr"`
`msg.Subject = "Suivi du " & Format(Date, "dd/mm/yyyy")`
`msg.Attachments.Add "\\srv-fichiers\finance\exports\suivi.pdf"`
`msg.Send`

Après :

`EnvoyerCourriel "controle.gestion@exemple.fr", "Suivi du " & Format(Date, "dd/mm/yyyy"), "Suivi joint.", "\\srv-fichiers\finance\exports\suivi.pdf"`

Si la macro appelait `.Display` pour laisser l'utilisateur relire, remplacez-le par `AfficherBrouillon` avec les mêmes paramètres : le brouillon s'ouvre dans Outlook sur le web.

**Envoi par CDO vers smtp.office365.com avec un mot de passe (MEL-001)**

Supprimez la configuration CDO (`smtpserver`, `sendusername`, `sendpassword`), puis remplacez `msg.Send` par :

`EnvoyerCourriel dest, "Relance facture " & numero, texte, pdf`

Demandez ensuite à l'administrateur de changer le mot de passe qui figurait dans le code.

**Access : DoCmd.SendObject (OUT-002)**

Avant :

`DoCmd.SendObject acSendReport, "rptNonConformite", acFormatPDF, "qualite@exemple.fr", , , "Non-conformité " & idNc, "Fiche jointe.", False`

Après :

`Dim pdf As String: pdf = Environ("TEMP") & "\NC-" & idNc & ".pdf"`
`DoCmd.OutputTo acOutputReport, "rptNonConformite", acFormatPDF, pdf`
`EnvoyerCourriel "qualite@exemple.fr", "Non-conformité " & idNc, "Fiche jointe.", pdf`
`Kill pdf`

**Application VB6 (OUT-001 dans un projet .vbp)**

Ajoutez `ModGraphMail.bas` au projet, remplacez les appels à Outlook comme ci-dessus, puis recompilez. Le module compile en VB6 : la branche « #If VBA7 » est ignorée et la version 32 bits des déclarations est utilisée.

## Paramètres

Dans l'en-tête du module :

* `GRAPH_LOCATAIRE` : votre domaine `onmicrosoft.com` ou l'ID de l'annuaire ;
* `GRAPH_EXPEDITEUR` : la boîte d'envoi autorisée par l'administrateur ;
* `GRAPH_COFFRE` : le nom de l'identifiant rangé par `cmdkey`, `GraphMail` par défaut.

`EnvoyerCourriel destinataires, objet, corps, [pieces], [copie], [copieCachee], [expediteur]`

* Les destinataires sont séparés par un point-virgule, comme dans Outlook : `a@exemple.fr; Service achats <achats@exemple.fr>`.
* Un corps qui commence par `<` part en HTML. Sinon, il part en texte.
* Les pièces jointes sont des chemins séparés par `|`, jusqu'à 150 Mo chacune. Le message reste soumis à la limite d'Exchange Online, 35 Mo par défaut.
* En cas d'échec, une erreur est levée avec la cause donnée par Microsoft et ce qu'il faut vérifier. Interceptez-la avec `On Error GoTo` si la macro doit continuer.

## Licence

`ModGraphMail.bas` est sous licence MIT : vous pouvez le modifier et le redistribuer dans vos outils.

Si vous préférez qu'on le fasse pour vous, la remédiation à prix fixe est réalisée entièrement à distance. Vous envoyez les fichiers concernés et vous recevez les fichiers corrigés, avec les commandes de l'administrateur préparées pour votre boîte d'envoi.
