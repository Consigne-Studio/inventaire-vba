# Guide de l'administrateur Microsoft 365

Ce guide s'adresse à l'administrateur Microsoft 365 de l'organisation, salarié ou prestataire. Il prépare l'envoi de courriels par Microsoft Graph pour les macros, les bases Access et les applications VB6 qui utilisent le module `ModGraphMail`. Il remplace l'automatisation d'Outlook, que le nouvel Outlook ne permet plus, et le SMTP par mot de passe, qu'Exchange Online désactive par défaut fin décembre 2026.

Comptez 20 minutes. Il faut un compte qui peut créer une application dans Microsoft Entra ID (rôle Administrateur d'applications) et un compte qui peut gérer les rôles d'Exchange Online (rôle Administrateur Exchange). Aucun droit n'est donné sur les autres boîtes de l'organisation.

## Ce que vous allez créer

* Une **boîte d'envoi** : une boîte partagée existante ou nouvelle, par exemple `envois@votre-domaine.fr`. Les courriels partiront de cette adresse et seront rangés dans ses Éléments envoyés.
* Une **application Entra ID** sans autorisation globale. Elle sert seulement d'identité.
* Un **droit « Application Mail.Send » limité à cette seule boîte**, par le contrôle d'accès par rôle (RBAC) pour les applications d'Exchange Online. C'est la méthode qui remplace désormais les stratégies d'accès d'application. Sans cette limite, `Mail.Send` permettrait d'envoyer au nom de n'importe quel utilisateur.

## 1. Créer l'application dans Entra ID

1. Ouvrez le centre d'administration Microsoft Entra, puis Applications, Inscriptions d'applications, Nouvelle inscription.
2. Nom : `Envoi de courriels - macros` ou un nom parlant. Types de comptes : « Comptes dans cet annuaire d'organisation uniquement ». Pas d'URI de redirection.
3. Notez l'**ID d'application (client)** et l'**ID de l'annuaire (locataire)**, affichés sur la page de présentation.
4. **N'ajoutez aucune autorisation d'API** et ne donnez aucun consentement administrateur. Si une autorisation `Mail.Send` est déjà présente, retirez-la et révoquez son consentement : les droits d'Entra s'ajoutent à ceux d'Exchange, et un `Mail.Send` accordé dans Entra annulerait la limite à une boîte.
5. Certificats et secrets, Nouveau secret client. Durée : 12 mois, ou 24 mois au plus. Copiez la **valeur** du secret tout de suite, car elle ne sera plus affichée. Notez la date d'expiration dans votre agenda, avec un rappel un mois avant.
6. Ouvrez Applications d'entreprise, retrouvez l'application par son nom, et notez son **ID d'objet**. C'est l'ID du principal de service, différent de celui de la page Inscriptions d'applications.

## 2. Déclarer l'application dans Exchange Online

Dans PowerShell, avec le module ExchangeOnlineManagement :

`Connect-ExchangeOnline`

`New-ServicePrincipal -AppId <ID d'application> -ObjectId <ID d'objet de l'application d'entreprise> -DisplayName "Envoi de courriels - macros"`

## 3. Limiter le droit d'envoi à une seule boîte

`New-ManagementScope -Name "Envoi macros - une boîte" -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'envois@votre-domaine.fr'"`

`New-ManagementRoleAssignment -App <ID d'objet de l'application d'entreprise> -Role "Application Mail.Send" -CustomResourceScope "Envoi macros - une boîte"`

Pour autoriser une seconde boîte plus tard, par exemple une boîte d'un autre service, élargissez le filtre : `"PrimarySmtpAddress -eq 'envois@votre-domaine.fr' -or PrimarySmtpAddress -eq 'compta@votre-domaine.fr'"`.

## 4. Vérifier

`Test-ServicePrincipalAuthorization -Identity "Envoi de courriels - macros" -Resource envois@votre-domaine.fr | Format-Table`

La colonne `InScope` doit afficher `True` pour le rôle `Application Mail.Send`. Faites le même test avec la boîte d'un utilisateur ordinaire : elle doit afficher `False`. Le test répond tout de suite, mais l'application elle-même peut mettre de 30 minutes à 2 heures avant de voir un changement de droits.

## 5. Ranger l'identifiant sur les postes

Sur chaque poste qui envoie, en session de l'utilisateur, dans une invite de commandes :

`cmdkey /generic:GraphMail /user:<ID d'application>  /pass`

La commande demande le mot de passe : collez la valeur du secret. Elle est rangée dans le Gestionnaire d'identification Windows de l'utilisateur et n'apparaît ni dans le code, ni dans un fichier, ni dans l'historique de la console. Pour un déploiement en nombre, la même commande peut être lancée par un script de connexion, à condition que le secret n'y figure pas en clair. Le plus simple reste de la lancer avec l'utilisateur.

Dans le module, renseignez `GRAPH_LOCATAIRE` (votre domaine `onmicrosoft.com` ou l'ID de l'annuaire) et `GRAPH_EXPEDITEUR` (la boîte d'envoi), puis lancez `GraphTester "vous@votre-domaine.fr"` depuis l'éditeur VBA. Le courriel de test doit arriver en moins d'une minute.

## 6. Après la bascule

* **SMTP par mot de passe** : quand plus aucun outil ne l'utilise, désactivez-le sur la boîte qui servait aux envois, avec `Set-CASMailbox -Identity envois@votre-domaine.fr -SmtpClientAuthenticationDisabled $true`. Changez aussi son ancien mot de passe, qui figurait dans du code ou dans un fichier de configuration.
* **Secret** : avant son expiration, créez un nouveau secret dans Entra ID, relancez `cmdkey` sur les postes, puis supprimez l'ancien secret.
* **Révocation** : supprimer le secret, ou l'attribution de rôle (`Remove-ManagementRoleAssignment`), coupe l'envoi immédiatement sur tous les postes.

## Questions fréquentes

**Pourquoi un secret plutôt qu'un certificat ?** VBA et VB6 ne savent pas signer une assertion de certificat sans bibliothèque supplémentaire. Le secret est limité à une seule boîte, n'est jamais écrit dans un fichier et se révoque en un clic. Pour une application .NET, préférez un certificat.

**Les utilisateurs doivent-ils avoir accès à la boîte d'envoi ?** Non pour l'envoi direct. Oui pour `AfficherBrouillon`, qui remplace `.Display` : le brouillon s'ouvre dans Outlook sur le web, dans la boîte d'envoi. Donnez alors aux utilisateurs concernés l'accès complet à la boîte partagée.

**Faut-il une licence pour la boîte d'envoi ?** Une boîte partagée de moins de 50 Go n'a pas besoin de licence. Les limites d'envoi d'Exchange Online s'appliquent : 10 000 destinataires par jour et 35 Mo par message par défaut.

**Le proxy de l'entreprise bloque-t-il l'envoi ?** Le module passe par les réglages Internet de l'utilisateur, comme le navigateur. Il doit pouvoir joindre `login.microsoftonline.com` et `graph.microsoft.com` en HTTPS. Si le proxy est réglé avec `netsh winhttp` plutôt que dans les options Internet, utilisez `GraphParametresAvances` avec `MSXML2.ServerXMLHTTP.6.0`.

Sources : Microsoft, « Role Based Access Control for Applications in Exchange Online » (learn.microsoft.com/graph/auth-limit-mailbox-access) ; message MC786329, retrait de SMTP AUTH en authentification de base (calendrier de janvier 2026).
