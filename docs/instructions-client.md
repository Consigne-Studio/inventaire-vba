# Après votre commande

Merci pour votre commande de diagnostic. Tout se fait à distance : vous lancez l'outil chez vous, vous nous envoyez un fichier sans code, nous vous renvoyons le rapport. Comptez 15 minutes de votre temps.

## 1. Télécharger l'outil

Téléchargez l'édition gratuite pour Windows : https://github.com/Consigne-Studio/inventaire-vba/releases/latest, fichier `inventaire-vba-1.0.0-windows-gratuite.zip`, puis décompressez-le. Aucune installation et aucun droit d'administrateur ne sont nécessaires. L'outil n'est pas encore signé : si Windows SmartScreen s'affiche, cliquez sur « Informations complémentaires » puis « Exécuter quand même ». L'empreinte du fichier figure dans `SHA256SUMS.txt`.

Formule Complet : la licence de parc incluse vous est envoyée avec l'accusé de réception. Vous pouvez faire le diagnostic avec l'édition gratuite, l'option `--diagnostic` lève la limite de 200 fichiers.

## 2. Lancer l'inventaire en mode diagnostic

Sur un poste qui voit vos partages, ouvrez une invite de commandes dans le dossier décompressé et lancez, en indiquant vos dossiers :

`inventaire-vba.exe \\serveur\partage1 \\serveur\partage2 --diagnostic --organisation "Votre organisation" -o C:\Diagnostic-VBA`

Ajoutez `--sans-chemins` si les noms de vos fichiers ou dossiers sont eux-mêmes confidentiels : ils seront remplacés par des alias (dossier-001, fichier-001). Gardez le fichier `diagnostic-correspondance.csv` écrit à côté : il relie chaque alias à son fichier réel. Ne l'envoyez pas, il vous servira à lire le rapport livré.

L'outil lit les fichiers sans les ouvrir dans Office, sans exécuter de macro et sans connexion réseau. Comptez quelques secondes par millier de fichiers.

**Bases Access.** Si vous avez des bases `.accdb` ou `.mdb` avec du code, lancez d'abord sur un poste équipé d'Access le script fourni, qui exporte le code depuis une copie, macros désactivées : `powershell -ExecutionPolicy Bypass -File exporter-access.ps1 -Dossier "\\serveur\Qualite" -Sortie C:\Diagnostic-VBA\access`. Ajoutez ensuite `C:\Diagnostic-VBA\access` aux dossiers de la commande ci-dessus.

## 3. Vérifier puis envoyer diagnostic.json

Ouvrez `diagnostic-apercu.html` dans le dossier de sortie : c'est exactement ce que contient le fichier à envoyer. Il ne contient ni code, ni extrait de code, ni auteur, ni signataire.

Envoyez `diagnostic.json` en réponse au courriel de confirmation de commande (l'adresse de contact figure aussi sur votre reçu), en indiquant le numéro de commande. Ajoutez si vous le souhaitez, en quelques lignes : les versions d'Office de votre parc (Microsoft 365, 2021, LTSC…), la part des postes en 64 bits, vos échéances (passage au nouvel Outlook, renouvellement des postes).

## 4. Recevoir le rapport

Nous accusons réception dans la journée ouvrée et vous donnons la date de livraison : 5 jours ouvrés (Essentiel) ou 10 jours ouvrés (Complet) après réception du fichier. Vous recevrez le rapport PDF relu et commenté, le classeur Excel et les exports CSV, puis vous pourrez poser vos questions par écrit pendant 30 jours.

Le fichier reçu est supprimé 90 jours après la livraison, ou plus tôt sur simple demande.
