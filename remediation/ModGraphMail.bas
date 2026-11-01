Attribute VB_Name = "ModGraphMail"
' ModGraphMail 1.0 : envoi de courriels par Microsoft Graph, sans Outlook, sans SMTP et sans DLL.
' Pour VBA (Excel, Word, Access, Office 32 ou 64 bits) et Visual Basic 6. Licence MIT.
' Fourni avec inventaire-vba (dossier remediation) : https://consigne-studio.github.io/inventaire-vba/
'
' Remplace les envois qui cassent :
'   - CreateObject("Outlook.Application") puis .Send ou .Display : le nouvel Outlook n'a pas de COM ;
'   - DoCmd.SendObject (Access) et MAPI : le nouvel Outlook ne fournit pas MAPI ;
'   - CDO ou SMTP vers smtp.office365.com avec un mot de passe : désactivé par défaut fin décembre 2026.
'
' Avant le premier envoi :
'   1. L'administrateur Microsoft 365 suit guide-administrateur.pdf : application Entra ID, droit
'      « Application Mail.Send » limité à UNE boîte d'envoi (RBAC pour les applications d'Exchange Online).
'   2. Sur chaque poste, l'identifiant de l'application est rangé dans le Gestionnaire d'identification
'      Windows (le secret est demandé, jamais écrit dans le code ni dans un fichier) :
'         cmdkey /generic:GraphMail /user:<ID d'application (client)> /pass
'   3. Renseigner les deux constantes ci-dessous, puis lancer GraphTester "vous@votre-domaine.fr".
'
' Utilisation :
'   EnvoyerCourriel "client@exemple.fr; Service achats <achats@exemple.fr>", "Relance facture 2024-118", _
'                   "<p>Bonjour,</p><p>Veuillez trouver la facture jointe.</p>", "C:\Factures\2024-118.pdf"
'   AfficherBrouillon ...   (mêmes paramètres) remplace .Display : le brouillon s'ouvre dans Outlook sur le web.
' Plusieurs pièces jointes : chemins séparés par « | ». Un corps qui commence par « < » est envoyé en HTML.
' En cas d'échec, une erreur est levée (Err.Number = vbObjectError + 4001 à 4003) avec la cause et quoi vérifier.
Option Explicit

' === À renseigner ===
Public Const GRAPH_LOCATAIRE As String = "votre-domaine.onmicrosoft.com"  ' domaine ou ID du locataire Microsoft 365
Public Const GRAPH_EXPEDITEUR As String = "envois@votre-domaine.fr"       ' boîte d'envoi autorisée par l'administrateur
Public Const GRAPH_COFFRE As String = "GraphMail"                         ' nom de l'identifiant (cmdkey /generic:…)

'#coffre-debut
#If VBA7 Then
Private Declare PtrSafe Function CredReadW Lib "advapi32.dll" (ByVal TargetName As LongPtr, ByVal CredType As Long, ByVal Flags As Long, ByRef Credential As LongPtr) As Long
Private Declare PtrSafe Sub CredFree Lib "advapi32.dll" (ByVal Buffer As LongPtr)
Private Declare PtrSafe Sub CopierMemoire Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
#Else
Private Declare Function CredReadW Lib "advapi32.dll" (ByVal TargetName As Long, ByVal CredType As Long, ByVal Flags As Long, ByRef Credential As Long) As Long
Private Declare Sub CredFree Lib "advapi32.dll" (ByVal Buffer As Long)
Private Declare Sub CopierMemoire Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
#End If
'#coffre-fin

' Limites de Microsoft Graph : une requête fait au plus 4 Mo. Au-delà de 2,5 Mo de pièces jointes, le message
' passe par un brouillon ; une pièce de 3 Mo ou plus est téléversée par blocs (multiples de 320 Kio).
Private Const SEUIL_ENVOI_DIRECT As Double = 2621440
Private Const SEUIL_PIECE As Double = 3145728
Private Const TAILLE_BLOC As Long = 983040
Private Const TAILLE_MAX As Double = 157286400

Private mJeton As String
Private mExpire As Date
Private mClientId As String
Private mSecret As String
Private mUrlConnexion As String
Private mUrlGraph As String
Private mHttp As String

' ---------------------------------------------------------------------------------------------------------
' Envoie un courriel depuis GRAPH_EXPEDITEUR (ou expediteur, si l'administrateur l'a aussi autorisé).
' Le message est enregistré dans les Éléments envoyés de la boîte d'envoi.
Public Function EnvoyerCourriel(ByVal destinataires As String, ByVal objet As String, ByVal corps As String, Optional ByVal piecesJointes As String = "", Optional ByVal copie As String = "", Optional ByVal copieCachee As String = "", Optional ByVal expediteur As String = "") As Boolean
    Dim boite As String, fichiers As Variant, total As Double, reponse As String, statut As Long, id As String
    boite = BoiteEnvoi(expediteur)
    fichiers = ListeFichiers(piecesJointes)
    total = TailleTotale(fichiers)
    If total <= SEUIL_ENVOI_DIRECT Then
        statut = Requete("POST", Graph() & "users/" & EncoderUrl(boite) & "/sendMail", "{""message"":" & MessageJson(destinataires, objet, corps, copie, copieCachee, PiecesJson(fichiers)) & ",""saveToSentItems"":true}", "application/json", True, reponse, "")
        If statut <> 202 Then Echec "Envoi du courriel", statut, reponse, boite
    Else
        id = CreerMessage(boite, destinataires, objet, corps, copie, copieCachee, fichiers, reponse)
        statut = Requete("POST", Graph() & "users/" & EncoderUrl(boite) & "/messages/" & EncoderUrl(id) & "/send", "", "", True, reponse, "")
        If statut <> 202 Then Echec "Envoi du brouillon", statut, reponse, boite
    End If
    EnvoyerCourriel = True
End Function

' Crée le message dans les Brouillons de la boîte d'envoi et renvoie son lien Outlook sur le web.
Public Function CreerBrouillon(ByVal destinataires As String, ByVal objet As String, ByVal corps As String, Optional ByVal piecesJointes As String = "", Optional ByVal copie As String = "", Optional ByVal copieCachee As String = "", Optional ByVal expediteur As String = "") As String
    Dim reponse As String
    CreerMessage BoiteEnvoi(expediteur), destinataires, objet, corps, copie, copieCachee, ListeFichiers(piecesJointes), reponse
    CreerBrouillon = ValeurJson(reponse, "webLink")
End Function

' Remplace .Display : ouvre le brouillon pour relecture. L'utilisateur doit avoir accès à la boîte d'envoi.
Public Sub AfficherBrouillon(ByVal destinataires As String, ByVal objet As String, ByVal corps As String, Optional ByVal piecesJointes As String = "", Optional ByVal copie As String = "", Optional ByVal copieCachee As String = "", Optional ByVal expediteur As String = "")
    Dim lien As String
    lien = CreerBrouillon(destinataires, objet, corps, piecesJointes, copie, copieCachee, expediteur)
    If Len(lien) > 0 Then CreateObject("Shell.Application").ShellExecute lien
End Sub

' Vérifie la configuration de bout en bout : connexion, droit d'envoi, réception.
Public Sub GraphTester(ByVal destinataire As String)
    EnvoyerCourriel destinataire, "Test ModGraphMail", "Ce message confirme que l'envoi par Microsoft Graph fonctionne depuis ce poste."
    MsgBox "Courriel de test envoyé à " & destinataire & " depuis " & GRAPH_EXPEDITEUR & ".", vbInformation, "ModGraphMail"
End Sub

' Facultatif : identifiants fournis par le code appelant (par exemple lus dans un coffre de l'entreprise)
' au lieu du Gestionnaire d'identification Windows.
Public Sub GraphIdentifiants(ByVal idApplication As String, ByVal secret As String)
    mClientId = idApplication
    mSecret = secret
    mJeton = ""
End Sub

' Facultatif : points de terminaison d'un cloud national, ou objet HTTP (MSXML2.ServerXMLHTTP.6.0 si le
' proxy est réglé par netsh winhttp plutôt que dans les options Internet).
Public Sub GraphParametresAvances(ByVal urlConnexion As String, ByVal urlGraph As String, ByVal objetHttp As String)
    mUrlConnexion = urlConnexion
    mUrlGraph = urlGraph
    mHttp = objetHttp
    mJeton = ""
End Sub

' ---------------------------------------------------------------------------------------------------------
Private Function CreerMessage(ByVal boite As String, ByVal destinataires As String, ByVal objet As String, ByVal corps As String, ByVal copie As String, ByVal copieCachee As String, ByVal fichiers As Variant, ByRef reponse As String) As String
    Dim statut As Long, id As String, base As String, i As Long, creation As String
    statut = Requete("POST", Graph() & "users/" & EncoderUrl(boite) & "/messages", MessageJson(destinataires, objet, corps, copie, copieCachee, ""), "application/json", True, reponse, "")
    If statut <> 201 Then Echec "Création du brouillon", statut, reponse, boite
    creation = reponse
    id = ValeurJson(reponse, "id")
    base = Graph() & "users/" & EncoderUrl(boite) & "/messages/" & EncoderUrl(id) & "/attachments"
    For i = 0 To UBound(fichiers)
        If TailleFichier(fichiers(i)) < SEUIL_PIECE Then
            statut = Requete("POST", base, PieceJson(fichiers(i)), "application/json", True, reponse, "")
            If statut <> 201 Then Echec "Ajout de " & NomFichier(fichiers(i)), statut, reponse, boite
        Else
            Televerser base, fichiers(i), boite
        End If
    Next
    reponse = creation
    CreerMessage = id
End Function

' Pièce de 3 Mo ou plus : session de téléversement, puis blocs envoyés sans jeton (l'adresse est signée).
Private Sub Televerser(ByVal base As String, ByVal chemin As String, ByVal boite As String)
    Dim taille As Double, statut As Long, reponse As String, adresse As String, debut As Double, n As Long, octets As Variant
    taille = TailleFichier(chemin)
    statut = Requete("POST", base & "/createUploadSession", "{""AttachmentItem"":{""attachmentType"":""file"",""name"":""" & Echapper(NomFichier(chemin)) & """,""size"":" & CStr(taille) & "}}", "application/json", True, reponse, "")
    If statut <> 201 Then Echec "Téléversement de " & NomFichier(chemin), statut, reponse, boite
    adresse = ValeurJson(reponse, "uploadUrl")
    octets = LireOctets(chemin)
    debut = 0
    Do While debut < taille
        n = TAILLE_BLOC
        If debut + n > taille Then n = CLng(taille - debut)
        statut = Requete("PUT", adresse, Tranche(octets, debut, n), "application/octet-stream", False, reponse, "bytes " & CStr(debut) & "-" & CStr(debut + n - 1) & "/" & CStr(taille))
        If statut <> 200 And statut <> 201 Then Echec "Téléversement de " & NomFichier(chemin), statut, reponse, boite
        debut = debut + n
    Loop
End Sub

Private Function Requete(ByVal methode As String, ByVal url As String, ByVal corps As Variant, ByVal typeContenu As String, ByVal avecJeton As Boolean, ByRef reponse As String, ByVal plage As String) As Long
    Dim h As Object, jetonAcces As String, numero As Long, description As String
    If avecJeton Then jetonAcces = Jeton()
    If Len(mHttp) = 0 Then mHttp = "MSXML2.XMLHTTP.6.0"
    Set h = CreateObject(mHttp)
    h.Open methode, url, False
    If Len(typeContenu) > 0 Then h.setRequestHeader "Content-Type", typeContenu
    If avecJeton Then h.setRequestHeader "Authorization", "Bearer " & jetonAcces
    If Len(plage) > 0 Then h.setRequestHeader "Content-Range", plage
    On Error Resume Next
    h.send corps
    numero = Err.Number
    description = Err.Description
    On Error GoTo 0
    If numero <> 0 Then Err.Raise vbObjectError + 4001, "ModGraphMail", "Connexion impossible à " & Hote(url) & " (" & description & "). Vérifiez l'accès à Internet, le proxy ou le pare-feu."
    reponse = h.responseText
    Requete = h.Status
End Function

' Jeton d'application (OAuth 2.0, identifiants client), gardé jusqu'à 5 minutes avant son expiration.
Private Function Jeton() As String
    Dim reponse As String, statut As Long, duree As String, formulaire As String
    If Len(mJeton) > 0 And Now < mExpire Then
        Jeton = mJeton
        Exit Function
    End If
    If Len(mClientId) = 0 Then LireCoffre GRAPH_COFFRE, mClientId, mSecret
    If Len(mClientId) = 0 Or Len(mSecret) = 0 Then Err.Raise vbObjectError + 4002, "ModGraphMail", "Identifiant « " & GRAPH_COFFRE & " » absent du Gestionnaire d'identification Windows. Sur ce poste, lancez : cmdkey /generic:" & GRAPH_COFFRE & " /user:<ID d'application> /pass"
    formulaire = "client_id=" & EncoderUrl(mClientId) & "&client_secret=" & EncoderUrl(mSecret) & "&scope=" & EncoderUrl(Left(Graph(), InStr(9, Graph(), "/")) & ".default") & "&grant_type=client_credentials"
    If Len(mUrlConnexion) = 0 Then mUrlConnexion = "https://login.microsoftonline.com/"
    statut = Requete("POST", mUrlConnexion & EncoderUrl(GRAPH_LOCATAIRE) & "/oauth2/v2.0/token", formulaire, "application/x-www-form-urlencoded", False, reponse, "")
    If statut <> 200 Then Echec "Connexion à Microsoft 365", statut, reponse, ""
    mJeton = ValeurJson(reponse, "access_token")
    duree = ValeurJson(reponse, "expires_in")
    If Not IsNumeric(duree) Then duree = "3599"
    mExpire = DateAdd("s", CLng(duree) - 300, Now)
    Jeton = mJeton
End Function

Private Function Graph() As String
    If Len(mUrlGraph) = 0 Then mUrlGraph = "https://graph.microsoft.com/v1.0/"
    Graph = mUrlGraph
End Function

Private Function BoiteEnvoi(ByVal expediteur As String) As String
    If Len(Trim(expediteur)) > 0 Then BoiteEnvoi = Trim(expediteur) Else BoiteEnvoi = GRAPH_EXPEDITEUR
End Function

' Message d'erreur utile : la cause renvoyée par Microsoft, puis quoi vérifier.
Private Sub Echec(ByVal etape As String, ByVal statut As Long, ByVal reponse As String, ByVal boite As String)
    Dim cause As String, code As String, conseil As String
    cause = ValeurJson(reponse, "message")
    If Len(cause) = 0 Then cause = ValeurJson(reponse, "error_description")
    If InStr(cause, vbCr) > 0 Then cause = Left(cause, InStr(cause, vbCr) - 1)
    If InStr(cause, vbLf) > 0 Then cause = Left(cause, InStr(cause, vbLf) - 1)
    code = ValeurJson(reponse, "code")
    If Len(code) = 0 Then code = ValeurJson(reponse, "error")
    If InStr(cause, "AADSTS7000215") > 0 Or InStr(cause, "AADSTS7000222") > 0 Then
        conseil = "Le secret est faux ou expiré : demandez-en un nouveau à l'administrateur, puis relancez cmdkey."
    ElseIf InStr(cause, "AADSTS700016") > 0 Or InStr(cause, "AADSTS90002") > 0 Then
        conseil = "Vérifiez GRAPH_LOCATAIRE et l'ID d'application rangé par cmdkey."
    ElseIf statut = 401 Or statut = 403 Then
        conseil = "L'application n'a pas le droit d'envoyer depuis " & boite & " : l'administrateur peut le vérifier avec Test-ServicePrincipalAuthorization (guide, étape 4). Une modification de droits peut mettre jusqu'à 2 heures à s'appliquer."
    ElseIf statut = 404 Then
        conseil = "Boîte d'envoi introuvable : vérifiez l'adresse " & boite & "."
    ElseIf statut = 413 Then
        conseil = "Message trop volumineux pour Exchange Online (35 Mo par défaut)."
    ElseIf statut = 429 Or statut = 503 Then
        conseil = "Limite d'envoi atteinte : réessayez dans quelques minutes."
    End If
    Err.Raise vbObjectError + 4003, "ModGraphMail", etape & " refusé (HTTP " & statut & IIfTexte(Len(code) > 0, ", " & code, "") & ") : " & cause & IIfTexte(Len(conseil) > 0, " " & conseil, "")
End Sub

Private Function IIfTexte(ByVal condition As Boolean, ByVal siVrai As String, ByVal siFaux As String) As String
    If condition Then IIfTexte = siVrai Else IIfTexte = siFaux
End Function

' ---------------------------------------------------------------------------------------------------------
' JSON du message (Graph, ressource message).
Private Function MessageJson(ByVal destinataires As String, ByVal objet As String, ByVal corps As String, ByVal copie As String, ByVal copieCachee As String, ByVal pieces As String) As String
    Dim j As String, typeCorps As String
    If Left(LTrim(corps), 1) = "<" Then typeCorps = "HTML" Else typeCorps = "Text"
    j = "{""subject"":""" & Echapper(objet) & """,""body"":{""contentType"":""" & typeCorps & """,""content"":""" & Echapper(corps) & """}"
    j = j & ",""toRecipients"":" & Adresses(destinataires)
    If Len(Trim(copie)) > 0 Then j = j & ",""ccRecipients"":" & Adresses(copie)
    If Len(Trim(copieCachee)) > 0 Then j = j & ",""bccRecipients"":" & Adresses(copieCachee)
    MessageJson = j & pieces & "}"
End Function

' « a@x.fr; Nom Prénom <b@y.fr> » : séparateur point-virgule, comme dans Outlook.
Private Function Adresses(ByVal liste As String) As String
    Dim morceaux As Variant, i As Long, x As String, nom As String, adresse As String, j As String, p As Long
    morceaux = Split(liste, ";")
    For i = 0 To UBound(morceaux)
        x = Trim(morceaux(i))
        If Len(x) > 0 Then
            nom = ""
            adresse = x
            p = InStr(x, "<")
            If p > 0 And Right(x, 1) = ">" Then
                nom = Trim(Replace(Left(x, p - 1), """", ""))
                adresse = Trim(Mid(x, p + 1, Len(x) - p - 1))
            End If
            If Len(j) > 0 Then j = j & ","
            j = j & "{""emailAddress"":{""address"":""" & Echapper(adresse) & """"
            If Len(nom) > 0 Then j = j & ",""name"":""" & Echapper(nom) & """"
            j = j & "}}"
        End If
    Next
    If Len(j) = 0 Then Err.Raise vbObjectError + 4003, "ModGraphMail", "Aucun destinataire."
    Adresses = "[" & j & "]"
End Function

Private Function PiecesJson(ByVal fichiers As Variant) As String
    Dim i As Long, j As String
    For i = 0 To UBound(fichiers)
        If Len(j) > 0 Then j = j & ","
        j = j & PieceJson(fichiers(i))
    Next
    If Len(j) > 0 Then PiecesJson = ",""attachments"":[" & j & "]"
End Function

Private Function PieceJson(ByVal chemin As String) As String
    PieceJson = "{""@odata.type"":""#microsoft.graph.fileAttachment"",""name"":""" & Echapper(NomFichier(chemin)) & """,""contentBytes"":""" & Base64Fichier(chemin) & """}"
End Function

' Chemins séparés par « | » (interdit dans un nom de fichier Windows) ; vérifie qu'ils existent.
Private Function ListeFichiers(ByVal pieces As String) As Variant
    Dim morceaux As Variant, i As Long, x As String, valides As String
    ListeFichiers = Array()
    If Len(Trim(pieces)) = 0 Then Exit Function
    morceaux = Split(pieces, "|")
    For i = 0 To UBound(morceaux)
        x = Trim(morceaux(i))
        If Len(x) > 0 Then
            If Not FichierExiste(x) Then Err.Raise vbObjectError + 4003, "ModGraphMail", "Pièce jointe introuvable : " & x
            If TailleFichier(x) > TAILLE_MAX Then Err.Raise vbObjectError + 4003, "ModGraphMail", "Pièce jointe de plus de 150 Mo : " & x
            If Len(valides) > 0 Then valides = valides & "|"
            valides = valides & x
        End If
    Next
    If Len(valides) > 0 Then ListeFichiers = Split(valides, "|")
End Function

Private Function TailleTotale(ByVal fichiers As Variant) As Double
    Dim i As Long, t As Double
    For i = 0 To UBound(fichiers)
        t = t + TailleFichier(fichiers(i))
    Next
    TailleTotale = t
End Function

Private Function NomFichier(ByVal chemin As String) As String
    Dim p As Long
    p = InStrRev(Replace(chemin, "/", "\"), "\")
    NomFichier = Mid(chemin, p + 1)
End Function

Private Function Hote(ByVal url As String) As String
    Dim x As String
    x = Mid(url, InStr(url, "//") + 2)
    If InStr(x, "/") > 0 Then x = Left(x, InStr(x, "/") - 1)
    Hote = x
End Function

' Chaîne JSON en ASCII pur : les caractères accentués passent en \uXXXX (aucune question d'encodage).
Private Function Echapper(ByVal s As String) As String
    Dim i As Long, c As Long, r As String
    For i = 1 To Len(s)
        c = AscW(Mid(s, i, 1))
        If c < 0 Then c = c + 65536
        If c = 34 Then
            r = r & "\"""
        ElseIf c = 92 Then
            r = r & "\\"
        ElseIf c = 10 Then
            r = r & "\n"
        ElseIf c = 13 Then
            r = r & "\r"
        ElseIf c = 9 Then
            r = r & "\t"
        ElseIf c >= 32 And c <= 126 Then
            r = r & ChrW(c)
        Else
            r = r & "\u" & Right("000" & Hex(c), 4)
        End If
    Next
    Echapper = r
End Function

' Valeur d'une clé dans une réponse JSON de Graph (première occurrence ; texte ou nombre).
Private Function ValeurJson(ByVal json As String, ByVal cle As String) As String
    Dim p As Long, i As Long, c As String, r As String
    p = InStr(1, json, """" & cle & """", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(cle) + 2, json, ":")
    If p = 0 Then Exit Function
    i = p + 1
    Do While i <= Len(json)
        c = Mid(json, i, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        i = i + 1
    Loop
    If Mid(json, i, 1) = """" Then
        i = i + 1
        Do While i <= Len(json)
            c = Mid(json, i, 1)
            If c = """" Then Exit Do
            If c = "\" Then
                i = i + 1
                c = Mid(json, i, 1)
                If c = "n" Then
                    r = r & vbLf
                ElseIf c = "r" Then
                    r = r & vbCr
                ElseIf c = "t" Then
                    r = r & vbTab
                ElseIf c = "u" Then
                    r = r & ChrW(CLng("&H" & Mid(json, i + 1, 4)))
                    i = i + 4
                ElseIf c <> "b" And c <> "f" Then
                    r = r & c
                End If
            Else
                r = r & c
            End If
            i = i + 1
        Loop
    Else
        Do While i <= Len(json)
            c = Mid(json, i, 1)
            If c = "," Or c = "}" Or c = "]" Or c = " " Then Exit Do
            r = r & c
            i = i + 1
        Loop
    End If
    ValeurJson = r
End Function

' Encodage d'URL (RFC 3986) en UTF-8.
Private Function EncoderUrl(ByVal s As String) As String
    Dim i As Long, c As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 45 Or c = 46 Or c = 95 Or c = 126 Then
            r = r & ch
        ElseIf c < 128 Then
            r = r & "%" & Right("0" & Hex(c), 2)
        ElseIf c < 2048 Then
            r = r & "%" & Hex(192 + (c \ 64)) & "%" & Hex(128 + (c And 63))
        Else
            r = r & "%" & Hex(224 + (c \ 4096)) & "%" & Hex(128 + ((c \ 64) And 63)) & "%" & Hex(128 + (c And 63))
        End If
    Next
    EncoderUrl = r
End Function

'#binaire-debut
' ---------------------------------------------------------------------------------------------------------
' Fichiers et coffre Windows : instructions natives de VBA et VB6, MSXML pour le Base64 (livré avec Windows).
Private Function FichierExiste(ByVal chemin As String) As Boolean
    FichierExiste = CreateObject("Scripting.FileSystemObject").FileExists(chemin)
End Function

Private Function TailleFichier(ByVal chemin As String) As Double
    TailleFichier = FileLen(chemin)
End Function

Private Function LireOctets(ByVal chemin As String) As Variant
    Dim f As Integer, o() As Byte
    f = FreeFile
    Open chemin For Binary Access Read Shared As #f
    If LOF(f) > 0 Then
        ReDim o(0 To LOF(f) - 1)
        Get #f, , o
    End If
    Close #f
    LireOctets = o
End Function

Private Function Tranche(ByRef octets As Variant, ByVal debut As Double, ByVal n As Long) As Variant
    Dim r() As Byte, i As Long, d As Long
    ReDim r(0 To n - 1)
    d = CLng(debut)
    For i = 0 To n - 1
        r(i) = octets(d + i)
    Next
    Tranche = r
End Function

Private Function Base64Fichier(ByVal chemin As String) As String
    Dim e As Object
    If FileLen(chemin) = 0 Then Exit Function
    Set e = CreateObject("MSXML2.DOMDocument.6.0").createElement("b")
    e.DataType = "bin.base64"
    e.nodeTypedValue = LireOctets(chemin)
    Base64Fichier = Replace(Replace(e.Text, vbLf, ""), vbCr, "")
End Function

' Identifiant générique du Gestionnaire d'identification Windows : nom d'utilisateur = ID d'application,
' mot de passe = secret. Rien n'est écrit sur le disque par ce module.
Private Sub LireCoffre(ByVal cible As String, ByRef utilisateur As String, ByRef secret As String)
#If VBA7 Then
    Dim p As LongPtr, blob As LongPtr, nom As LongPtr
#Else
    Dim p As Long, blob As Long, nom As Long
#End If
    Dim taille As Long, n As Long
    If CredReadW(StrPtr(cible), 1, 0, p) = 0 Then Exit Sub
#If Win64 Then
    CopierMemoire taille, p + 32, 4
    CopierMemoire blob, p + 40, 8
    CopierMemoire nom, p + 72, 8
#Else
    CopierMemoire taille, p + 24, 4
    CopierMemoire blob, p + 28, 4
    CopierMemoire nom, p + 48, 4
#End If
    If taille > 0 And blob <> 0 Then
        secret = String(taille \ 2, vbNullChar)
        CopierMemoire ByVal StrPtr(secret), blob, taille
    End If
    If nom <> 0 Then
        n = lstrlenW(nom)
        utilisateur = String(n, vbNullChar)
        CopierMemoire ByVal StrPtr(utilisateur), nom, n * 2
    End If
    CredFree p
End Sub
'#binaire-fin
