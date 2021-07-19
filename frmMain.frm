VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Internet-Browser"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   840
   ClientWidth     =   7545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3360
      Top             =   1560
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Oben ausrichten
      Height          =   540
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   953
      ButtonWidth     =   820
      ButtonHeight    =   794
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Neu"
            Object.ToolTipText     =   "Neu"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Öffnen"
            Object.ToolTipText     =   "Öffnen"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Speichern"
            Object.ToolTipText     =   "Speichern"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Back"
            Object.ToolTipText     =   "Zurück"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Forward"
            Object.ToolTipText     =   "Weiter"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Abbrechen"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Aktualisieren"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Home"
            Object.ToolTipText     =   "Startseite"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            Object.ToolTipText     =   "Suchen"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Drucken"
            Object.ToolTipText     =   "Drucken"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ShowPrint"
            Object.ToolTipText     =   "Seite einrichten"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Ausschneiden"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Kopieren"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Einfügen"
            ImageIndex      =   14
         EndProperty
      EndProperty
      Begin VB.PictureBox XPMenu1 
         Height          =   375
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   2
         Top             =   1680
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5010
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5768
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2698
            MinWidth        =   2698
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   926
            MinWidth        =   931
            Picture         =   "frmMain.frx":1CCA
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2118
            TextSave        =   "19.07.2021"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2118
            MinWidth        =   2118
            TextSave        =   "11:25"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3360
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2264
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2976
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3088
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":379A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":534A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":59C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":603E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Datei"
      Begin VB.Menu mnuNew 
         Caption         =   "&Neu"
         Begin VB.Menu mnuFileNew 
            Caption         =   "Neues &Fenster"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuNewIE 
            Caption         =   "&Internet Explorer"
         End
         Begin VB.Menu mnuNewMail 
            Caption         =   "&E-Mail..."
         End
         Begin VB.Menu mnuNewFileBrowser 
            Caption         =   "&Datei-Browser..."
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Ö&ffnen..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "S&chließen"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Speichern unter..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPageSetup 
         Caption         =   "&Seite einrichten..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Dru&cken..."
      End
      Begin VB.Menu mnuShowPrintPreview 
         Caption         =   "Druck&vorschau..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuWork 
         Caption         =   "Offlinebetrieb"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnuShowUndo 
         Caption         =   "&Rückgängig..."
      End
      Begin VB.Menu mnuShowRetry 
         Caption         =   "&Wiederholen..."
      End
      Begin VB.Menu mnuFileBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowCutText 
         Caption         =   "&Ausschneiden..."
      End
      Begin VB.Menu mnuShowCopyText 
         Caption         =   "&Kopieren..."
      End
      Begin VB.Menu mnuShowPaste 
         Caption         =   "&Einfügen..."
      End
      Begin VB.Menu mnuShowDelete 
         Caption         =   "&Löschen..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Alles makieren..."
      End
      Begin VB.Menu mnuClearSelectAll 
         Caption         =   "Alles demakieren..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Ansicht"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Symbolleiste"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status&leiste"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVollbild 
         Caption         =   "&Vollbild"
      End
   End
   Begin VB.Menu mnuViewUpdateFile 
      Caption         =   "Aktuelle &Seite"
      Begin VB.Menu mnuChanges 
         Caption         =   "&Änderungen zeigen..."
      End
      Begin VB.Menu mnuCookies 
         Caption         =   "&Cookies anzeigen..."
      End
      Begin VB.Menu mnuShowProperties 
         Caption         =   "&Eigenschaften..."
      End
      Begin VB.Menu mnuShowSendLink 
         Caption         =   "&Link senden..."
      End
      Begin VB.Menu mnuShowTextSource 
         Caption         =   "&Quelltext anzeigen..."
      End
      Begin VB.Menu mnuShowProtocol 
         Caption         =   "&Protokoll anzeigen..."
      End
      Begin VB.Menu mnuSiteOpenWith 
         Caption         =   "Seite &öffnen mit..."
         Begin VB.Menu mnuShowOpenUrlBrowser 
            Caption         =   "Sta&ndardbrowser..."
         End
         Begin VB.Menu mnuShowOpenUrlIE 
            Caption         =   "&Internet Explorer..."
         End
      End
      Begin VB.Menu mnuShowHostName 
         Caption         =   "&Server anzeigen..."
      End
      Begin VB.Menu mnuShowURL 
         Caption         =   "&URL anzeigen..."
      End
   End
   Begin VB.Menu mnuViewBrowser 
      Caption         =   "&Navigation"
      Begin VB.Menu mnuNavigateAt 
         Caption         =   "&Wechseln zu..."
      End
      Begin VB.Menu mnuFileBar9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "&Zurück..."
      End
      Begin VB.Menu mnuForward 
         Caption         =   "&Weiter..."
      End
      Begin VB.Menu mnuFileBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStop 
         Caption         =   "Abbre&chen"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Aktualisieren"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewHome 
         Caption         =   "Star&tseite..."
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "Suchen..."
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Werkzeuge"
      Begin VB.Menu mnuViewBrwInfo 
         Caption         =   "&Browserinformation anzeigen..."
      End
      Begin VB.Menu mnuViewFileBrowser 
         Caption         =   "&Datei-Browser..."
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "D&atei herunterladen..."
      End
      Begin VB.Menu mnuViewPicture 
         Caption         =   "Internet-&Bildbetrachter..."
      End
      Begin VB.Menu mnuGetUrlCache 
         Caption         =   "&Internet-Cache verwalten..."
      End
      Begin VB.Menu mnuShowInetBlock 
         Caption         =   "&Internet sperren/freigeben..."
      End
      Begin VB.Menu mnuShowIP 
         Caption         =   "IP-&Adresse auslesen..."
      End
      Begin VB.Menu mnuViewJava 
         Caption         =   "&Javainfos anzeigen..."
      End
      Begin VB.Menu mnuViewMPlayer 
         Caption         =   "&Media Player..."
      End
      Begin VB.Menu mnuViewMail 
         Caption         =   "&Nachricht schreiben..."
      End
      Begin VB.Menu mnuShowSource 
         Caption         =   "&Quelltextanzeiger..."
      End
      Begin VB.Menu mnuShowRSSConverter 
         Caption         =   "&RSS-Reader..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Optionen..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Fenster"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "Neues &Fenster"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Über&lappend"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "&Nebeneinander"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Ü&bereinander"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Symbole anordnen"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "Inf&o..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuShowInfoWindows 
         Caption         =   "Info über &Windows..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mbDontNavigateNow As Boolean
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Const EM_UNDO = &HC7
Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHAutoComplete Lib "shlwapi" (ByVal _
      hWnd As Long, ByVal dwFlags As Long) As Long



Private Const SHACF_DEFAULT = &H0
Private Const SHACF_FILESYSTEM = &H1
Private Const SHACF_URLHISTORY = &H2
Private Const SHACF_URLMRU = &H4
Private Const SHACF_USETAB = &H8
Private Const SHACF_FILESYS_ONLY = &H10
Private Const SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)

Private Const SHACF_AUTOSUGGEST_FORCE_ON = &H10000000
Private Const SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000
Private Const SHACF_AUTOAPPEND_FORCE_ON = &H40000000
Private Const SHACF_AUTOAPPEND_FORCE_OFF = &H80000000


' Benötigte API-Deklarationen
Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" ( _
  ByVal hWnd As Long, _
  ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As Long) As Long


Public Function GetTempDir2(Optional ByVal AddBackslash As Boolean) _
 As String
  Dim nTempDir As String

  On Error Resume Next
  nTempDir = Environ$("temp")
  If Len(nTempDir) = 0 Then
    nTempDir = Environ$("tmp")
  End If
  If Len(nTempDir) Then
    If AddBackslash Then
      GetTempDir2 = nTempDir & "\"
    Else
      GetTempDir2 = nTempDir
    End If
    End If
End Function

Public Function EnableAutoComplete(hWnd As Long, dwFlags As Long) _
         As Boolean

  On Error GoTo Err_AutoComplete

  SHAutoComplete hWnd, dwFlags
  EnableAutoComplete = True
  Exit Function

Err_AutoComplete:
  EnableAutoComplete = False
End Function



Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    tbToolBar.Refresh
    LoadNewDoc
End Sub

Function STATUSTEXT(swIndex As Long, swText As String)
sbStatusBar.Panels(swIndex).Text = swText
End Function

Function SSL_status(Status As Boolean)
If Status = False Then sbStatusBar.Panels(3).Visible = False Else sbStatusBar.Panels(3).Visible = True
End Function

Private Sub cboAddress_Click()
If ActiveForm Is Nothing Then
Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.StartingAddress = cboAddress.Text
    frmD.Caption = "Dokument " & lDocumentCount
    frmD.Show
    Else
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    ActiveForm.brwWebBrowser.Navigate cboAddress.Text
    End If
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub

Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.StartingAddress = "about:blank"
    frmD.Caption = "Dokument " & lDocumentCount
    frmD.Show
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuDownload_Click()
On Error Resume Next
    Dim DownloadForm As New frmDownload
    DownloadForm.Show
End Sub

Private Sub mnuGetUrlCache_Click()
On Error Resume Next
    Dim CacheForm As New frmGetUrlCache
    CacheForm.Show
End Sub


Private Sub mnuPageSetup_Click()
On Error Resume Next
If ActiveForm Is Nothing Then MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
ActiveForm.brwWebBrowser.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER
End Sub


Private Sub mnuShowDelete_Click()
On Error Resume Next
If ActiveForm Is Nothing Then MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
ActiveForm.brwWebBrowser.SelText = vbNullString
ActiveForm.brwWebBrowser.ExecWB OLECMDID_DELETE, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuNavigateAt_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
ActiveForm.brwWebBrowser.Navigate InputBox("Geben Sie eine Internet-Adresse ein:", , "")
End If
End Sub

Private Sub mnuNewFileBrowser_Click()
On Error Resume Next
    Static lDocumentCount As Long
    Dim frmfbrw As frmFileBrowser
    lDocumentCount = lDocumentCount + 1
    Set frmfbrw = New frmFileBrowser
    frmfbrw.Caption = "Dokument " & lDocumentCount
    StartingAddress = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    frmfbrw.Show
    ActiveForm.brwWebBrowser.Navigate StartingAddress
End Sub

Private Sub mnuNewIE_Click()
Set ie = New InternetExplorer
  ie.Navigate "about:blank"
  ie.StatusBar = True  'Statusleiste aktivieren
  ie.MenuBar = True    'Menü aktivieren
  ie.ToolBar = 1        'Toolbar aktivieren
  ie.FullScreen = False  'Vollbild deaktivieren
  ie.Visible = True     'Internet Explorer anzeigen
End Sub

Private Sub mnuNewMail_Click()
mnuViewMail_Click
End Sub

Private Sub mnuSelectAll_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuClearSelectAll_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_CLEARSELECTION, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowCopyText_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
Clipboard.SetText ActiveForm.brwWebBrowser.SelRTF
ActiveForm.brwWebBrowser.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowCutText_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
Clipboard.SetText ActiveForm.brwWebBrowser.SelRTF
ActiveForm.brwWebBrowser.SelText = vbNullString
ActiveForm.brwWebBrowser.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowHostName_Click()
On Error Resume Next
If ActiveForm Is Nothing Then
MsgBox "Es ist kein Fenster offen.", vbCritical
Else
StartingAddress = "javascript:alert('Der Servername heißt: ' + document.location.hostname)"
ActiveForm.brwWebBrowser.Navigate StartingAddress
End If
End Sub

Private Sub mnuShowInetBlock_Click()
FIBForm.Show
End Sub

Private Sub mnuShowInfoWindows_Click()
Call ShellAbout(Me.hWnd, Me.Caption, "Lizenziert unter der GNU General Public License.", Me.Icon)
End Sub

Private Sub mnuShowIP_Click()
Dim IPForm As New frmIPForm
IPForm.Show
End Sub

Private Sub mnuShowPaste_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.SelRTF = Clipboard.GetText
ActiveForm.brwWebBrowser.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowPrintPreview_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowProperties_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowRSSConverter_Click()
Set RSS2HTML = CreateObject("RSS2HTMLScoutLib.RSS2HTMLScout")

rssfile = InputBox("Bitte geben sie die URL zu der RSS-Datei an!", "RSS-Reader", "")
If rssfile = "" Then Exit Sub

 RSS2HTML.ItemsPerFeed = 20 ' display only 20 latest items
 ' ##### we can add more than one RSS feed #########
 RSS2HTML.AddFeed rssfile, 300  ' update every 300 minutes (5 hours)
 RSS2HTML.Execute
 RSS2HTML.SaveOutputToFile GetTempDir2 & "\rss.html" ' save output to HTML file
 Set RSS2HTML = Nothing
 Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.StartingAddress = GetTempDir2 & "\rss.html"
    frmD.Caption = "Dokument " & lDocumentCount
    frmD.Show
End Sub

Private Sub mnuShowSendLink_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
ActiveForm.brwWebBrowser.Navigate "mailto:?subject=E-Mail schreiben an: " & ActiveForm.brwWebBrowser.LocationName & "&body=Habe eine Tolle Seite gefunden. Guck mal unter " & ActiveForm.brwWebBrowser.LocationURL
End If
End Sub

Private Sub mnuShowSource_Click()
Dim frmtxt As New frmQuellCode
frmtxt.StartingAddress = "about:blank"
frmtxt.Show
End Sub

Private Sub mnuShowTextSource_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
Dim pURL As String
pURL = ActiveForm.brwWebBrowser.LocationURL
DoEvents
Dim frmtxt As New frmQuellCode
frmtxt.StartingAddress = pURL
frmtxt.Show
End If
End Sub

Private Sub mnuShowOpenUrlBrowser_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
Dim pURL As String
pURL = ActiveForm.cboAddress.Text
DoEvents
Call ShellExecute(Me.hWnd, "open", pURL, "", "", 1)
End If
End Sub

Private Sub mnuShowOpenUrlIE_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
Dim pURL As String
pURL = ActiveForm.cboAddress.Text
DoEvents
Set ie = New InternetExplorer
  ie.Navigate pURL
  ie.StatusBar = True  'Statusleiste aktivieren
  ie.MenuBar = True    'Menü aktivieren
  ie.ToolBar = 1        'Toolbar aktivieren
  ie.FullScreen = False  'Vollbild deaktivieren
  ie.Visible = True     'Internet Explorer anzeigen
End If
End Sub

Private Sub mnuShowUndo_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_UNDO, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuShowRetry_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_REDO, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub




Private Sub mnuViewBrwInfo_Click()
On Error Resume Next
    If ActiveForm Is Nothing Then LoadNewDoc
    StartingAddress = "javascript:alert('Sie verwenden ' + navigator.appName + ' \nCodename: ' + navigator.userAgent + '. \nSie benutzen ein ' + navigator.platform + '-Betriebssystem.')"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
End Sub

Private Sub mnuViewJava_Click()
On Error Resume Next
    If ActiveForm Is Nothing Then LoadNewDoc
    StartingAddress = "javascript:alert('Java ist ' +navigator.javaEnabled());"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
End Sub

Private Sub mnuSave_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
On Error Resume Next
ActiveForm.brwWebBrowser.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End If
End Sub

Private Sub mnuViewMail_Click()
On Error Resume Next
    If ActiveForm Is Nothing Then LoadNewDoc
    StartingAddress = "mailto:"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
End Sub



Private Sub mnuViewFileBrowser_Click()
On Error Resume Next
    Static lDocumentCount As Long
    Dim frmfbrw As frmFileBrowser
    lDocumentCount = lDocumentCount + 1
    Set frmfbrw = New frmFileBrowser
    frmfbrw.Caption = "Document " & lDocumentCount
    StartingAddress = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
    frmfbrw.Show
    ActiveForm.brwWebBrowser.Navigate StartingAddress
End Sub


Private Sub mnuViewMPlayer_Click()
Dim FMediaPlayer As New frmMedia
FMediaPlayer.Show
End Sub

Private Sub mnuViewPicture_Click()
Dim InetPicForm As New frmInetPicture
InetPicForm.Show
End Sub

Private Sub mnuViewStop_Click()
    If ActiveForm Is Nothing Then
    MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
    Else
    ActiveForm.timTimer.Enabled = False
            ActiveForm.brwWebBrowser.Stop
            ActiveForm.Caption = ActiveForm.brwWebBrowser.LocationName
            End If
End Sub

Private Sub mnuVollbild_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
    If mnuVollbild.Checked = False Then
    mnuVollbild.Checked = True
    Else
    mnuVollbild.Checked = False
End If
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewWebBrowser_Click()
    Dim frmB As New frmDocument
    frmB.StartingAddress = "about:blank"
    frmB.Show
End Sub

Private Sub mnuViewOptions_Click()
   On Error Resume Next
    Dim OptionForm As New frmOption
    OptionForm.Show
End Sub

Private Sub mnuViewRefresh_Click()
  If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
    ActiveForm.brwWebBrowser.Refresh
End If
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.brwWebBrowser.SelRTF = Clipboard.GetText

End Sub


Private Sub mnuFileExit_Click()
    'Formular entladen
    Unload Me

End Sub


Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then
   MsgBox "Es ist kein Fenster da, das gedruckt werden soll.", vbCritical
   Else
    StartingAddress = "javascript:window.print()"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
    End If
End Sub
Private Sub mnuCookies_Click()
    On Error Resume Next
     If ActiveForm Is Nothing Then
   MsgBox "Es ist kein Fenster offen.", vbCritical
   Else
    StartingAddress = "javascript:alert('Cookie: ' + document.cookie)"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
    End If
End Sub
Private Sub mnuChanges_Click()
    On Error Resume Next
     If ActiveForm Is Nothing Then
   MsgBox "Es ist kein Fenster offen.", vbCritical
   Else
    StartingAddress = "javascript:alert('Diese Seite wurde zuletzt am ' + document.lastModified + ' aktualisiert.')"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
    End If
End Sub

Private Sub mnuShowProtocol_Click()
    On Error Resume Next
     If ActiveForm Is Nothing Then
   MsgBox "Es ist kein Fenster offen.", vbCritical
   Else
    StartingAddress = "javascript:alert('Diese Seite benutzt ' + window.document.protocol)"
    ActiveForm.brwWebBrowser.Navigate StartingAddress
    End If
End Sub
Private Sub mnuShowURL_Click()
    On Error Resume Next
     If ActiveForm Is Nothing Then
   MsgBox "Es ist kein Fenster offen.", vbCritical
   Else
    MsgBox "Die Adresse dieser Seite heisst: " + ActiveForm.brwWebBrowser.LocationURL, vbInformation
    End If
End Sub

Private Sub mnuFileClose_Click()
   If ActiveForm Is Nothing Then
   MsgBox "Es wurden bereits alle Fenster geschlossen.", vbCritical
   Else
   Unload ActiveForm
   End If
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String

    LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Öffnen"
        .CancelError = False
        'Zu erledigen: Festlegen der Flags und Attribute des Standarddialog-Steuerelements
        .Filter = "Alle Dateien (*.*)|*.*"
        .ShowOpen
        If Len(.filename) = 0 Then
        Exit Sub
        End If
        If .CancelError = True Then
        Exit Sub
        End If
        sFile = .filename
    End With
    ActiveForm.brwWebBrowser.Navigate sFile
    ActiveForm.Caption = sFile

End Sub
Private Sub mnuBack_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
ActiveForm.brwWebBrowser.GoBack
End If
End Sub
Private Sub mnuForward_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
ActiveForm.brwWebBrowser.GoForward
End If
End Sub

Private Sub mnuViewHome_Click()

    If ActiveForm Is Nothing Then LoadNewDoc
    
    ActiveForm.brwWebBrowser.GoHome

End Sub
Private Sub mnuViewSearch_Click()

    If ActiveForm Is Nothing Then LoadNewDoc
    
    ActiveForm.brwWebBrowser.GoSearch

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub


Private Sub mnuWork_Click()
If ActiveForm Is Nothing Then
MsgBox "Bitte öffnen Sie ein neues Fenster!", vbExclamation
Exit Sub
Else
If mnuWork.Checked = False Then
mnuWork.Checked = True
ActiveForm.brwWebBrowser.Offline = True
fMainForm.STATUSTEXT 1, "Offline"
Me.Caption = App.Title & " - [Offlinebetrieb]"
Else
mnuWork.Checked = False
ActiveForm.brwWebBrowser.Offline = False
Me.Caption = App.Title
fMainForm.STATUSTEXT 1, "Online"
End If
End If
End Sub




Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Neu"
            LoadNewDoc
        Case "Öffnen"
            mnuFileOpen_Click
        Case "Speichern"
            mnuSave_Click
        Case "Drucken"
            mnuFilePrint_Click
        Case "ShowPrint"
            mnuShowPrintPreview_Click
        Case "Copy"
            mnuShowCopyText_Click
        Case "Cut"
            mnuShowCutText_Click
        Case "Paste"
            mnuShowPaste_Click
        Case "Back"
            mnuBack_Click
        Case "Forward"
            mnuForward_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Home"
            mnuViewHome_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Stop"
            mnuViewStop_Click
    End Select
End Sub

