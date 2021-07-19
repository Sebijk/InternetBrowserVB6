VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocument 
   ClientHeight    =   6375
   ClientLeft      =   3060
   ClientTop       =   3450
   ClientWidth     =   8565
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmBrowser.frx":058A
   ScaleHeight     =   6375
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picAddress 
      Align           =   1  'Oben ausrichten
      BorderStyle     =   0  'Kein
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   8565
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8565
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   300
         Width           =   5355
      End
      Begin VB.Label lblAddress 
         Caption         =   " &Adresse:"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Tag             =   "&Adresse:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   6180
      Top             =   1500
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1200
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":08CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":16F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2514
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2C26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   2280
      Left            =   50
      TabIndex        =   0
      Top             =   765
      Width           =   2520
      ExtentX         =   4445
      ExtentY         =   4022
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   159
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3338
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":626A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":919C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":C0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":F000
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":11F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":14E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":17D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1ACC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1DBFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":20B2C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Private Sub Form_Load()
    'On Error Resume Next
    Me.Show
    
    Form_Resize
    fMainForm.EnableAutoComplete cboAddress.hWnd, 0&
    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 20
    If Len(StartingAddress) > 0 Then
        'versuche auf Startadresse zu positionieren
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If


End Sub

Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub

Private Sub brwWebBrowser_NavigateError(ByVal pDisp As Object, _
  URL As Variant, Frame As Variant, StatusCode As Variant, _
  Cancel As Boolean)
  Dim sHTML As String
  
  'URLs und Statuscodes überpüfen
  If StatusCode = "200" Then Exit Sub
  If StatusCode = "301" Then Exit Sub
  If StatusCode = "302" Then Exit Sub
  If StatusCode = "403" Then Exit Sub
  
  frage = MsgBox("Die Webseite könnte möglicherweise Fehler enthalten! Wollen Sie trotzdem fortfahren?", vbExclamation + vbYesNo)
  If frage = vbYes Then Exit Sub
  If frage = vbNo Then
  ' Eigene Fehlerseite erstellen
  sHTML = "about:" & _
    "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & _
    "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">" & _
    "<title>Sebijks Internet-Browser Fehler</title>" & _
    "</head>" & _
    "<body><basefont size=2 face=Tahoma>" & _
    "<table width=""730"" cellpadding=""0"" cellspacing=""0"" border=""0""><tr>" & _
    "<td width=""60"" align=""left"" valign=""top"" rowspan=""2"">" & _
    "</td>" & _
    "<td valign=""middle"" align=""left"" width=""*"" bgcolor=""#E8EAEF"">" & _
    "<h1 id=""mainTitle""><b>Fehler: Die Seite kann nicht angezeigt werden</b></h1>" & _
    "<font face=""Arial"" size=""2"">Sebijks Internet-Browser kann die Seite nicht anzeigen, weil sie entweder nicht gefunden wurde oder die Webseite Fehler enth&auml;lt!</font>" & _
    "</td></tr><tr><td class=""errorCodeAndDivider"" align=""right"" bgcolor=""#E8EAEF"">&nbsp;" & _
    "<div class=""divider""></div></td></tr><tr><td>&nbsp;</td>" & _
    "<td valign=""top"" align=""left"" bgcolor=""#E8EAEF"">" & _
    "<h3 id=""likelyCauses"">L&ouml;sungen:</h3>" & _
    "<ul><li id=""causeNotConnected"">Wenn Sie auch keine andere Website aufrufen k&ouml;nnen, &uuml;berpr&uuml;fen Sie bitte die <br>Netzwerk-/Internetverbindung.</li>" & _
    "<li id=""causeSiteProblem"">Die Webseite ist nicht erreichbar, versuchen sie es sp&auml;ter erneuert.</li>" & _
    "<li id=""causeErrorInAddress"">Bitte &uuml;berpr&uuml;fen Sie die Adresse auf Tippfehler, wie " & _
    "<b>ww.beispiel.de</b> statt <b>www.beispiel.de</b></li></ul></td></tr>" & _
    "<tr><td>&nbsp;</td><td align=""left"" valign=""middle"" bgcolor=""#E8EAEF""><h4>" & _
    "<a href=""javascript:history.back(1)""><font color=""#000FFF"">Zur&uuml;ck zur vorherigen Seite</font></a><p>" & _
    "<a href=""http://www.sebijk.de""><font color=""#000FFF"">Zur Home of the Sebijk.de</font></a></p>" & _
    "<p>Fehlercode: " & CStr(StatusCode) & "<p>Sebijks Internet - Browser</p></h4>" & _
    "<p>URL : <a href=" & URL & "><font color=""#000FFF"">" & URL & "</font></a></td></tr></table></basefont></body>"


  ' Jetzt Fehlerseite anzeigen
  With brwWebBrowser
    .Silent = True
    .Navigate sHTML
  End With
  cboAddress.Text = ""
  End If
End Sub

Private Sub brwwebbrowser_BeforeNavigate2(ByVal pDisp As Object, ByRef URL As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If Media(URL) = True Then Cancel = True: DoEvents: Exit Sub
End Sub
Function Media(ByVal sUrl As String) As Boolean
    Dim FMediaForm As New frmMedia
    If Right(LCase(sUrl), 4) = ".wav" Then
        Media = True
        DoEvents
        FMediaForm.StartingAddress = sUrl
        FMediaForm.Show
    ElseIf Right(LCase(sUrl), 4) = ".avi" Then
        Media = True
        DoEvents
        FMediaForm.StartingAddress = sUrl
        FMediaForm.Show
    ElseIf Right(LCase(sUrl), 4) = ".mp3" Then
        Media = True
        DoEvents
        FMediaForm.StartingAddress = sUrl
        FMediaForm.Show
    ElseIf Right(LCase(sUrl), 4) = ".wma" Then
        Media = True
        DoEvents
        FMediaForm.StartingAddress = sUrl
        FMediaForm.Show
    ElseIf Right(LCase(sUrl), 4) = ".wmv" Then
        Media = True
        DoEvents
        FMediaForm.StartingAddress = sUrl
        FMediaForm.Show
    End If
End Function
Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = brwWebBrowser.LocationName
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    mbDontNavigateNow = True
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    mbDontNavigateNow = False
End Sub


Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    brwWebBrowser.Navigate cboAddress.Text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  If WindowState = vbMinimized Then Exit Sub

  cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub


Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "In Bearbeitung..."
    End If
End Sub

' Zustand der Statusleiste
Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)
  fMainForm.STATUSTEXT 1, Text
End Sub
Private Sub brwWebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
  frage = MsgBox("Die Webseite will ein neues Fenster/Popup öffnen! Möchten Sie das zulassen?", vbExclamation + vbYesNo)
  If frage = vbNo Then Cancel = True
  If frage = vbYes Then
  ' Neue Instanz der Popup-Form erstellen
  Dim oForm As frmDocument
  Set oForm = New frmDocument

  With oForm
    ' als Browser registrieren
    .brwWebBrowser.RegisterAsBrowser = True

    ' WebBrowser-Object zuweisen
    Set ppDisp = .brwWebBrowser.Object

    ' Form anzeigen
    .Show
  End With
  End If
End Sub

Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
If Progress = 0 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(11).Picture
If Progress > 200 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(1).Picture
If Progress > 300 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(2).Picture
If Progress > 400 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(3).Picture
If Progress > 500 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(4).Picture
If Progress > 600 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(5).Picture
If Progress > 700 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(6).Picture
If Progress > 800 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(7).Picture
If Progress > 9000 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(8).Picture
If Progress > 1000 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(9).Picture
If Progress > 2000 Then fMainForm.sbStatusBar.Panels(2).Picture = ImageList1.ListImages(10).Picture
End Sub

Private Sub brwWebBrowser_SetSecureLockIcon(ByVal SecureLockIcon As Long)
If SecureLockIcon = 0 Then
    fMainForm.SSL_status False
Else
    fMainForm.SSL_status True
End If
End Sub

Private Sub brwWebBrowser_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
Cancel = True
Unload Me
End Sub
