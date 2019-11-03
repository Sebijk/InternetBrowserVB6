VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmQuellCode 
   Caption         =   "Quellcodeanzeiger"
   ClientHeight    =   5250
   ClientLeft      =   3060
   ClientTop       =   3450
   ClientWidth     =   7440
   Icon            =   "frmQuellCode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox brwWebBrowser 
      Height          =   1575
      Left            =   50
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   840
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6000
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Oben ausrichten
      BorderStyle     =   0  'Kein
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7440
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7440
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   300
         Width           =   5355
      End
      Begin VB.Label lblAddress 
         Caption         =   " &Adresse:"
         Height          =   255
         Left            =   0
         TabIndex        =   0
         Tag             =   "&Adresse:"
         Top             =   60
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmQuellCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim strUrl As String

Option Explicit


Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    Form_Resize


   cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

If Len(StartingAddress) > 0 Then

        cboAddress.AddItem cboAddress.Text
        'versuche auf Startadresse zu positionieren
        MousePointer = vbHourglass
        DoEvents
        DoEvents
        'brwWebBrowser.Text = GetHTML(StartingAddress)
        MousePointer = vbDefault
    End If


End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
  On Error Resume Next
 If KeyAscii = vbKeyReturn Then
 cboAddress_Click
End If
End Sub

Private Sub cboAddress_Click()

    On Error GoTo errorhandler
    strUrl = cboAddress.Text
    'Feld muss mindestens 11 Zeichen ("http://www.") enthalten
    If Len(strUrl) > 11 Then
        'Html-Dokument in das Textfeld kopieren
        brwWebBrowser.Text = Inet1.OpenURL(strUrl)
    Else
        MsgBox "Geben Sie einen gültigen Dokumentennamen in das Feld URL ein"
    End If
    Exit Sub
errorhandler:
    MsgBox "Fehler beim Öffnen der URL", , Err.Description
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  If WindowState = vbMinimized Then Exit Sub
  'cboAddress.Width = ScaleWidth - cboAdress.Left - 100
  cboAddress.Width = Me.ScaleWidth - 100
  brwWebBrowser.Width = Me.ScaleWidth - 100
  'rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
  'brwWebBrowser.RightMargin = brwWebBrowser.Width - 400
  brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Inet1.Cancel
End Sub


