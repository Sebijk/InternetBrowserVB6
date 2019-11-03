VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmDoc3 
   Caption         =   "frmDocument"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   5775
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   3795
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5280
      Top             =   1080
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5400
      ExtentX         =   9513
      ExtentY         =   6586
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
      Location        =   ""
   End
   Begin VB.Label lblAddress 
      Caption         =   "&Adresse:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "&Adresse:"
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "frmDoc3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean
Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    tbToolBar.Refresh
    Form_Resize


    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15


    If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'versuche auf Startadresse zu positionieren
        timTimer.Enabled = True
        brwWebBrowser.Navigate StartingAddress
    End If


End Sub



Private Sub brwWebBrowser_DownloadComplete()
    On Error Resume Next
    Me.Caption = brwWebBrowser.LocationName
End Sub


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
    cboAddress.Width = Me.ScaleWidth - 100
    brwWebBrowser.Width = Me.ScaleWidth - 100
    brwWebBrowser.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub


Private Sub timTimer_Timer()
    If brwWebBrowser.Busy = False Then
        timTimer.Enabled = False
        Me.Caption = brwWebBrowser.LocationName
    Else
        Me.Caption = "Bearbeiten..."
    End If
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
      

    timTimer.Enabled = True
      

    Select Case Button.Key
        Case "Back"
            brwWebBrowser.GoBack
        Case "Forward"
            brwWebBrowser.GoForward
        Case "Refresh"
            brwWebBrowser.Refresh
        Case "Home"
            brwWebBrowser.GoHome
        Case "Search"
            brwWebBrowser.GoSearch
        Case "Stop"
            timTimer.Enabled = False
            brwWebBrowser.Stop
            Me.Caption = brwWebBrowser.LocationName
    End Select


End Sub

