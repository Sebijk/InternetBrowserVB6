VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmMedia 
   Caption         =   "Media Player"
   ClientHeight    =   5325
   ClientLeft      =   105
   ClientTop       =   345
   ClientWidth     =   5925
   Icon            =   "frmMedia.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   5925
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5400
      Top             =   1680
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  'Oben ausrichten
      BorderStyle     =   0  'Kein
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   5925
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5925
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   5355
      End
      Begin VB.Label lblAddress 
         Caption         =   " &Adresse:"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Tag             =   "&Adresse:"
         Top             =   60
         Width           =   3075
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   4575
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5895
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String

Sub Form_Load()
Form_Resize
cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 20
If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'versuche auf Startadresse zu positionieren
        timTimer.Enabled = True
        MediaPlayer1.filename = StartingAddress
        MediaPlayer1.Play
    End If
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If WindowState = vbMinimized Then Exit Sub
  cboAddress.Width = Me.ScaleWidth - 100
    MediaPlayer1.Width = Me.ScaleWidth - 100
    MediaPlayer1.Height = Me.ScaleHeight - (picAddress.Top + picAddress.Height) - 100
End Sub

Private Sub MediaPlayer1_ToolTipText(ByVal Text As String)
  fMainForm.STATUSTEXT 1, Text
End Sub

Private Sub cboAddress_Click()
    If mbDontNavigateNow Then Exit Sub
    timTimer.Enabled = True
    MediaPlayer1.filename = cboAddress.Text
    MediaPlayer1.Play
    cboAddress.AddItem cboAddress.Text
End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    'On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        cboAddress_Click
    End If
End Sub
