VERSION 5.00
Begin VB.Form frmInetBlock 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Sebijks Internet-Browser"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6840
   Icon            =   "frmInetBlock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1185
   ScaleWidth      =   6840
   Begin VB.CommandButton cmdHide 
      Caption         =   "Verstecken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Schließen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   480
   End
   Begin VB.CommandButton But 
      Caption         =   "Internet stoppen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   2
      Tag             =   "F"
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   -120
      Width           =   6615
      Begin VB.Label Label1 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Internet ist geöffnet !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmInetBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub But_Click()

    If Timer1.Enabled Then
        Label1.Caption = "Internet ist geöffnet !"
        Label1.ForeColor = &HC000&
        But.Caption = "Internet sperren"
        Me.Caption = "Sebijks Internet-Browser - [Internet ist geöffnet]"
        fMainForm.Caption = "Sebijks Internet-Browser"
        Timer1.Enabled = False
    Else
        Me.Caption = "Sebijks Internet-Browser - [Internet ist gesperrt]"
        fMainForm.Caption = "Sebijks Internet-Browser - [Internet gesperrt]"
        Label1.Caption = "Internet ist gesperrt !"
        Label1.ForeColor = vbRed
        But.Caption = "Internet freigeben"
        
        Timer1.Enabled = True
        Call Timer1_Timer
    End If
    
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdHide_Click()
FIBForm.Hide
End Sub

Private Sub Form_Load()
       
    Timer1.Interval = 100
    Timer1.Enabled = False
    
End Sub
Private Sub Timer1_Timer()
    
    Static Work As Boolean
    
    If Work Then Exit Sub
    Work = True
    
    Call RefreshStack
    Call EnumEntries
    
    DoEvents
    Work = False
    
End Sub
