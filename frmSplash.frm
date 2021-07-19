VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   4320
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9000
      Begin VB.Label Label1 
         Caption         =   "Sebijk's Internet-Browser wird geladen..."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3720
         Width           =   6495
      End
      Begin VB.Image imgLogo 
         Height          =   3135
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Lizenziert unter der GNU General Public License"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   1
         Top             =   2520
         Width           =   3735
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Produkt"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3360
         TabIndex        =   3
         Top             =   1200
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

