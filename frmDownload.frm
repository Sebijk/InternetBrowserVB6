VERSION 5.00
Begin VB.Form frmDownload 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Datei herunterladen"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "frmDownload"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "IE-Download starten"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "http://"
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label 
      Caption         =   "Geben Sie hier an, welche Datei heruntergeladen werden soll:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Private Declare Function DoFileDownload Lib "shdocvw.dll" _
        (ByVal lpszFile As String) As Long

Private Sub Command1_Click()
  Dim Result&, URL$
    URL = StrConv(Text1.Text, vbUnicode)
    Call DoFileDownload(URL)
End Sub

