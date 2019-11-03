VERSION 5.00
Begin VB.Form frmFavoriten 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Favoriten"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFavoriten 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmFavoriten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const R_FAVORITEN = &H6
Private Const NOERROR = 0
Private Sub Form_Load()
Kill FileFavs.Path & "\" & lstFavoriten.Text & ".url"
DoEvents
FileFavs.Refresh
LoadFavoriten
FileFavs.ListIndex = lstFavoriten.ListIndex
End Sub

Private Function LoadFavoriten()
lstFavoriten.Clear
DoEvents
Dim favs As Variant
For ii = 0 To FileFavs.ListCount - 1
favs = Split(FileFavs.List(ii), ".url")
lstFavoriten.AddItem favs(0)
DoEvents
Next ii
End Function

Private Sub lstFavoriten_DblClick()
Dim strURL As String
strURL = GetFavorite("InternetShortcut", "URL", FileFavs.Path & "\" & lstFavoriten.Text & ".url")
web.Navigate strURL
End Sub
