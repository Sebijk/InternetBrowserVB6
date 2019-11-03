VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmInetPicture 
   Caption         =   "Internet-Bildbetrachter"
   ClientHeight    =   5250
   ClientLeft      =   3060
   ClientTop       =   3450
   ClientWidth     =   7440
   Icon            =   "frmInetPicture.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmInetPicture.frx":74F2
   ScaleHeight     =   5250
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox brwWebBrowser 
      BackColor       =   &H80000009&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   3
      Top             =   720
      Width           =   3015
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
Attribute VB_Name = "frmInetPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StartingAddress As String
Dim mbDontNavigateNow As Boolean

Option Explicit

Private textbuffer As String
Private binarybuffer() As Byte
Private mode As Integer

Public Sub GetRESOURCE(ByVal WWW_Adresse As String)
  On Error GoTo GetRESOURCEERR
  
  Dim filename As String
  Dim fnum As Long
  
  filename = fMainForm.GetTempDir2
  If Not Right(filename, 1) = "\" Then filename = filename & "\"
  filename = filename & LastPath(WWW_Adresse)
    
  MousePointer = vbHourglass
  brwWebBrowser.Picture = LoadPicture()
  DoEvents

  mode = 2

  Inet1.Execute WWW_Adresse
  
  While Inet1.StillExecuting
    DoEvents
  Wend
  
  If UBound(binarybuffer) - LBound(binarybuffer) + 1 > 1 Then
    fnum = FreeFile
    Open filename For Binary Access Write As #fnum
      Put #fnum, , binarybuffer()
    Close #fnum
    
    ReDim binarybuffer(0)
    
    brwWebBrowser.Picture = LoadPicture(filename)
  End If
    
  MousePointer = vbDefault
  Exit Sub
  
GetRESOURCEERR:
  MsgBox (Err.Description)
  Resume Next
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.Show
    Form_Resize


   cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15

If Len(StartingAddress) > 0 Then
        cboAddress.Text = StartingAddress
        cboAddress.AddItem cboAddress.Text
        'versuche auf Startadresse zu positionieren
        cboAddress_Click
    End If


End Sub


Private Sub cboAddress_KeyPress(KeyAscii As Integer)
  On Error Resume Next
 If KeyAscii = vbKeyReturn Then
 cboAddress_Click
End If
End Sub

Private Sub cboAddress_Click()
    
    MousePointer = vbHourglass
    DoEvents
    
    Call GetRESOURCE(cboAddress.Text)

    MousePointer = vbDefault
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

Private Sub Inet1_StateChanged(ByVal State As Integer)
  If State = icResponseCompleted Then
    If mode = 1 Then
        ' Textdaten abrufen
        Dim data As String
        
        textbuffer = ""
        data = "dummy"
        
        While Len(data) > 0
            data = Inet1.GetChunk(1024, icString)
            textbuffer = textbuffer & data
        Wend
    ElseIf mode = 2 Then
        ' Binärdaten abrufen
        Dim binbuf() As Byte
        Dim s1 As String, s2 As String
        binbuf = Inet1.GetChunk(1024, icByteArray)
        
        ' Überprüfung auf Dimensionierung
        ' wahrscheinlich Inet-Steuerelement spezifisch
        If LBound(binbuf) <= UBound(binbuf) Then
            binarybuffer = binbuf
            
            While LBound(binbuf) <= UBound(binbuf)
                binbuf = Inet1.GetChunk(1024, icByteArray)
                
                If LBound(binbuf) <= UBound(binbuf) Then
                    s1 = StrConv(binarybuffer, vbUnicode)
                    s2 = StrConv(binbuf, vbUnicode)
                    binarybuffer = StrConv(s1 & s2, vbFromUnicode)
                End If
            Wend
        Else
            ReDim binarybuffer(0)
        End If
    Else
        Call MsgBox("Ungültiger Modus im Inet1_StateChanged!", _
                    vbExclamation + vbOKOnly, App.Title)
    End If
  End If

  If Inet1.ResponseCode <> 0 Then _
     MsgBox (Inet1.ResponseCode & " : " & Inet1.ResponseInfo)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Inet1.Cancel
End Sub

Private Function LastPath(ByVal Path As String) As String
    Dim aa As String, BB As String
    Dim x As Long
    
    For x = Len(Path) To 1 Step -1
        aa = Mid$(Path, x, 1)
        If aa = "/" Or aa = "\" Then
            Exit For
        Else
            BB = aa & BB
        End If
    Next x
    LastPath = BB
End Function

