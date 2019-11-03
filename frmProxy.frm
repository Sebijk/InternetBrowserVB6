VERSION 5.00
Begin VB.Form frmProxy 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Proxy-Server konfigurieren"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Beenden 
      Caption         =   "Beenden"
      Height          =   285
      Left            =   800
      TabIndex        =   13
      Top             =   1365
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Wiederherstellen"
      Height          =   285
      Left            =   5280
      TabIndex        =   12
      ToolTipText     =   " Resore settings before start of program "
      Top             =   1365
      Width           =   1410
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Leeren"
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Top             =   1365
      Width           =   1050
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Übernehmen"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   1365
      Width           =   1170
   End
   Begin VB.OptionButton OptProxyEnable 
      Caption         =   "deaktivieren"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   8
      Top             =   855
      Value           =   -1  'True
      Width           =   1410
   End
   Begin VB.OptionButton OptProxyEnable 
      Caption         =   "aktivieren"
      Height          =   195
      Index           =   0
      Left            =   1365
      TabIndex        =   7
      Top             =   855
      Width           =   1275
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Lesen"
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1365
      Width           =   1050
   End
   Begin VB.TextBox TxtPort 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   6975
      TabIndex        =   2
      Top             =   420
      Width           =   600
   End
   Begin VB.TextBox TxtProxIP 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   5265
      TabIndex        =   1
      Top             =   420
      Width           =   1590
   End
   Begin VB.TextBox txtProxy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   135
      TabIndex        =   0
      Top             =   420
      Width           =   4995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proxystatus:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   855
      Width           =   1050
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Proxyname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   210
      Width           =   930
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6975
      TabIndex        =   4
      Top             =   210
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Proxy IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5310
      TabIndex        =   3
      Top             =   210
      Width           =   720
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dieser Source stammt von http://www.ActiveVB.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Sollten Sie Fehler entdecken oder Fragen haben, dann
'mailen Sie mir bitte unter: Reinecke@ActiveVB.de
'Ansonsten viel Spaß und Erfolg mit diesem Source !
'**************************************************************

' Autor: K. Langbein Klaus@ActiveVB.de

' Beschreibung: Neben dem Internet Explorer verwenden auch andere
' MS-Komponenten, z.B. das MSInet.ocx, die Einstellungen des IE.
' Mit Hilfe dieses Programms koennen die Einträge in der Registry
' gelesen oder auch gesetzt werden, so dass sie automatisch vom
' MSInet oder Webbrowser-Control verwendet werden.
Option Explicit

Dim RegRoot As Long ' Registry Root z.B. HKEY_CURRENT_USER
Dim RegKey$         ' Der zu veraendernde Schluessel
Dim Sett As Long    ' Flag zum Unterdruecken mancher Funktionen

Dim OldProxy$
Dim OldIP$
Dim OldPort$
Dim oldEnabled As Long
Function Enable_HttpProxy(ByVal OnOff As Long) As Long

    Dim result As Long
    
    If OnOff <> 0 Then
        OnOff = 1      ' Die Registry verwendet eine 1 fuer "Wahr"
    End If
    
    result = RegValueSet(RegRoot, RegKey$, "ProxyEnable", OnOff)
    
    Enable_HttpProxy = result

End Function
Function Set_HttpProxy(ByVal prox$) As Long

    Dim result As Long
    
    prox$ = "http=" + prox$
    result = RegValueSet(RegRoot, RegKey$, "ProxyServer", prox$)
    
    Set_HttpProxy = result

End Function





Function SplitVB5(Source$, Delim$) As String

    ' Vb6 Benutzer benoetigen diese Funktion nicht
    Dim pos As Integer
    Dim LeftPart$
    pos = InStr(1, Source$, Delim$, 1)
    If pos > 0 Then
        LeftPart$ = Left$(Source$, pos - 1)
        Source$ = Mid$(Source$, pos + Len(Delim$))
    Else
        LeftPart$ = Source$
        Source$ = ""
    End If
    
    SplitVB5 = LeftPart$

End Function

Function is_ip(ByVal Source$) As Long

    ' Testet ob ein String wie eine IP-Adresse
    ' (also 4 dreistellige Zahlen, durch Punkt getrennt)
    ' aufgebaut ist.
    
    Dim test$
    Dim cnt As Long
    Dim i As Long
    
   
    For i = 1 To 3
        test$ = SplitVB5(Source$, ".")
        If IsNumeric(test$) Then
            cnt = cnt + 1
        End If
    Next i
    If IsNumeric(Source$) Then
        cnt = cnt + 1
    End If
    
    If cnt = 4 Then
        is_ip = -1
    End If
    
End Function

Function Read_HttpTimeout() As Long

    ' Hier eine weitere Funktion mit der man auslesen kann
    ' nach welcher Zeit (s), eine Seite neu geladen werden
    ' soll, anstatt aus dem Cache gelesen zu werden.

    Dim retval As Long
    Dim ret As Long
    
    ret = RegValueGet(RegRoot, RegKey$, "HttpDefaultExpiryTimeSecs", retval)
    
    Read_HttpTimeout = retval
    
End Function


Function Read_Proxy() As String

    Dim retstr$
    Dim pos As Long
    
    Dim ret
    ret = RegValueGet(RegRoot, RegKey$, "ProxyServer", retstr$)
    pos = InStr(1, retstr$, "http=", 1)
    If pos > 0 Then
        retstr$ = Mid$(retstr$, pos + 5) ' das "http:" entfernen
    End If
    pos = InStr(1, retstr$, ";", 1)
    If pos > 0 Then
        retstr$ = Left(retstr$, pos - 1)
    End If
    Read_Proxy = retstr$
    
End Function

Function Read_ProxyEnable() As Long

    Dim retval As Long
    Dim ret As Long
    
    ret = RegValueGet(RegRoot, RegKey$, "ProxyEnable", retval)
    
    Read_ProxyEnable = retval * -1
    
End Function


Sub ini_RegKeys()

    Dim result As Long

    RegRoot = HKEY_CURRENT_USER
    
    ' Dieser Schlüssel wird unter Windows 95 für MS Internet Explorer verwendet.
    ' Andere Betriebssyteme verwenden eventuell einen anderen Schüssel.
    RegKey$ = "Software\Microsoft\Windows\CurrentVersion\Internet Settings"
       
    'Testen ob Schlüssel existiert
    result = RegKeyExist(RegRoot, RegKey$)
    If result <> 0 Then
        MsgBox "Fehler!"
    End If
    
   
End Sub





Private Sub Beenden_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
    
    txtProxy.Text = OldProxy$
    TxtProxIP.Text = OldIP$
    TxtPort.Text = OldPort$
    If oldEnabled = 1 Then
        OptProxyEnable(0).Value = -1
    Else
        OptProxyEnable(1).Value = -1
    End If
    Call cmdSet_Click
        
End Sub

Private Sub cmdClear_Click()

    Sett = 1
    TxtProxIP.Text = ""
    txtProxy.Text = ""
    TxtPort.Text = ""
    Sett = 0
    
End Sub

Private Sub cmdRead_Click()

    Dim test$
    Dim result$
    Call cmdClear_Click
    Sett = 1
    test$ = Read_Proxy()
    result$ = SplitVB5(test$, ":")
    If is_ip(result$) Then
        TxtProxIP.Text = result$
    Else
        txtProxy.Text = result$
    End If
    If test$ <> "" Then
        TxtPort.Text = test$
    End If
        
    
    If Read_ProxyEnable() Then
        OptProxyEnable(0).Value = -1
    Else
        OptProxyEnable(1).Value = -1
    End If
    cmdSet.Enabled = 0
    Sett = 0
    
End Sub

Private Sub cmdSet_Click()

    Dim Proxy$
    Dim ret As Long
    
    Proxy$ = TxtProxIP.Text ' Die Ip wird bevorzugt, da hiermit kein
                            ' Nameserver aufgerufen werden muss.
    If Proxy$ = "" Then
        Proxy$ = txtProxy.Text ' Wenn das IP-Feld lerr ist,
    End If                     ' den Namen verwenden
    
    If Proxy$ = "" Then
        MsgBox "Geben Sie einen gültigen Proxserver an!", 64
        Exit Sub
    End If
    
    If TxtPort.Text = "" Then
        MsgBox "You must supply a proxy port !"
        Exit Sub
    Else
        If Val(TxtPort.Text) <> 0 Then
            Proxy$ = Proxy$ + ":" + TxtPort.Text ' Der Port wird nach ":" angehängt.
        Else
            MsgBox "Sie müssen einen Proxy Port eingeben!"
            Exit Sub
        End If
    End If
       
    ret = Set_HttpProxy(Proxy$)
    
    If OptProxyEnable(0).Value Then
        ret = Enable_HttpProxy(-1)
    Else
        ret = Enable_HttpProxy(0)
    End If
    cmdSet.Enabled = 0
    
    Call cmdRead_Click ' und wieder auslesen
    
End Sub

Private Sub Form_Load()

    Call ini_RegKeys    ' Schlüssel eintragen
    Call cmdRead_Click  ' Erstmal lesen...
    
    ' Urspruengliche Werte werden gespeichert.
    OldProxy$ = txtProxy.Text
    OldIP$ = TxtProxIP.Text
    OldPort$ = TxtPort.Text
    If OptProxyEnable(0).Value = -1 Then
        oldEnabled = 1
    End If
    
End Sub

Private Sub OptProxyEnable_Click(Index As Integer)

    If Sett = 1 Then
        Exit Sub
    End If
    cmdSet.Enabled = -1
  
End Sub


Private Sub TxtPort_Change()
    If Sett = 1 Then
        Exit Sub
    End If
    cmdSet.Enabled = -1
End Sub

Private Sub TxtProxIP_Change()
    If Sett = 1 Then
        Exit Sub
    End If
    cmdSet.Enabled = -1
End Sub

Private Sub txtProxy_Change()
    If Sett = 1 Then
        Exit Sub
    End If
    cmdSet.Enabled = -1
End Sub


