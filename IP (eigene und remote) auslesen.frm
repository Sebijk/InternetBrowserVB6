VERSION 5.00
Begin VB.Form frmIPForm 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "IP-Adresse auslesen"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "IP (eigene und remote) auslesen.frx":0000
   LinkTopic       =   "frmIPForm"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   4680
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Beenden"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "IP des Einwahlrechners:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "IP des eigenen Rechners:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmIPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !
'
'Autor:  Helge Rex   helge@activevb.de
'
'Auslesen der lokale Internet-Adresse (IP) und die Adresse
'des Einwahlrechners (Remote)

Option Explicit
 
'   API, zum Ermitteln des Handles zur aktiven DFÜ-Verbindung
Private Declare Function RasEnumConnections Lib "rasapi32.dll" _
    Alias "RasEnumConnectionsA" ( _
        lpRasCon As Any, _
        lpcb As Long, _
        lpcConnections As Long _
) As Long
 
'   API, mit der die zugangsdaten ermittelt werden
Private Declare Function RasGetProjectionInfo Lib "rasapi32.dll" _
    Alias "RasGetProjectionInfoA" ( _
        ByVal hRasConn As Long, _
        ByVal rasProjectionType As Long, _
        lpProjection As Any, _
        lpcb As Long _
) As Long
 
'   Eine kleine Speicherschieber-Funktion
Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
        Destination As Any, _
        Source As Any, _
        ByVal Length As Long _
        )
 
'   Ein paar Konstanten
Private Const RAS_MaxEntryName = 256
Private Const RAS_MaxDeviceType = 16
Private Const RAS_MaxDeviceName = 32
 
'   Datentyp für die DFÜ-Verbindungen
Private Type RASType
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS_MaxDeviceName) As Byte
End Type
 
'   Struktur für das TCP/IP-Protokoll
Private Type VBRASPPPIP
    dwSize As Long
    dwError As Long
    szClientIp As String
    szServerIp As String
End Type
 
'   helper function
Private Sub BytesToString(strToCopyTo As String, AbPosition As Byte, Laenge As Long)
    '   Speicher reservieren
    Dim strTemp As String
    Dim lngLen As Long
    
    '   Speicher zum Hineinkopieren bereitstellen
    strTemp = String(Laenge + 1, 0)
    
    '   Daten kopieren
    CopyMemory ByVal strTemp, AbPosition, Laenge
    
    '   Länge bis zum NullChar ermitteln
    lngLen = InStr(strTemp, Chr$(0)) - 1
    
    '   Rückgabe setzen
    strToCopyTo = Left$(strTemp, lngLen)
End Sub
 
Private Function VBRasGetRASPPPIP(hRasConn As Long, udtRASIP As VBRASPPPIP) As Long
    '    Speicher reservieren
    Dim Buffer() As Byte
    Dim Result As Long
    Dim StructSize As Long
    
    '   Größe der UDT festlegen
    StructSize = 40&
    
    '   Speicher für die API vorbereiten
    ReDim Buffer(StructSize - 1)
    
    '   Größe der UDT in die UDT kopieren
    CopyMemory Buffer(0), StructSize, 4
    
    '   IP-Adressen ermitteln
    Result = RasGetProjectionInfo(hRasConn, &H8021&, Buffer(0), StructSize)
    
    '   Rückgabe setzen
    VBRasGetRASPPPIP = Result
    
    '   War der Aufruf erfolgreich?
    If Result = 0 Then
        '   Ja, alle Daten kopieren
        With udtRASIP
            '   Größe der UDT kopieren
            CopyMemory .dwSize, Buffer(0), 4
            
            '   Fehlercode kopieren
            CopyMemory .dwError, Buffer(4), 4
            
            '   locale IP kopieren
            BytesToString .szClientIp, Buffer(8), 16
                        
            '   remote IP kopieren
            BytesToString .szServerIp, Buffer(24), 16
        End With
    End If
End Function
 
Private Function GetDFUEHandle() As Long
    '   Speicher reservieren
    Dim RAS(0 To 255) As RASType
    Dim StructSize As Long
    Dim DFUECount As Long
    Dim Result As Long
 
    '   Größe der Struktur festlegen
    RAS(0).dwSize = 412
    
    '   Größe der gesamten Abfrage festlegen
    StructSize = (UBound(RAS) - LBound(RAS) + 1) * RAS(0).dwSize
    
    '   Die DFÜ-Verbindungen abfragen
    Result = RasEnumConnections(RAS(0), StructSize, DFUECount)
 
    '   Wurde eine DFÜ-Verbindung gefunden?
    If (DFUECount <> 0) Then
        '   Ja, Handle zurückgeben
        GetDFUEHandle = RAS(0).hRasCon
    Else
        '   Nein, Nix zurückgeben
        GetDFUEHandle = 0
    End If
End Function
 
Private Sub Command1_Click()
    '   Dialog schließen
    Unload Me
End Sub
 
Private Sub Form_Load()
    '   Label und Textbox (eigene IP) beschriften
    Me.Text1.Text = vbNullString
    Me.Text1.Locked = True
    
    '   Label und Textbox (remote IP) beschriften
    Me.Text2.Text = vbNullString
    Me.Text2.Locked = True
    
    '   Timer setzen (5 Sekunden)
    Me.Timer1.Interval = 5000
    Me.Timer1.Enabled = True
    
    ' Gleich aufrufen
    Timer1_Timer
End Sub
 
Private Sub Timer1_Timer()
    '   Speicher reservieren
    Dim RASIP As VBRASPPPIP
    Dim RASHandle As Long
    
    '   Handle der Verbindung ermitteln
    RASHandle = GetDFUEHandle
    
    '   Wurde ein Handle gefunden?
    If (RASHandle <> 0) Then
        '   Ja, IPs abfragen
        Call VBRasGetRASPPPIP(RASHandle, RASIP)
        
        '   IPs mitteilen
        Me.Text1.Text = RASIP.szClientIp
        Me.Text2.Text = RASIP.szServerIp
    Else
        '   Nicht verbunden
        Me.Text1.Text = vbNullString
        Me.Text2.Text = vbNullString
    End If
End Sub



