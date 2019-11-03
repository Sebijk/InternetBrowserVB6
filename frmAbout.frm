VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info über Sebijk's Internet-Browser"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Tag             =   "Info Projekt1"
   Begin VB.CommandButton cmdShowInfoWindows 
      Caption         =   "Windows-Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   8
      Top             =   4080
      Width           =   1452
   End
   Begin VB.CommandButton cmdIEVersion 
      Caption         =   "IE-Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   7
      Top             =   3600
      Width           =   1452
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   780
      Left            =   120
      Picture         =   "frmAbout.frx":1E32
      ScaleHeight     =   720
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4260
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2640
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System-Info..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "&System-Info..."
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Verwendet Trident Engine, um Webseiten anzuzeigen. Symbol vom Desktop Tango Project. Einige Quelltexte stammen von ActiveVB.de"
      Height          =   975
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sebijk's Internet-Browser"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   5
      Tag             =   "Anwendungstitel"
      Top             =   240
      Width           =   4092
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Innen ausgefüllt
      Index           =   1
      X1              =   240
      X2              =   5672
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Lizenziert unter der GNU General Public License. Projektseite: https://www.github.com/Sebijk/sjBrowserVB6"
      ForeColor       =   &H00000000&
      Height          =   1740
      Left            =   240
      TabIndex        =   3
      Tag             =   "Warnung: ..."
      Top             =   2685
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

' Registrierungsschlüssel - Sicherheitsoptionen...
Const KEY_ALL_ACCESS = &H2003F

Option Explicit

' Registrierungsschlüssel - Sicherheitsoptionen...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Null-terminierte Unicode-Zeichenfolge
Const REG_DWORD = 4                      ' 32-Bit-Zahl


Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"


' Konstanten
Private Const KEY_QUERY_VALUE = &H1

' Ermitteln des Installationsverzeichnisses
' des Internet Explorers
Public Function IE_InstallPath() As String
  Dim sBuffer As String
  Dim lhKeyOpen As Long
  Dim nResult As Long
  Dim sKey As String
  
  ' Key, in dem der Pfad gespeichert ist
  sKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IEXPLORE.EXE"
  
  ' Puffer für den Installationspfad
  sBuffer = Space$(255)
  
  ' Registrier-Zweig öffnen
  If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0, _
    KEY_QUERY_VALUE, lhKeyOpen) = ERROR_SUCCESS Then
  
    ' Eintrag "Path" lesen
    If RegQueryValueEx(lhKeyOpen, "Path", 0, REG_SZ, sBuffer, 255) = 0 Then
      ' Installationspfad
      If InStr(sBuffer, Chr$(0)) > 0 Then _
        sBuffer = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
    
      If Right$(sBuffer, 1) = ";" Then _
        sBuffer = Left$(sBuffer, Len(sBuffer) - 1)
        
      IE_InstallPath = RTrim$(sBuffer)
    End If
    RegCloseKey (lhKeyOpen)
  End If
End Function


Private Sub cmdIEVersion_Click()
Dim sIEPath As String

sIEPath = IE_InstallPath()
If sIEPath <> "" Then
  MsgBox "Der Internet Explorer befindet sich unter: " & sIEPath
Else
  MsgBox "Internet Explorer ist scheinbar nicht vorhanden!"
End If

End Sub

Private Sub cmdShowInfoWindows_Click()
Call ShellAbout(fMainForm.hWnd, fMainForm.Caption, "© 2005-2006 Home of the Sebijk.de", fMainForm.Icon)
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub



Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Versuchen Namen und Pfad des Systeminfo-Programms aus der Registrierung zu lesen...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Versuchen nur den Pfad des Systeminfo-Programms aus der Registrierung zu lesen...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Sicherstellen, daß es sich um bekannte 32-Bit-Version handelt.
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Fehler - Datei nicht gefunden...
                Else
                        GoTo SysInfoErr
                End If
        ' Fehler - Registrierungseintrag nicht gefunden...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "Systeminformationsprogramm zur Zeit nicht verfügbar.", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long               ' Schleifenzähler
        Dim rc As Long              ' Rückgabe-Code
        Dim hKey As Long            ' Handle für geöffneten Registrierungsschlüssel
        Dim hDepth As Long          '
        Dim KeyValType As Long      ' Datentyp eines Registrierungsschlüssels
        Dim tmpVal As String        ' Teporärer Speicher für einen Registrierungswert
        Dim KeyValSize As Long      ' Größe der Registrierungsschlüssel-Variablen
        '------------------------------------------------------------
        ' Registrierungsschlüssel unter Stammverzeichnis
        ' öffnen {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehlerbehandlung...
        

        tmpVal = String$(1024, 0)                               ' Speicher für Variable reservieren
        KeyValSize = 1024                                       ' Größe der Variablen speichern
        

        '------------------------------------------------------------
        ' Registrierungswert abrufen...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Fehlerbehandlung
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Bestimmen des Datentyps für die Konvertierung...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Durchsuchen der Datentypen...
        Case REG_SZ                                             ' Datentyp Zeichenfolge
                KeyVal = tmpVal                                     ' Kopieren der Zeichenfolge
        Case REG_DWORD                                          ' Datentyp Doppelwort
                For i = Len(tmpVal) To 1 Step -1                    ' Konvertieren der einzelnen Bits
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Wert Zeichen für Zeichen erstellen
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Doppelwort in Zeichenfoge umwandeln
        End Select
        

        GetKeyValue = True                                      ' Wert für Erfolg zurückgeben
        rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
        Exit Function                                           ' Funktion verlassen
        

GetKeyError:    ' Aufräumen, nachdem ein Fehler aufgetreten ist...
        KeyVal = ""                                             ' Rückgabewert auf leere Zeichenfolge setzen
        GetKeyValue = False                                     ' Wert für Fehler zurückliefern
        rc = RegCloseKey(hKey)                                  ' Registrierungsschlüssel schließen
End Function

