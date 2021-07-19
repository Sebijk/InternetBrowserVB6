VERSION 5.00
Begin VB.Form frmGetUrlCache 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Internet-Cache verwalten"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   Icon            =   "frmGetUrlCache.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11835
   Begin VB.ListBox List2 
      Height          =   3960
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   9255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cookies aus Liste extrahieren"
      Height          =   495
      Left            =   9480
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Entfernen"
      Height          =   495
      Left            =   9480
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Beenden"
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   7680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Auflisten..."
      Height          =   495
      Left            =   9480
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmGetUrlCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Beispiel : Cache auflisten. Ausgewähltes Item aus dem Inhalt löschen. 2.
'Hinweis  : Zur Ausgabe verwendet dieses Beispiel eine Listbox. Beachten Sie die Aufnahmekapazität der ListBox.

Private CountFiles      As Long
Private c               As New Collection

Private Sub Command3_Click()
    'Item löschen...
    Dim Index As Long
    
    Index = List1.ListIndex
    If Index = -1 Then
        Call MsgBox("Kein Eintrag zum löschen markiert...", vbOKOnly Or vbCritical, "Fehler")
        Exit Sub
    Else
        Dim result As Long
        
        result = DeleteUrl(List1.List(List1.ListIndex))
        Command3.Enabled = False
        Command4.Enabled = False
        Select Case result
            Case 1
                Call MsgBox(result & vbNewLine & "Eintrag wurde gelöscht...", vbOKOnly Or vbInformation, "Info")
                Call Command2_Click
            Case 0, 2, 5
                Call MsgBox(result & vbNewLine & "Eintrag konnte nicht gelöscht werden...", vbOKOnly Or vbInformation, "Info")
                Call Command2_Click
        End Select
    End If
    Command3.Enabled = True
    Command4.Enabled = True
End Sub

Private Function DeleteUrl(ByVal URL As String) As Long
    Dim result As Long
    
    result = DeleteUrlCacheEntry(URL)
    Debug.Print result
    Select Case result
        Case Is = 0, 2, 5: DeleteUrl = result 'Löschvorgang konnte nicht ausgeführt werden
        Case Else: DeleteUrl = result         'Löschvorgang wurde ausgeführt
    End Select
End Function

Private Sub GetCacheContent()
    Dim CacheBuffer             As Long
    Dim SubsequentHandle        As Long
    Dim btHeap                  As Long
    Dim CacheEntry              As INTERNET_CACHE_ENTRY_INFO
    Dim FilesFromMem            As String
    Dim ptrResult               As String
        
    CacheBuffer = 0
    SubsequentHandle = FindFirstUrlCacheEntry(vbNullString, ByVal 0, CacheBuffer)
    Debug.Print SubsequentHandle
    If (SubsequentHandle = 0) And (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
    btHeap = LocalAlloc(LMEM_FIXED, CacheBuffer)
    If btHeap <> 0 Then
        Call CopyMemory(ByVal btHeap, CacheBuffer, 4)
        SubsequentHandle = FindFirstUrlCacheEntry(vbNullString, ByVal btHeap, CacheBuffer)
        If SubsequentHandle <> 0 Then
            Debug.Print SubsequentHandle
            Do
                Call CopyMemory(CacheEntry, ByVal btHeap, Len(CacheEntry))
                If CacheEntry.CacheEntryType And NORMAL_CACHE_ENTRY = NORMAL_CACHE_ENTRY Then
                    ptrResult = String$(lstrlen(ByVal CacheEntry.lpszSourceUrlName), 0)
                    Call lstrcpy(ByVal ptrResult, ByVal CacheEntry.lpszSourceUrlName)
                    CountFiles = CountFiles + 1
                    c.Add Item:=ptrResult, Key:=CStr(CountFiles)
                End If
                Call LocalFree(btHeap)
                CacheBuffer = 0
                Call FindNextUrlCacheEntry(SubsequentHandle, ByVal 0, CacheBuffer)
                btHeap = LocalAlloc(LMEM_FIXED, CacheBuffer)
                Call CopyMemory(ByVal btHeap, CacheBuffer, 4)
            Loop While FindNextUrlCacheEntry(SubsequentHandle, ByVal btHeap, CacheBuffer)
            End If
        End If
    End If
    Call LocalFree(btHeap)
    Call FindCloseUrlCache(SubsequentHandle)
End Sub

Private Sub Command1_Click()
    Call Unload(Me)
End Sub

Private Sub Command2_Click()
    Dim Items As Variant
    
    Me.MousePointer = vbHourglass
    Me.Caption = "Bitte warten..."
    Command1.Enabled = False
    Call GetCacheContent
    
    If List1.ListCount > 0 Then List1.Clear
    For Each Items In c
        List1.AddItem Items
    Next
    
    Me.Caption = CStr(CountFiles) & " Einträge gefunden..."
    Me.MousePointer = vbDefault
    CountFiles = 0
    
    Dim n As Long
    
    For n = 1 To c.Count
        DoEvents
        c.Remove 1
    Next
    Command3.Enabled = True
    Command1.Enabled = True
    Command4.Enabled = True
End Sub

Private Sub Command4_Click()
    Dim GetCookie As String
    Dim n As Long
    If List2.ListCount > 0 Then List2.Clear
    For n = 0 To List1.ListCount - 1
        GetCookie = List1.List(n)
        If InStr(GetCookie, "Cookie") <> 0 Then
            Call List2.AddItem(GetCookie)
        End If
    Next
End Sub

Private Sub Form_Load()
    Command3.Enabled = False
    Command4.Enabled = False
End Sub
