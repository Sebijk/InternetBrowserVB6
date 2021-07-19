VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Optionen"
   ClientHeight    =   3750
   ClientLeft      =   2760
   ClientTop       =   3735
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ProxySettings 
      Caption         =   "Proxyserver konfigurieren"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CommandButton deletehistory 
      Caption         =   "Verlauf löschen"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton GetUrlCache 
      Caption         =   "Internet-Cache verwalten"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton InetSettings 
      Caption         =   "Internetoptionen"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   240
      Picture         =   "frmOption.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label TextLabel 
      Caption         =   "Wählen Sie eine Option aus, um Sebijk's Internet- Browser zu konfigurieren"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 ' clsid: (shlguid.h)  => SHDOCVW.DLL
Private Const clsid_CUrlHistory = "{3C374A40-BAE4-11CF-BF7D-00AA006946EE}"

' clsid: (urlhist.h)  => IUrlHistoryStg2
Private Const clsid_IUrlHistoryStg2 = "{AFA0DC11-C313-11D0-831A-00C04FD5AE38}"

' vtbl:  (urlhist.h)  => IUnknown-Release()
Private Const IUrlHistoryStg2_Release As Long = 8&

' vtbl:  (urlhist.h)  => HRESULT=ClearHistory()
Private Const IUrlHistoryStg2_ClearHistory As Long = 36&

' const: (WTYPES.h)   => ClassContext
Private Const CLSCTX_INPROC_SERVER As Long = 1&

' const: (WINERROR.h)
Private Const S_OK As Long = 0&

' const: (OAIDL.h)    => CallConvention
Private Const CC_STDCALL As Long = 4&

Private Declare Function CLSIDFromString Lib "ole32" ( _
    ByVal lpszProgID As Long, ByVal pCLSID As Long) As Long
    
Private Declare Function CoCreateInstance Lib "ole32" ( _
    ByVal rclsid As String, ByVal pUnkOuter As Long, _
    ByVal dwClsContext As Long, ByVal riid As String, _
    ByRef ppv As Long) As Long
    
Private Declare Sub DispCallFunc Lib "oleaut32" ( _
    ByVal ppv As Long, ByVal oVft As Long, _
    ByVal cc As Long, ByVal rtTYP As VbVarType, _
    ByVal paCNT As Long, ByRef paTypes As Long, _
    ByRef paValues As Long, ByRef fuReturn As Variant)


' hstDelete     Delete IE URL-history
'
' CALL:         hstDelete()
'
' IN:           ---
'
' OUT:          log     success
'
Public Function hstDelete() As Boolean
    Dim oid As String     ' object-id
    Dim iid As String     ' interface-id
    Dim ipt As Long       ' interface-ptr
    Dim ret As Variant

    oid = cnvCLSID(clsid_CUrlHistory)
    iid = cnvCLSID(clsid_IUrlHistoryStg2)
    
    If CoCreateInstance(oid, 0&, CLSCTX_INPROC_SERVER, iid, _
        ipt) = S_OK Then
        
        DispCallFunc ipt, IUrlHistoryStg2_ClearHistory, _
            CC_STDCALL, vbLong, 0, 0&, 0&, ret
            
        DispCallFunc ipt, IUrlHistoryStg2_Release, _
            CC_STDCALL, vbLong, 0, 0&, 0&, ret
            
        hstDelete = True
    End If
End Function


' cnvCLSID      Converts clsid-string to binary string (unicode)
'
' CALL:         cnvCLSID(clsid)
'
' IN:           chr:clsid   i.e. {3C374A40-BAE4-11CF-BF7D-00AA006946EE}
'
' OUT:          chr         16-byte converted string
'
Private Function cnvCLSID(clsid As String) As String
    Dim B1(15) As Byte
    
    CLSIDFromString StrPtr(clsid), VarPtr(B1(0))
    cnvCLSID = StrConv(B1, vbUnicode)
End Function

Private Sub deletehistory_Click()
Dim frage As VbMsgBoxResult
frage = MsgBox("Wollen Sie wirklich den Verlauf löschen?", vbExclamation + vbYesNo)
If frage = vbYes Then hstDelete Else Exit Sub
End Sub

Private Sub GetUrlCache_Click()
On Error Resume Next
    Dim CacheForm As New frmGetUrlCache
    CacheForm.Show
End Sub

Private Sub InetSettings_Click()
On Error Resume Next
Shell ("control.exe inetcpl.cpl")
End Sub

Private Sub ProxySettings_Click()
 Dim ProxyForm As New frmProxy
 ProxyForm.Show
End Sub

Private Sub TextLabel_Click()

End Sub
