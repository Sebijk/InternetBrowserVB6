Attribute VB_Name = "Module1"

Public fMainForm As frmMain
Public FIBForm As frmInetBlock
Public Const scUserAgent = "Mozilla/5.0 (compatible; MSIE 8.0; Internet-Browser)"
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()


Sub Main()
    InitCommonControls
    frmSplash.Show
    frmSplash.Refresh
    For i = 1 To 5000
        For j = 1 To 7000
        Next j
    Next i
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash
    fMainForm.Show
    Set FIBForm = New frmInetBlock
    Load FIBForm
    FIBForm.Hide
End Sub


