Attribute VB_Name = "Modinetblock"
Option Explicit

Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (pTcpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type MIB_TCPROW
    dwState      As Long
    dwLocalAddr  As Long
    dwLocalPort  As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
End Type

Private udtRow As MIB_TCPROW

Private pDataRef As Long, nRows As Long, pTablePtr As Long

Public Function RefreshStack() As Boolean
    Dim nRet As Long
    
    pDataRef = 0
    
    nRet = AllocateAndGetTcpExTableFromStack(pTablePtr, 0, GetProcessHeap, 0, 2)

    If nRet = 0 Then
        CopyMemory nRows, ByVal pTablePtr, 4
        RefreshStack = True
    Else
        RefreshStack = False
    End If
    
End Function
Public Function EnumEntries() As Boolean
    Dim I As Long
    Dim nState As Long, nLocalAddr As Long, nLocalPort As Long
    Dim nRemoteAddr As Long, nRemotePort As Long, nProcId As Long
    
    EnumEntries = True
    
    If nRows = 0 Or pTablePtr = 0 Then
        EnumEntries = False
        Exit Function
    End If

    For I = 0 To nRows
        CopyMemory nState, ByVal pTablePtr + (pDataRef + 4), 4
        CopyMemory nLocalAddr, ByVal pTablePtr + (pDataRef + 8), 4
        CopyMemory nLocalPort, ByVal pTablePtr + (pDataRef + 12), 4
        CopyMemory nRemoteAddr, ByVal pTablePtr + (pDataRef + 16), 4
        CopyMemory nRemotePort, ByVal pTablePtr + (pDataRef + 20), 4
        CopyMemory nProcId, ByVal pTablePtr + (pDataRef + 24), 4
    
        DoEvents
        
        If nProcId < 70000 And nProcId > 0 And nState > 0 And nState < 13 Then
            
            Call TerminateConnection(nLocalAddr, _
                                     nLocalPort, _
                                     nRemoteAddr, _
                                     nRemotePort)
        End If
        
        pDataRef = pDataRef + 24
        DoEvents
    Next I
    
End Function
Public Sub TerminateConnection(xLocalAddr As Long, _
                               xLocalPort As Long, _
                               xRemoteAddr As Long, _
                               xRemotePort As Long)
    
    
    udtRow.dwLocalAddr = xLocalAddr
    udtRow.dwLocalPort = xLocalPort
    udtRow.dwRemoteAddr = xRemoteAddr
    udtRow.dwRemotePort = xRemotePort
    udtRow.dwState = 12
     
    Call SetTcpEntry(udtRow)

End Sub
