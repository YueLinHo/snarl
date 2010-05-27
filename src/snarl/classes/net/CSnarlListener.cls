VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnarlListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           CSnarlListener.cls
    '/
    '/  Description:    Represents a listening socket bound to port 9887 on a specific IP address
    '/
    '/  © 2008 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mConn() As CSnarlSocket
Dim mCount As Long
Dim mIp As String

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Public Sub Go(ByVal LocalIP As String)

    Set theSocket = New CSocket
    theSocket.Bind 9887, LocalIP
    theSocket.Listen

    mIp = LocalIP
    g_Debug "CSnarlListener.Go(): listening on " & theSocket.LocalIP & ":" & theSocket.LocalPort & "..."

End Sub

Public Sub Quit()
Dim sz As String

    If Not (theSocket Is Nothing) Then
        sz = theSocket.LocalIP & ":" & theSocket.LocalPort
        theSocket.CloseSocket
        Set theSocket = Nothing
        g_Debug "CSnarlListener.Quit(): listener on '" & sz & "' stopped"

    End If

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "CSnarlListener.OnConnectionRequest(): requestID=0x" & g_HexStr(requestID) & " if='" & mIp & "'", LEMON_LEVEL_PROC

    mCount = mCount + 1
    ReDim Preserve mConn(mCount)
    Set mConn(mCount) = New CSnarlSocket
    mConn(mCount).Accept requestID

End Sub
