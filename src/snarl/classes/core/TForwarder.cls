VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TForwarder.cls
    '/
    '/  Description:    Represents a single fire-and-forget notification forwarding instance
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mInfo As notification_info
Dim mUID As Long

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Public Sub Go(ByVal RemoteComputer As String, ByRef NotificationInfo As notification_info, ByVal UID As Long)

    LSet mInfo = NotificationInfo
    mUID = UID

    ' /* we're a fire-and-forget notification, so we initiate an asynchronous connection
    '    to the destination and start a timer (currently 3 seconds).  If the socket
    '    connects before the timer expires, we send the notification info */

    ' /* TO DO: make the timeout user configurable */

    Set theSocket = New CSocket
    theSocket.Connect RemoteComputer, 9887
    Set theTimer = new_BTimer(3000, True)

End Sub

Public Function UID() As Long

    UID = mUID

End Function

Private Sub theSocket_OnConnect()

    Debug.Print "### " & theSocket.RemoteHostIP & ": connected"
    Set theTimer = Nothing

Dim sz() As String

    ' /* we ask for the app title as well, so on entry, notificationinfo.Title
    '    actually looks like this: <creating app>#?<title> */

    sz = Split(Replace$(mInfo.Title, vbCrLf, "\n"), "#?")

    ' /* we must also escape any CRLF's back to \n's otherwise the CR will
    '    act as a packet terminator */

    theSocket.SendData "snp://notify?title=" & g_EncodeURL(sz(1)) & _
                       "&text=" & g_EncodeURL(Replace$(mInfo.Text, vbCrLf, "\n") & "\n\n" & sz(0) & " on " & g_GetComputerName()) & _
                       "&icon=" & mInfo.Icon & _
                       vbCr

End Sub

Private Sub theSocket_OnSendComplete()

    Debug.Print "### " & theSocket.RemoteHostIP & ": send complete"
    g_ForwardRemove mUID

End Sub

Private Sub theTimer_Pulse()

    g_Debug "TForwarder:: timed out connecting to '" & theSocket.RemoteHostIP & "'"
    g_ForwardRemove mUID

End Sub
