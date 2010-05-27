VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDelayedNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TDelayedNotification.cls
'/
'/  Description:    Holding point for a notification whose icon is being downloaded
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

'// "http://www.google.com/ig/images/weather/chance_of_rain.gif"

Dim mClass As TAlert
Dim mInfo As T_NOTIFICATION_INFO
Dim mTempIcon As String

Dim WithEvents theIconRequest As BWebRequest
Attribute theIconRequest.VB_VarHelpID = -1

Public Function Init(ByVal URL As String) As Boolean

    On Error Resume Next

    Err.Clear
    Set theIconRequest = New BWebRequest
    If Err.Number <> 0 Then
        g_Debug "TDelayedNotification.Init(): couldn't create BWebRequest", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mTempIcon = g_GetSafeTempIconPath()
    If mTempIcon = "" Then
        g_Debug "TDelayedNotification.Init(): couldn't create local temporary file", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    g_Debug "TDelayedNotification.Init(): temporary icon is '" & mTempIcon & "'"

    Init = theIconRequest.GetFile(URL, mTempIcon)

End Function

Friend Sub SetTo(ByRef Info As T_NOTIFICATION_INFO, ByRef Class As TAlert)

    LSet mInfo = Info
    Set mClass = Class

End Sub

Private Sub theIconRequest_Completed(ByVal WasSuccessful As Boolean)
Dim sz As String

    If Not WasSuccessful Then
        ' /* didn't download for some reason so use the class default icon, if one exists */
        If Not (mClass Is Nothing) Then
            sz = mClass.DefaultIcon

        Else
            ' /* to-do: allow user to specify an icon? */

        End If

    Else
        sz = mTempIcon

    End If

    mInfo.IconPath = sz

    If Not (g_NotificationRoster Is Nothing) Then _
        g_NotificationRoster.Add mClass, mInfo

End Sub
