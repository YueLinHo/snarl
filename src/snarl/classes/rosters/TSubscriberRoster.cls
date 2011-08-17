VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSubscriberRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TSubscriberRoster.cls
    '/
    '/  Description:    Manages subscribed GNTP and SNP3 clients
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mSubscribers As BTagList
Dim mRefs As Long

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */

    End If

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT

    g_Debug "TSubscriberRoster.Open()", LEMON_LEVEL_PROC_ENTER

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        Set mSubscribers = new_BTagList()
        g_Debug "roster initialised"

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mSubscribers.CountItems

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
Dim i As Long

'    If mCount Then
'        Name = LCase$(Name)
'        For i = 1 To mCount
'            If LCase$(mApp(i).Name) = Name Then
'                MRoster_FindItem = i
'                Exit Function
'
'            End If
'        Next i
'
'        g_Debug "TApplicationRoster.FindItem(): app '" & Name & "' not found", LEMON_LEVEL_WARNING
'
'    Else
'        g_Debug "TApplicationRoster.FindItem(): no apps registered!", LEMON_LEVEL_CRITICAL
'
'    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.mObject

'    If (Index > 0) And (Index <= mCount) Then _
        Set MRoster_ItemAt = mApp(Index)

End Function

Private Function MRoster_NextItem(Item As melon.mObject) As Boolean

'    If mIndex <= mCount Then
'        Set Item = mApp(mIndex)
'        mIndex = mIndex + 1
'        MRoster_NextItem = True
'
'    End If

End Function

Private Sub MRoster_Rewind()

'    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "24-Jun-2011"           ' // creation date, not last modification date

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "subscriber.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

Public Function Add(ByRef Socket As CSocket, ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Data As BPackedData) As Boolean

    If (Socket Is Nothing) Or (Data Is Nothing) Then _
        Exit Function

Dim ps As TSubscriber

    If mSubscribers.Find(Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort), ps) Then
        ' /* already subscribed */
        g_Debug "TSubscriberRoster.Add(): '" & Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort) & "' is already subscribed", LEMON_LEVEL_INFO

        ' /* check password */
        If ps.Password <> Data.ValueOf("password") Then
            g_Debug "TSubscriberRoster.Add(): password mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_AUTH_FAILURE

        ElseIf Data.Exists("app-sig") Then
            ' /* add specific apps */
            Add = ps.AddApps(Data)

        Else
            ' /* cannot do this if already subscribed */
            g_Debug "TSubscriberRoster.Add(): cannot do wildcard subscribe", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_FAILED

        End If

    Else
        ' /* not already subscribed */
        Set ps = New TSubscriber
        ps.SetTo Socket, SubscriberType, Data
        mSubscribers.Add ps
        Add = True

    End If

End Function

Public Function Remove(ByRef Socket As CSocket, ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Data As BPackedData) As Boolean

    If (Socket Is Nothing) Or (Data Is Nothing) Then _
        Exit Function

Dim ps As TSubscriber
Dim i As Long

    i = mSubscribers.IndexOf(Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort))
    If i Then
        Debug.Print "TSubscriberRoster.Remove(): subscriber '" & Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort) & "' found"

        Set ps = mSubscribers.TagAt(i)

        If ps.SubscriptionType <> SubscriberType Then
            ' /* wrong type */
            g_Debug "TSubscriberRoster.Remove(): '" & Socket.RemoteHostIP & "' type mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_INVALID_ARG

        ElseIf ps.Password <> Data.ValueOf("password") Then
            ' /* wrong password */
            g_Debug "TSubscriberRoster.Remove(): password mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_AUTH_FAILURE

        Else
            ' /* remove */
            mSubscribers.Remove i
            Remove = True

        End If

    Else
        g_Debug "TSubscriberRoster.Remove(): '" & Socket.RemoteHostIP & "' is not subscribed", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_NOT_SUBSCRIBED

    End If

End Function

Friend Sub Broadcast(ByRef Content As T_NOTIFICATION_INFO)
Dim ps As TSubscriber

    With mSubscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            If ps.IsSubscribedTo(Content.ClassObj.App.Signature) Then _
                ps.send Content.OriginalContent, Content.ClassObj.App

        Loop

    End With

End Sub

Public Function List() As BTagList

    Set List = mSubscribers

End Function
