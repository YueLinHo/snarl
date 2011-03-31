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

'    Window As CSnarlWindow
'    Column As Integer           ' // used to help determine placement of other notifications
'    StyleName As String         ' // R2.31: expanded
'    SchemeName As String        ' // R2.31: expanded
'    Info As T_NOTIFICATION_INFO ' // R2.4 DR7: full originating information

Dim mWindow As CSnarlWindow
Dim mColumn As Integer          ' // used to help determine placement of other notifications
'Dim mStyleName As String        ' // R2.31: expanded
'Dim mSchemeName As String       ' // R2.31: expanded

Dim mInfo As T_NOTIFICATION_INFO
Dim mTempIcon As String
Dim mArgs As BPackedData

Dim WithEvents theIconRequest As BWebRequest
Attribute theIconRequest.VB_VarHelpID = -1

Public Function Init(ByVal URL As String, ByRef Args As BPackedData) As Boolean

    On Error Resume Next

    err.Clear
    Set theIconRequest = New BWebRequest
    If err.Number <> 0 Then
        g_Debug "TDelayedNotification.Init(): couldn't create BWebRequest", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mTempIcon = g_GetSafeTempIconPath() & "." & g_GetExtension(URL)
    If mTempIcon = "" Then
        g_Debug "TDelayedNotification.Init(): couldn't create local temporary file", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    Set mArgs = Args

    g_Debug "TDelayedNotification.Init(): saving as '" & mTempIcon & "'"
    Init = theIconRequest.GetFile(URL, mTempIcon)

End Function

Friend Sub SetTo(ByRef Info As T_NOTIFICATION_INFO, Optional ByRef Window As CSnarlWindow, Optional ByVal Column As Integer, Optional ByRef Frame As BRect) ', Optional ByVal StyleName As String, Optional ByVal SchemeName As String)

    LSet mInfo = Info
    Set mWindow = Window
    mColumn = Column
'    mStyleName = StyleName
'    mSchemeName = SchemeName

    If Not (mWindow Is Nothing) Then
        mWindow.Window.MoveTo Frame.Left, Frame.Top
        mWindow.Show

    End If

End Sub

Private Sub theIconRequest_Completed(ByVal WasSuccessful As Boolean)
Dim sz As String

    Debug.Print "TDelayedNotification.Completed(): " & WasSuccessful

    If Not WasSuccessful Then
        ' /* didn't download for some reason so use the class default icon, if one exists */
        If Not (mInfo.ClassObj Is Nothing) Then
            sz = mInfo.ClassObj.DefaultIcon

        Else
            ' /* to-do: allow user to specify an icon? */

        End If

    Else
        sz = mTempIcon

    End If

    mInfo.IconPath = sz

    If Not (g_NotificationRoster Is Nothing) Then _
        g_NotificationRoster.Add mInfo, mArgs

End Sub

Public Function Window() As CSnarlWindow

    Set Window = mWindow

End Function

Public Function Column() As Integer

    Column = mColumn

End Function

Friend Function Info() As T_NOTIFICATION_INFO

    LSet Info = mInfo

End Function

Public Function StyleName() As String

    StyleName = mInfo.StyleName

End Function

'Private Sub uFixupContent(ByRef Info As T_NOTIFICATION_INFO)
'
'    If (Info.ClassObj Is Nothing) Then
'        Debug.Print "########### uFixupContent missing class object ###########"
'        Exit Sub
'
'    End If
'
'    With Info
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'    End With
'
'Dim ppn As BPackedData
'
'    ' /* new */
'    Set ppn = New BPackedData
'
'    With ppn
'        .Add "id", Info.ClassObj.Name
'        .Add "title", Info.Title
'        .Add "text", Info.Text
'        .Add "timeout", CStr(Info.Timeout)
'        .Add "icon", Info.IconPath
'        .Add "priority", CStr(Info.Priority)
'        .Add "callback", Info.DefaultAck
'        .Add "value", Info.Value
'
''        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
'            .Add "merge", "1"
'
'    End With
'
'Dim ppd As BPackedData
'Dim szn As String
'Dim szv As String
'
'    ' /* add in all other content */
'
'    Set ppd = New BPackedData
'    With ppd
'        If .SetTo(Info.OriginalContent) Then
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                If Not ppn.Exists(LCase$(szn)) Then _
'                    ppn.Add szn, szv
'
'            Loop
'        End If
'    End With
'
'    Info.OriginalContent = ppn.AsString()
'
'End Sub

Public Sub Unset()

    Set mWindow = Nothing

End Sub

Public Sub Hide()

    If Not (mWindow Is Nothing) Then _
        mWindow.Hide

End Sub
