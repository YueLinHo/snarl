VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunFilePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunFilePanel.cls
    '/
    '/  Description:    Provides the Run File style configuration panel
    '/
    '/  Created:        13-Jan-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mSelIndex As Long

Dim mFile() As String
Dim mFiles As Long

Dim mIsRunFile As Boolean
Dim mConfFile As String

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pc As BControl
Dim sz As String
Dim i As Long

    Select Case Control.GetName()
    Case "list"
        mSelIndex = Val(Value)

        ' /* toolbar */
'        prefskit_SafeEnable mPanel, "toolbar_options", (Value <> "0")
'        prefskit_SafeSetText mPanel, "the_title", Replace$(mSelected.Title, vbCrLf, "¶ ")

'        If mPanel.Find("toolbar_options", pc) Then
''            prefskit_SetItem pc, 1, "enabled", IIf(mSelected.DefaultAck <> "", 1&, 0&)
'            prefskit_SetItem pc, 2, "enabled", IIf(mSelIndex <> 0, 1&, 0&)
'
'        End If

    Case "toolbar_options"

        Select Case Val(Value)
        Case 1
            ' /* add */
            sz = fd_QuickPickFile()
            If sz <> "" Then _
                uAddEntry sz

        Case 2
            If mSelIndex <> 0 Then _
                uRemEntry mSelIndex

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Function Go(ByVal IsRunFile As Boolean) As Long
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg
Dim sn As String
Dim sd As String

    mIsRunFile = IsRunFile
    If mIsRunFile Then
        mConfFile = "runfile"

    Else
        mConfFile = "image"

    End If

    If (mPanel Is Nothing) Then

        mFiles = 0
        ReDim mFile(0)

        With New CConfFile
            If .SetTo(g_MakePath(App.Path) & "etc\" & mConfFile & ".conf", True) Then
                .Rewind
                Do While .GetEntry(sn, sd)
                    mFiles = mFiles + 1
                    ReDim Preserve mFile(mFiles)
                    mFile(mFiles) = sd

                Loop
            End If

        End With

        
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "Run File Settings"
            .SetWidth 400
            .SetWindow 1

            Set pPage = new_BPrefsPage("", , Me)

            With pPage
                .SetMargin 0

                If mIsRunFile Then
                    .Add new_BPrefsControl("label", "", "Each scheme represents a file to run.  Snarl includes arguments as follows:")
                    .Add new_BPrefsControl("label", "", "<file> ""<notification title>"" ""<notification text>"" ""<application>""")

                End If

                Set pm = New CTempMsg
'                pm.Add "item-height", 36&
                Set pc = new_BPrefsControl("listbox", "list", , , , pm)
                pc.SizeTo 0, 160
                .Add pc

'                .Add new_BPrefsControl("fancyedit", "file_to_run")

''                .Add new_BPrefsControl("label", "", "Callback")
''                .Add new_BPrefsControl("label", "the_callback", "")

                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Add...|Remove")

''                .Add new_BPrefsControl("fancybutton2", "fb>ack", "Invoke Callback", , , , False)
'
''                .Add new_BPrefsControl("seperator", "")
'                .Add new_BPrefsControl("fancybutton2", "fb>clear", "Clear List")
        
            End With

            .AddPage pPage
            .Go

            uRefreshList

        End With

    End If

    Go = mPanel.hWnd

End Function

'Private Sub uUpdateList()
'
'    If (g_NotificationRoster Is Nothing) Then _
'        Exit Sub
'
'Dim pc As BControl
'
'    If Not mPanel.Find("lb>missed", pc) Then _
'        Exit Sub
'
'Dim pInfo As T_NOTIFICATION_INFO
'Dim pm As CTempMsg
'Dim sz As String
'Dim i As Long
'Dim n As Long
'Dim sza As String
'
'    n = Val(pc.GetValue())
'
'    With g_NotificationRoster
'
'        prefskit_SafeEnable mPanel, "fb>clear", (.MissedCount > 0)
'
'        If .MissedCount Then
'
'            ' /* build the content string */
'
'            For i = 1 To .MissedCount
'                .MissedAt i, pInfo
'                sza = Replace$(pInfo.Title, "|", ":")
'                sza = Replace$(sza, vbCrLf, "¶")
'                sz = sz & g_FormattedMidStr(sza, 54) & "#?0#?"
'
'                sza = Replace$(pInfo.Text, "|", ":")
'                sza = Replace$(sza, vbCrLf, "¶")
'                sz = sz & g_FormattedMidStr(sza, 54) & "|"
''                sz = sz & g_FormattedMidStr(Replace$(pInfo.Title, "|", ":"), 54) & "#?0#?" & "From " & pInfo.Sender & " (" & pInfo.Class & ") at " & Format$(pInfo.DateStamp, "hh:mm:ss") & "|"
'
'            Next i
'
'            pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
'
'            ' /* set the icons */
'
'            Set pm = New CTempMsg
'            For i = 1 To .MissedCount
'                .MissedAt i, pInfo
'                pm.Replace "index", i
'                pm.Remove "image-object"
'                pm.Add "image-object", pInfo.Icon
'                pc.DoExCmd B_EXTENDED_COMMANDS.B_SET_ITEM, pm
'
'            Next i
'
'        Else
'            pc.SetText ""
'
'        End If
'
'    End With
'
'    pc.SetValue CStr(n + 1)
'
'End Sub
'
'Public Sub UpdateList()
'
'    If Not (mPanel Is Nothing) Then _
'        uUpdateList
'
'End Sub

Private Sub uAddEntry(ByVal Path As String)

    mFiles = mFiles + 1
    ReDim Preserve mFile(mFiles)
    mFile(mFiles) = Path

Dim i As Long

    With New CConfFile
        For i = 1 To mFiles
            .Add "target", mFile(i)

        Next i

        .Save g_MakePath(App.Path) & "etc\" & mConfFile & ".conf"

    End With

    uRefreshList

End Sub

Private Sub uRemEntry(ByVal Index As Long)
Dim i As Long

    If Index < mFiles Then
        For i = Index To (mFiles - 1)
            mFile(i) = mFile(i + 1)

        Next i

    End If

    mFiles = mFiles - 1
    ReDim Preserve mFile(mFiles)

    With New CConfFile
        For i = 1 To mFiles
            .Add "target", mFile(i)

        Next i

        .Save g_MakePath(App.Path) & "etc\" & mConfFile & ".conf"

    End With

    uRefreshList

End Sub

Private Sub uRefreshList()
Dim pc As BControl

    If Not mPanel.Find("list", pc) Then _
        Exit Sub

Dim sz As String
Dim n As Long
Dim i As Long

    n = Val(pc.GetValue())

    If mFiles Then
        For i = 1 To mFiles
            sz = sz & mFile(i)
            If i < mFiles Then _
                sz = sz & "|"

        Next i

    End If

    pc.SetText sz
    pc.SetValue CStr(n)

End Sub

