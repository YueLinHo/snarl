VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMissedNotificationsPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel

Dim mSelected As T_NOTIFICATION_INFO
Dim mSelIndex As Long

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
    Case "lb>missed"
        mSelIndex = Val(Value)
        g_NotificationRoster.MissedAt mSelIndex, mSelected

        ' /* toolbar */
        prefskit_SafeEnable mPanel, "toolbar_options", (Value <> "0")

        If mPanel.Find("toolbar_options", pc) Then
            prefskit_SetItem pc, 1, "enabled", IIf(mSelected.DefaultAck <> "", 1&, 0&)
            prefskit_SetItem pc, 2, "enabled", IIf(mSelected.Actions.CountItems, 1&, 0&)

        End If

        prefskit_SafeSetText mPanel, "the_title", Replace$(mSelected.Title, vbCrLf, "¶ ")
        prefskit_SafeSetText mPanel, "the_text", Replace$(mSelected.Text, vbCrLf, "¶ ")
        prefskit_SafeEnable mPanel, "fb>ack", (mSelected.DefaultAck <> "")

Debug.Print "********************************* " & mSelected.Actions.CountItems

        sz = "Created " & Format$(mSelected.DateStamp, "hh:mm:ss") & _
             IIf(DateDiff("d", Now, mSelected.DateStamp) = "0", " today", " on " & Format$(mSelected.DateStamp, "short date")) & _
             " by " & Replace$(mSelected.ClassObj.App.Name, vbCrLf, "¶ ")

        If Not (mSelected.Socket Is Nothing) Then _
            sz = sz & " on " & mSelected.Socket.RemoteHostIP

'        If mSelected.DefaultAck <> "" Then _
            sz = sz & "  Callback: " & mSelected.DefaultAck

        prefskit_SafeSetText mPanel, "the_detail", sz
'        prefskit_SafeSetText mPanel, "the_callback", Replace$(mSelected.DefaultAck, vbCrLf, "¶ ")

'        prefskit_SafeSetText mPanel, "sending_app", "Sent from " & mSelected.Sender & " using class '" & mSelected.Class & "'"
'        prefskit_SafeSetText mPanel, "sending_class", mSelected.Class

    Case "toolbar_options"

        Select Case Val(Value)
        Case 1
            If mSelected.DefaultAck <> "" Then _
                g_ProcessAck mSelected.DefaultAck

        Case 2
            With New OMMenu
                For i = 1 To mSelected.Actions.CountItems
                    .AddItem .CreateItem("", mSelected.Actions.TagAt(i).Name)
                    If i < mSelected.Actions.CountItems Then _
                        .AddSeparator

                Next i

                .Track Control.Page.Panel.hWnd
            
            End With

        Case 3
            ' /* copy to clipboard */
            Clipboard.SetText mSelected.Title & vbCrLf & mSelected.Text

        Case 4
            ' /* display */
            If Not g_IsDNDModeEnabled Then _
                g_NotificationRoster.Add mSelected, Nothing

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl

    Select Case Control.GetName()
    Case "fb>clear"
        If Not (g_NotificationRoster Is Nothing) Then
            g_NotificationRoster.ClearMissed
            uUpdateList

            prefskit_SafeSetText mPanel, "the_title", ""
            prefskit_SafeSetText mPanel, "the_text", ""
            prefskit_SafeSetText mPanel, "sending_app", ""

        End If


    Case "fb>ack"
        If mSelected.DefaultAck <> "" Then
            g_ProcessAck mSelected.DefaultAck

        ElseIf IsWindow(mSelected.hWndReply) <> 0 Then
            If mSelected.uReplyMsg <> 0 Then
                PostMessage mSelected.hWndReply, mSelected.uReplyMsg, SNARL_CALLBACK_INVOKED, ByVal mSelected.Token

            Else
                Debug.Print "TMissedNotificationsPanel.fb>ack(): bad reply message"

            End If

        Else
            Debug.Print "TMissedNotificationsPanel.fb>ack(): no ack or reply window"
        
        End If

    End Select

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

Public Sub Go()
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then
        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "Missed Notifications"
            .SetWidth 400

            Set pPage = new_BPrefsPage("", , Me)

            With pPage
                .SetMargin 0

                .Add new_BPrefsControl("label", "", "While you were away...")

                Set pm = New CTempMsg
                pm.Add "item-height", 36&
                Set pc = new_BPrefsControl("listbox", "lb>missed", , , , pm)
                pc.SizeTo 0, 160
                .Add pc

'                .Add new_BPrefsControl("label", "", "Title")
'                Set pm = New CTempMsg
'                pm.Add "locked", 1&
'                .Add new_BPrefsControl("fancyedit", "the_title", "", , , pm)
'
'                .Add new_BPrefsControl("label", "", "Text")
'                Set pm = New CTempMsg
'                pm.Add "locked", 1&
'                .Add new_BPrefsControl("fancyedit", "the_text", "", , , pm)
        
'                .Add new_BPrefsControl("label", "", "Source")
                .Add new_BPrefsControl("label", "the_detail", Space$(256))
        
'                .Add new_BPrefsControl("label", "sending_app", Space$(128))
'                .Add new_BPrefsControl("label", "sending_class", "Class:")

'                .Add new_BPrefsControl("label", "", "Callback")
'                .Add new_BPrefsControl("label", "the_callback", "")

                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Invoke|Actions|Copy to Clipboard|Display", , , , False)

'                .Add new_BPrefsControl("fancybutton2", "fb>ack", "Invoke Callback", , , , False)
        
'                .Add new_BPrefsControl("seperator", "")
                .Add new_BPrefsControl("fancybutton2", "fb>clear", "Clear List")
        
            End With

            .AddPage pPage


            .Go

            uUpdateList

            g_SetWindowIconToAppResourceIcon .hWnd

        End With

    Else

        g_ShowWindow mPanel.hWnd, True, True

    End If

End Sub

Private Sub uUpdateList()

    If (g_NotificationRoster Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If Not mPanel.Find("lb>missed", pc) Then _
        Exit Sub

Dim pInfo As T_NOTIFICATION_INFO
Dim pm As CTempMsg
Dim sz As String
Dim i As Long
Dim n As Long
Dim sza As String

    n = Val(pc.GetValue())

    With g_NotificationRoster

        prefskit_SafeEnable mPanel, "fb>clear", (.MissedCount > 0)

        If .MissedCount Then

            ' /* build the content string */

            For i = 1 To .MissedCount
                .MissedAt i, pInfo
                sza = Replace$(pInfo.Title, "|", ":")
                sza = Replace$(sza, vbCrLf, "¶")
                sz = sz & g_FormattedMidStr(sza, 54) & "#?0#?"

                sza = Replace$(pInfo.Text, "|", ":")
                sza = Replace$(sza, vbCrLf, "¶")
                sz = sz & g_FormattedMidStr(sza, 54) & "|"
'                sz = sz & g_FormattedMidStr(Replace$(pInfo.Title, "|", ":"), 54) & "#?0#?" & "From " & pInfo.Sender & " (" & pInfo.Class & ") at " & Format$(pInfo.DateStamp, "hh:mm:ss") & "|"

            Next i

            pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

            ' /* set the icons */

            Set pm = New CTempMsg
            For i = 1 To .MissedCount
                .MissedAt i, pInfo
                pm.Replace "index", i
                pm.Remove "image-object"
                pm.Add "image-object", pInfo.Icon
                pc.DoExCmd B_EXTENDED_COMMANDS.B_SET_ITEM, pm

            Next i

        Else
            pc.SetText ""

        End If

    End With

    pc.SetValue CStr(n + 1)

End Sub

Public Sub UpdateList()

    If Not (mPanel Is Nothing) Then _
        uUpdateList

End Sub
