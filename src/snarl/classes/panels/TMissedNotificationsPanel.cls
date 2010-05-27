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
Dim mPage As BPrefsPage

Dim mSelected As G_NOTIFICATION_CONTENT
Dim mSelIndex As Long

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    With mPage
        .Add new_BPrefsControl("label", "", "While you were away...")

        Set pm = New CTempMsg
        pm.Add "item-height", 28&

        Set pc = new_BPrefsControl("listbox", "lb>missed", , , , pm)
        pc.SizeTo 0, 160
        .Add pc

        .Add new_BPrefsControl("label", "", "Text")

        Set pm = New CTempMsg
        pm.Add "locked", 1&
        .Add new_BPrefsControl("fancyedit", "fe>text", "", , , pm)

        .Add new_BPrefsControl("fancybutton2", "fb>ack", "Acknowledge Notification", , , , False)

        .Add new_BPrefsControl("seperator", "")
        .Add new_BPrefsControl("fancybutton2", "fb>clear", "Clear List")

    End With

    uUpdateList

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pc As BControl

    Select Case Control.GetName()
    Case "lb>missed"
        mSelIndex = Val(Value)
        g_NotificationRoster.MissedAt mSelIndex, mSelected

        If mPage.Find("fe>text", pc) Then _
            pc.SetText Replace$(mSelected.Text, vbCrLf, " \n ")

        If mPage.Find("fb>ack", pc) Then _
            pc.SetEnabled (mSelected.Ack <> "")


    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl

    Select Case Control.GetName()
    Case "fb>clear"
        If Not (g_NotificationRoster Is Nothing) Then
            g_NotificationRoster.ClearMissed
            uUpdateList

        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 0

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

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
    Set mPage = Nothing

End Sub

Private Sub KPrefsPanel_Ready()

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)

End Sub

Public Sub Go()

    If Not (mPanel Is Nothing) Then
        g_ShowWindow mPanel.hWnd, True, True
        Exit Sub

    End If
    
    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .AddPage new_BPrefsPage("", , Me)
        .SetTitle "Missed Notifications"
        .SetWidth 400
        .Go
        g_SetWindowIconToAppResourceIcon .hWnd

    End With

End Sub

Private Sub uUpdateList()

    If (g_NotificationRoster Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If Not mPage.Find("lb>missed", pc) Then _
        Exit Sub

Dim pInfo As G_NOTIFICATION_CONTENT
Dim pm As CTempMsg
Dim sz As String
Dim i As Long
Dim n As Long

    n = Val(pc.GetValue())

    With g_NotificationRoster
        If .CountMissed Then

            ' /* build the content string */

            For i = 1 To .CountMissed
                .MissedAt i, pInfo
                sz = sz & Format$(pInfo.Timestamp, "hh:mm") & " " & pInfo.Title & "|"

            Next i

            pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

            ' /* set the icons */

            Set pm = New CTempMsg

            For i = 1 To .CountMissed
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

'    If n = 0 Then _
        n = 1

    pc.SetValue CStr(n)

End Sub

Public Sub UpdateList()

    If (mPanel Is Nothing) Then _
        Exit Sub

    uUpdateList

End Sub
