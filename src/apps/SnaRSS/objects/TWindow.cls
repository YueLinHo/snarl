VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2

Dim mhWnd As Long
Dim mFeedRoster As TFeedRoster

Dim WithEvents theAddFeedPanel As TAddFeedPanel
Attribute theAddFeedPanel.VB_VarHelpID = -1

Implements KPrefsPanel
Implements KPrefsPage
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    On Error Resume Next

    Select Case uMsg
    Case WM_CREATE
        frmMain.Add "WM_CREATE"
        slInitApp hWnd, App.ProductName, "snaRSS", g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "icon.png"
        mhWnd = hWnd
        uStart


    Case WM_DESTROY
        frmMain.Add "WM_DESTROY"
        If gToken Then _
            sn41UnregisterApp gToken


    Case WM_CLOSE
        PostQuitMessage 0


    Case sn41Broadcast()
        Select Case wParam
        Case SNARL_LAUNCHED
            frmMain.Add "Snarl launched (V" & CStr(lParam) & ")"
            If LoWord(lParam) >= 41 Then _
                uRegisterWithSnarl

        Case SNARL_QUIT
            frmMain.Add "Snarl quit"
            gToken = 0                      ' // no point in unregistering...

        End Select


    Case sn41AppMsg()
        Select Case wParam
        Case SNARL41_APP_PREFS
            frmMain.Add "_APP_PREFS"
            uDoPrefs

        Case SNARL41_APP_ABOUT
            sn41EZNotify gToken, "", _
                         "snaRSS " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments, _
                         App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision), , _
                         g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png"

        End Select


    Case slMsg()
        Select Case wParam
        Case SL_ABOUT
            sn41EZNotify gToken, "", _
                         "snaRSS " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments, _
                         App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision), , _
                         g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png"

        Case SL_PREFS
            frmMain.Add "_APP_PREFS"
            uDoPrefs

        Case SL_QUIT
            PostQuitMessage 0

'        Case SL_SHOW_ALL
'        Case SL_HIDE_ALL
'        Case SL_COUNT_WINDOWS
'        Case SL_SWITCH_TO

        End Select


    Case WM_NOTIFICATION
        frmMain.Add "NOTIFICATION: " & g_HexStr(wParam) & " " & CStr(lParam)
'        n = LoWord(wParam)

'        Select Case n
'        Case SNARL_NOTIFICATION_ACK, SNARL_NOTIFICATION_MENU
'            If Not uFind(lParam, pObj) Then _
'                Exit Function
'
'            frmmain.Add "item found: " & pObj.Class
'
'            If n = SNARL_NOTIFICATION_ACK Then
'                pObj.Display
'
'            ElseIf n = SNARL_NOTIFICATION_MENU Then
'
'                sn41Hide lParam
'                pObj.UnRead = False
'
'                Select Case HiWord(wParam)
'                Case 1
'                    pObj.Display
'
'                Case 2
'                    pObj.Reply.Display
'
'                Case 3
'                    pObj.ReplyAll.Display
'
'                Case 4
'                    pObj.Forward.Display
'
'                End Select
'
'            End If
'
'
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
'
'        End Select


    Case WM_TEST
        frmMain.Add "TEST: " & wParam

'        If wParam = 0 Then
'            uNotifyMail Nothing
'
'        ElseIf wParam = 1 Then
'            uNotifyMail myItems.Item(1)
'
'        ElseIf wParam = 2 Then
'            uFindFirst olMeetingRequest, pObj
'            Debug.Print VarType(pObj)
'            Err.Clear
'            uNotifyAppointment pObj
'            Debug.Print Err.Description
'
'        End If

'    Case Else
'        frmmain.add "WM_" & g_HexStr(uMsg)

    End Select

End Function

Private Sub uRegisterWithSnarl()

    On Error Resume Next

    gToken = sn41RegisterApp(App.ProductName, App.Title, g_MakePath(App.Path) & IIf(g_IsIDE(), "bin\", "") & "icon.png", mhWnd, WM_NOTIFICATION, SNARL41_APP_IS_WINDOWLESS Or SNARL41_APP_HAS_ABOUT Or SNARL41_APP_HAS_PREFS)
    If gToken Then
        frmMain.List1.Tag = CStr(gToken)
        frmMain.Add "Registered with Snarl (token=" & CStr(gToken) & ")"
'        sn41AddClass gtoken, "info", "Informational notifications"
'        sn41AddClass gtoken, "warn", "Warning notifications"
'        sn41AddClass gtoken, "crit", "Critical notifications"
'        sn41AddClass gtoken, "none", "Other notifications"

    Else
        frmMain.Add "Error registering with Snarl (" & CStr(sn41GetLastError()) & ")"

    End If

End Sub

Private Sub uStart()
Dim pFeedInfo As CConfSection

    ' /* set up */

    Set mFeedRoster = New TFeedRoster

    ' /* read config */

    With New CConfFile3
        .SetFile g_MakePath(App.Path) & "snaRSS.conf"
        If .Load Then
            .Rewind
            Do While .GetNextSection(pFeedInfo)
                mFeedRoster.AddFeed pFeedInfo.GetValueWithDefault("URL"), ""

            Loop
        End If
    End With

'        mFeedRoster.AddFeed "http://newsrss.bbc.co.uk/rss/newsonline_uk_edition/sci/tech/rss.xml", ""
'        mFeedRoster.AddFeed "http://newsrss.bbc.co.uk/rss/newsonline_uk_edition/england/rss.xml", ""
'        mFeedRoster.AddFeed "http://www.theregister.co.uk/headlines.atom", ""

    If sn41IsSnarlRunning() Then
        uRegisterWithSnarl

    Else
        frmMain.Add "Snarl not running, waiting..."

    End If

End Sub

Private Sub uDoPrefs()
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pm As CTempMsg

    If (gPanel Is Nothing) Then
        Set gPanel = New BPrefsPanel
        With gPanel
            .SetHandler Me
            .SetTitle "snaRSS Preferences"
            .SetWidth 400

            Set pPage = new_BPrefsPage("", Nothing, Me)

            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("banner", "", "RSS Feeds")

                Set pm = New CTempMsg
                pm.Add "item-height", 38
                pm.Add "checkboxes", 1&
                Set pCtl = new_BPrefsControl("listbox", "feed_list", , , "1", pm)
                pCtl.SizeTo 0, 160
                .Add pCtl

                Set pCtl = new_BPrefsControl("fancyplusminus", "feed_add_remove", "")
                .Add pCtl

                .Add new_BPrefsControl("fancytoolbar", "feed_toolbar", "Show Headline|Show Summary|Refresh|Feed Information", , , , False)

                .Add new_BPrefsSeparator

                .Add new_BPrefsControl("fancybutton2", "quit_app", "Quit snaRSS")
                .Add new_BPrefsControl("label", "", "snaRSS " & CStr(App.Major) & "." & CStr(App.Minor) & " (Build " & CStr(App.Revision) & ") " & App.LegalCopyright, , , , False)
                .Add new_BPrefsControl("label", "", "http://www.fullphat.net", , , , False)

            End With

            .AddPage pPage
            .Go
            g_SetWindowIconToAppResourceIcon .hWnd

            uUpdateFeedList

        End With
    End If

    g_ShowWindow gPanel.hWnd, True, True

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pList As BControl
Dim sz As String
Dim i As Long

    If gPanel.Find("feed_list", pList) Then _
        i = Val(pList.GetValue)

    Select Case Control.GetName

    Case "feed_list"
        prefskit_SafeEnable gPanel, "feed_toolbar", (Val(Value) <> 0)

    Case "feed_toolbar"

        If i = 0 Then _
            Exit Sub

        Select Case Value

        Case "1"
            mFeedRoster.FeedAt(i).ShowHeadline

        Case "2"
            mFeedRoster.FeedAt(i).ShowSummary

        Case "3"
            mFeedRoster.FeedAt(i).Refresh
            uUpdateFeedList

        Case "4"
            mFeedRoster.FeedAt(i).FeedInfo

        End Select


    Case "feed_add_remove"
        If Value = "+" Then
            If g_IsPressed(VK_CONTROL) Then
                theAddFeedPanel_AddFeed Clipboard.GetText()

            Else
                Set theAddFeedPanel = New TAddFeedPanel
                theAddFeedPanel.Go gPanel.hWnd

            End If

        ElseIf (Value = "-") And (i > 0) Then
            mFeedRoster.Remove i
            uUpdateConfig
            uUpdateFeedList

        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "quit_app" Then _
        PostQuitMessage 0

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

    If (Control.GetName = "feed_list") And (Notification = "update") Then _
        uUpdateFeedList

End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
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

    Set gPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub uUpdateFeedList()
Dim szCurrent As String
Dim pList As BControl
Dim pFeed As TFeed
Dim sz As String
Dim i As Long

    If Not gPanel.Find("feed_list", pList) Then _
        Exit Sub

    szCurrent = pList.GetValue
    If szCurrent = "0" Then _
        szCurrent = "1"

    With mFeedRoster
        .Rewind
        Do While .NextFeed(pFeed)
            sz = sz & g_FormattedMidStr(pFeed.TitleOrURL, 52) & "#?0#?" & pFeed.Status() & "|"

        Loop

    End With

    Debug.Print "**"
    Debug.Print sz
    Debug.Print "**"

    pList.SetText g_SafeLeftStr(sz, Len(sz) - 1)
    pList.SetValue szCurrent

    With mFeedRoster
        .Rewind
        Do While .NextFeed(pFeed)
            i = i + 1
            prefskit_SetItem pList, i, "checked", IIf(pFeed.IsEnabled, 1&, 0&)

        Loop

    End With

End Sub

'Private Function uSpecialMidStr(ByVal Text As String, ByVal MaxLen As Long) As String
'
'    If Len(Text) < MaxLen Then
'        uSpecialMidStr = Text
'        Exit Function
'
'    End If
'
'Dim i As Long
'
'    i = Fix((MaxLen - 3) / 2)
'    uSpecialMidStr = g_SafeLeftStr(Text, i) & "…" & g_SafeRightStr(Text, i)  '"…"
'
'End Function

Private Sub uUpdateConfig()
Dim pFile As CConfFile3
Dim pItem As CConfSection
Dim pFeed As TFeed

    ' /* update config */

    Set pFile = New CConfFile3
    pFile.SetFile g_MakePath(App.Path) & "snaRSS.conf"

    With mFeedRoster
        .Rewind
        Do While .NextFeed(pFeed)
            Set pItem = New CConfSection
            pItem.SetName "feed"
            pItem.Add "url", pFeed.URL
            pFile.Add pItem

        Loop

    End With

    pFile.Save

End Sub

Private Sub theAddFeedPanel_AddFeed(ByVal URL As String)

    If URL <> "" Then
        If mFeedRoster.AddFeed(URL, "") Then
            uUpdateConfig
            uUpdateFeedList
            prefskit_SetValue gPanel, "feed_list", prefskit_DoCmd(gPanel, "feed_list", B_COUNT_ITEMS)

        End If
    End If

End Sub


