VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGeneralPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage
'        .Add new_BPrefsControl("banner", "", "Launch Options")

        .SetMargin 96

        .Add new_BPrefsControl("banner", "", "Launch Options")

        Set pm = New CTempMsg
        pm.Add "text", "Start at login?"
        pm.Add "align", 1
        pm.Add "inset_by", 0
        .Add new_BPrefsControl("fullphat/fancytoggle", "run_on_logon", , "", IIf(g_Prefs.run_on_logon, "1", "0"), pm)
        .Add new_BPrefsControl("fancybutton2", "start", "Stop Snarl", "")
        .Add new_BPrefsControl("label", "lbl_running", "Snarl is running", "")
'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("banner", "", "Miscellaneous")

        ' /* miscellaneous */

        .Add new_BPrefsControl("banner", "", "Miscellaneous")
        .Add new_BPrefsControl("fancytoggle2", "show_msg_on_start", "Show message on startup?", "", IIf(g_Prefs.show_msg_on_start, "1", "0"), pm)
        .Add new_BPrefsControl("fancytoggle2", "sticky_snarls", "Sticky notifications?", , IIf(g_Prefs.sticky_snarls, "1", "0"))
        .Add new_BPrefsControl("fancytoggle2", "log_only", "Log only (don't display)?", , IIf(g_Prefs.log_only, "1", "0"))
        .Add new_BPrefsControl("fancytoggle2", "notify_on_first_register", "Show notification on first registration only?", , IIf(g_Prefs.notify_on_first_register, "1", "0"))
        .Add new_BPrefsControl("fancytoggle2", "auto_update", "Check for updates on launch?", "", IIf(g_Prefs.auto_update, "1", "0"))

Dim nIdle As Long

        .Add new_BPrefsControl("banner", "", "Behaviour")
        .Add new_BPrefsControl("label", "", "Enable sticky notifications after the following period of inactivity:")

        nIdle = ROUNDUP(g_Prefs.idle_timeout / 300)
        If (nIdle < 0) Or (nIdle > 3) Then _
            nIdle = 0

        .Add new_BPrefsControl("fancycycle", "idle_timeout", "Never|5 Minutes|10 Minutes|15 Minutes", "", CStr(nIdle + 1))

'        .Add new_BPrefsControl("fancytoggle2", "auto_sticky_on_screensaver", "Enable sticky notifications when screensaver starts?", "", g_ConfigGet("auto_sticky_on_screensaver"))

'        ' /* hotkey */
'
''        .Add new_BPrefsControl("separator", "")
'
'        .Add new_BPrefsControl("banner", "", "Hotkeys")
'
'        .Add new_BPrefsControl("fancytoggle2", "use_hotkey", "Use a hotkey to activate Snarl's Preferences?", "", IIf(g_Prefs.use_hotkey, "1", "0"), pm)
'
'        .Add new_BPrefsControl("key_picker", "hotkey_prefs", , , CStr(MOD_CONTROL) & "," & CStr(g_Prefs.hotkey_prefs), , g_Prefs.use_hotkey)
'        .Add new_BPrefsControl("label", "", "Press the key you want to use in the box the above.  Note that while you can select the key which activates Snarl, the modifiers (the combination of SHIFT and CTRL keys) used are automatically set.                            ")



'        .Add new_BPrefsControl("separator", "")
'        .Add new_BPrefsControl("fancybutton2", "go_app_manager", "Launch App Manager", "App Manager:")
'        .Add new_BPrefsControl("label", "", "The App Manager allows you to control other Snarl applications which don't have their own user interface.")
'
''        Set pc = new_BPrefsControl("label", "")
''        pc.SizeTo 0, 64
''        .Add pc
''
''        Set pm = New CTempMsg
''        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\snarl.png"
''        pm.Add "image-height", 32
''        pm.Add "valign", "centre"
''        .Add new_BPrefsControl("labelex", "", "Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ") " & App.LegalCopyright, , , pm)

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl
Dim sz() As String
Dim dw As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()
    Case "run_on_logon"
        g_Prefs.run_on_logon = CBool(Value)

    Case "show_msg_on_start"
        g_Prefs.show_msg_on_start = CBool(Value)

    Case "sticky_snarls"
        g_Prefs.sticky_snarls = CBool(Value)

    Case "log_only"
        g_Prefs.log_only = CBool(Value)

    Case "use_hotkey"
        ' /* R2.2: we have a separate config entry now */
        g_Prefs.use_hotkey = CBool(Value)
        If Control.Page.Find("hotkey_prefs", pc) Then _
            pc.SetEnabled CBool(Value)

        frmAbout.bSetHotkeys

    Case "hotkey_prefs"

        ' /* the key picker control should return a pair of values separated by a comma.  The first
        '    value is the set of modifiers; the second value is the keycode of the key pressed */

        sz() = Split(Value, ",")
        If UBound(sz()) <> 1 Then _
            Exit Sub

        ' /* we're only interested in the keycode here */

        dw = Val(sz(1))
        If dw <> g_Prefs.hotkey_prefs Then
            If frmAbout.bSetHotkeys(dw) Then
                ' /* registered okay so store the new keycode */
                g_Prefs.hotkey_prefs = dw
                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: hotkey changed to #" & sz(1)

            Else
                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: couldn't set hotkey to #" & sz(1), LEMON_LEVEL_WARNING
                sz(1) = CStr(g_Prefs.hotkey_prefs)

            End If
        End If

        Control.SetValue CStr(MOD_CONTROL) & "," & sz(1)

    Case "notify_on_first_register"
        g_Prefs.notify_on_first_register = CBool(Value)

    Case "auto_update"
        g_Prefs.auto_update = CBool(Value)

    Case "auto_sticky_on_screensaver"
        g_ConfigSet Control.GetName, Value
        Exit Sub

    Case "idle_timeout"
        dw = (Val(Value) - 1) * 300
        Debug.Print "new idle = " & dw & " secs"
        g_ConfigSet Control.GetName, CStr(dw)
        Exit Sub

    Case Else
        Exit Sub

    End Select

    g_WriteConfig

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
Dim pc As BControl

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()
    Case "start"
        ' /* toggle running state */
        g_SetRunning Not g_IsRunning
        Control.SetText IIf(g_IsRunning, "Stop Snarl", "Start Snarl")
        If Control.Page.Find("lbl_running", pc) Then _
            pc.SetText IIf(g_IsRunning, "Snarl is running", "Snarl is stopped")

    Case "go_app_manager"
        ShellExecute 0, "open", g_MakePath(App.Path) & "Snarl_App_Manager.exe", vbNullString, vbNullString, SW_SHOW

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
'    mPage.setmargin 0

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)

End Sub

Private Function KPrefsPage_hWnd() As Long

End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)

End Sub






