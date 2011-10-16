VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TConfigureClassPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TConfigureClassPanel.cls
    '/
    '/  Description:    Provides the floating class configuration panel
    '/
    '/  Created:        9-Sep-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mhWndPanel As Long

Dim mIgnoreChanges As Boolean           ' // when set KPrefsPage_ControlChanged() doesn't update config
Dim mCurrentForward As Long
Dim mCurrentForwardName As String
Dim mCurrentCondition As Long

Public Event Done()

Implements KPrefsPanel
Implements KPrefsPage
Implements MWndProcSink

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    If (gSelectedClass Is Nothing) Or (mIgnoreChanges) Then _
        Exit Sub

Dim pc As BControl
Dim sz As String

    ' /* update values */

    Select Case Control.GetName
    Case "show_on_screen", "use_style", "custom_timeout", "use_custom_icon", "custom_icon", _
         "use_custom_sound", "custom_sound", "ignore_input", "use_custom_ack", "custom_ack", "log_during_dnd", "block_spam", _
         "override-normal", "override-away", "override-busy", "block_if_foreground", "use_custom_actions", "custom_actions", _
         "log_in_history"

        gSelectedClass.SetValue Control.GetName, Value

        Debug.Print Control.GetName

        If Control.GetName = "custom_timeout" Then
            Control.SetText IIf(Value = "0", "Sticky", Value & " second" & IIf(Value = "1", "", "s"))

        ElseIf (Control.GetName = "use_style") And (Value = "1") Then
            gSelectedClass.SetStyleAndScheme uSelectedItem("style_name"), uSelectedItem("scheme_name")

        End If

    Case "duration", "position", "duplicates"
        gSelectedClass.SetValue Control.GetName, CStr(Val(Value) - 1)


    Case "priority"
        Select Case Val(Value)
        Case 1
            ' /* app decides */
            gSelectedClass.SetValue "custom_priority", "0"

        Case 2
            gSelectedClass.SetValue "custom_priority", "1"
            gSelectedClass.SetValue "priority", "-1"

        Case 3
            gSelectedClass.SetValue "custom_priority", "1"
            gSelectedClass.SetValue "priority", "0"

        Case 4
            gSelectedClass.SetValue "custom_priority", "1"
            gSelectedClass.SetValue "priority", "1"

        End Select


    Case "style_name"
        ' /* 'style' combo */
        uUpdateSchemeList

    Case "scheme_name"
        ' /* 'scheme' combo */
        If mPanel.Find("style_name", pc) Then
            sz = prefskit_GetItem(pc)               ' // get the selected style from the style combo
            If sz <> "" Then _
                gSelectedClass.SetStyleAndScheme sz, prefskit_GetItem(Control)

        End If


    Case "lb>notify"
        mCurrentForward = Val(Value)
        mCurrentForwardName = LCase$(Replace$(prefskit_GetItem(Control), ": ", "/"))
        uSetConditions

    Case "when_list"
        mCurrentCondition = Val(Value)

    End Select

    ' /* update control status */

    uSetControls

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName

    Case "test"
        If Not (gSelectedClass Is Nothing) Then _
            gSelectedClass.TestNotification

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
Dim fChecked As Boolean
Dim sz As String

    Debug.Print Control.GetName & " - " & Notification

    Select Case Control.GetName
    Case "lb>notify"
        If (Notification = "checked") And (mCurrentForward <> 0) Then
            ' /* get item name */
            sz = prefskit_GetItem(Control, "text", mCurrentForward)

            Debug.Print gSelectedClass.Name & " > " & sz

            ' /* get tickbox state */
            If prefskit_GetItem(Control, "checked", mCurrentForward) = "1" Then
                gSelectedClass.AddForward sz

            Else
                gSelectedClass.RemForward sz

            End If

            uSetConditions

        End If


    Case "when_list"
        ' /* user is ticking items */
        If (Notification = "checked") And (mCurrentCondition <> 0) Then

            ' /* get tickbox state */
            fChecked = (prefskit_GetItem(Control, "checked", mCurrentCondition) = "1")

            Select Case mCurrentCondition

            Case 1
                ' /* when away */
                gSelectedClass.SetForward mCurrentForwardName, 1, fChecked

            Case 2
                ' /* when busy */
                gSelectedClass.SetForward mCurrentForwardName, 2, fChecked

            Case 3
                ' /* (other times) */
                gSelectedClass.SetForward mCurrentForwardName, 4, fChecked

            End Select

            uSetConditions

        End If

    End Select

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
Dim rc As RECT

    RaiseEvent Done
    g_ShowWindow mhWndPanel, True, True

    GetWindowRect mhWndPanel, rc
    g_MoveWindow mhWndPanel, Fix((g_ScreenWidth - (rc.Right - rc.Left)) / 2), rc.Top

End Sub

Private Sub KPrefsPanel_Ready()

    Debug.Print "TConfigureClassPanel.Ready"

    ClassChanged                ' // set controls to their correct values

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long)
Dim pp As BPrefsPage
Dim pm As CTempMsg
Dim pc As BControl
'Dim sz As String

    mhWndPanel = hWndPanel
    Set mPanel = New BPrefsPanel

    With mPanel
        .SetHandler Me
        .SetWindow 1                                            ' // so it doesn't appear during Go()
        .SetWidth 360

        ' /* general page */

'        Set pp = new_BPrefsPage("General", , Me)
        Set pp = new_BPrefsPage("Appearance", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-app.png"), Me)
        With pp
            .SetMargin 80
            .SetWidth 300

            ' /* display */
            .Add new_BPrefsControl("fancytoggle2", "show_on_screen", "Show on screen?", "Visibility:", gSelectedClass.GetValue("show_on_screen"))

            ' /* style */
            .Add new_BPrefsControl("fancytoggle2", "use_style", "Use a specific style?", "Style:", gSelectedClass.GetValue("use_style"))
            .Add new_BPrefsCombo("style_name", "", "", , 24)
            .Add new_BPrefsCombo("scheme_name", "", "", , 24)

            ' /* position */
            .Add new_BPrefsControl("fancycycle", "position", "Default|Top Left|Top Right|Bottom Left|Bottom Right", "Position:")

            ' /* R2.3: custom icon */
            .Add new_BPrefsControl("fancytoggle2", "use_custom_icon", "Use custom icon?", "Icon:")
            .Add new_BPrefsControl("fancyfilepicker", "custom_icon", , , "")

            ' /* view-only */
'            .Add new_BPrefsControl("fancytoggle2", "ignore_input", "Ignore mouse and other input events?", "Interaction:")

            ' /* test button */
            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With

        .AddPage pp


        Set pp = new_BPrefsPage("Behaviour", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-beh.png"), Me)
        With pp
            .SetMargin 80
            .SetWidth 300

'            If Not (gSelectedClass Is Nothing) Then _
'                sz = gSelectedClass.GetValue("custom_timeout")

            ' /* duration */
            .Add new_BPrefsControl("fancycycle", "duration", "Default|App Decides|Custom", "Duration:")
            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 60&
            pm.Add "freq", 10&
            pm.Add "label_size", 64&
            .Add new_BPrefsControl("fancyslider", "custom_timeout", , , , pm)

            ' /* priority */
            .Add new_BPrefsControl("fancycycle", "priority", "App Decides|Low|Normal|High", "Priority:")

            ' /* R2.3: custom sound */
            .Add new_BPrefsControl("fancytoggle2", "use_custom_sound", "Play custom sound?", "")
            .Add new_BPrefsControl("fancyfilepicker", "custom_sound", , "Sound:")

            ' /* R2.31: custom ack - renamed in R2.4 DR7 to "Callback" */
            .Add new_BPrefsControl("fancytoggle2", "use_custom_ack", "Custom callback?", "")
            .Add new_BPrefsControl("fancyedit", "custom_ack", , "Callback:")

'            ' /* R2.31: duplicates renamed to spamming */
'            .Add new_BPrefsControl("fancytoggle2", "block_spam", "Prevent notification flooding?", "Other:")

            ' /* R2.31: test button here too */
            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With

        .AddPage pp

'                .Add new_BPrefsControl("fancycycle", "fc>fwd.when", "Never|Always|Away|Do Not Disturb Set", "When:")
'            .Add new_BPrefsControl("label", "", "View-only mode makes the notification invisible to the mouse pointer, allowing you to access content beneath it.  View-only notifications cannot be interacted with.")



        Set pp = new_BPrefsPage("Redirection", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-fwd.png"), Me)
        With pp
            .SetMargin 0
            .SetWidth 300

            .Add new_BPrefsControl("label", "", "Send the notification content to the following:")

            ' /* list */

            Set pm = New CTempMsg
            pm.Add "checkboxes", 1
            Set pc = new_BPrefsControl("listbox", "lb>notify", g_StyleRoster.GetNonWindowStyleList(), "", "1", pm)
            pc.SizeTo 0, 136
            .Add pc
            mCurrentForward = 1

            ' /* R2.4 DR8: when to forward */

            .Add new_BPrefsControl("label", "", "When the following conditions are met:")

            Set pm = New CTempMsg
            pm.Add "checkboxes", 1
            Set pc = new_BPrefsControl("listbox", "when_list", "When away|When busy|(All other times)", "", "", pm)
            pc.SizeTo 0, 68
            .Add pc

            ' /* R2.31: test button here too */
            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With

        .AddPage pp




        Set pp = new_BPrefsPage("Presence", load_image_obj(g_MakePath(App.Path) & "etc\icons\presence.png"), Me)
        With pp
            .SetMargin 80
            .SetWidth 300

            .Add new_BPrefsControl("banner", "", "When I'm at my computer...")
            .Add new_BPrefsControl("fancycycle", "override-normal", "Default|Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority", "Action:", gSelectedClass.GetValue("override-normal"))

            .Add new_BPrefsControl("banner", "", "When I'm away...")
            .Add new_BPrefsControl("fancycycle", "override-away", "Default|Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority", "Action:", gSelectedClass.GetValue("override-away"))

'            ' /* R2.31: log during DnD */
'            .Add new_BPrefsControl("fancytoggle2", "log_during_dnd", "Always log when I'm busy", "")

            .Add new_BPrefsControl("banner", "", "When I'm busy...")
            .Add new_BPrefsControl("fancycycle", "override-busy", "Default|Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority", "Action:", gSelectedClass.GetValue("override-busy"))

            ' /* R2.31: test button here too */
            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With

        .AddPage pp



        Set pp = new_BPrefsPage("Advanced", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-ext.png"), Me)
        With pp
            .SetMargin 80
            .SetWidth 300

            ' /* duplicates */
'            .Add new_BPrefsControl("fancycycle", "duplicates", "Allow|Block|Delay", "Flooding:")

            .Add new_BPrefsControl("banner", "", "Behaviour")
            .Add new_BPrefsControl("fancytoggle2", "block_if_foreground", "Ignore if foreground application?", , gSelectedClass.GetValue("block_if_foreground"))

            .Add new_BPrefsControl("fancytoggle2", "ignore_input", "Ignore mouse and other input events?", "")
            .Add new_BPrefsControl("fancytoggle2", "block_spam", "Prevent notification flooding?", "")

            ' /* R2.4.2 DR3: custom actions */
            .Add new_BPrefsControl("fancytoggle2", "use_custom_actions", "Custom actions?", "")
            .Add new_BPrefsControl("fancyedit", "custom_actions", , "Actions:")
            .Add new_BPrefsControl("label", "", "<Action>=<command>.  Separate multiple actions with semi-colons.")

            .Add new_BPrefsControl("fancycycle", "log_in_history", "Always|Never|App Decides", "Log:", gSelectedClass.GetValue("log_in_history"))

            ' /* R2.31: test button here too */
            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With

        .AddPage pp




        .Go

        AddSubClass mhWndPanel, Me
        AddSubClass .hWnd, Me

        ' /* configure window look */

        g_SetWindowIconFromWindowIcon .hWnd, hWndPanel
        SetWindowLong .hWnd, GWL_STYLE, GetWindowLong(mPanel.hWnd, GWL_STYLE) And (Not WS_MINIMIZEBOX)

        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
'        EnableWindow hWndPanel, 0

Dim rc As RECT
Dim rc2 As RECT

        ' /* position */
        GetWindowRect hWndPanel, rc
        GetWindowRect .hWnd, rc2
        g_MoveWindow hWndPanel, Fix((g_ScreenWidth - (rc.Right - rc.Left) - (rc2.Right - rc2.Left)) / 2), rc.Top

        GetWindowRect hWndPanel, rc
        g_MoveWindow .hWnd, rc.Right - 4, rc.Top + Fix(((rc.Bottom - rc.Top) - (rc2.Bottom - rc2.Top)) / 2)
        g_ShowWindow .hWnd, True, True

    End With

End Sub

Public Sub Quit()

    RemoveSubClass mhWndPanel

    If Not (mPanel Is Nothing) Then
        RemoveSubClass mPanel.hWnd
        mPanel.Quit

    End If

End Sub

Public Sub ClassChanged()

    If (gSelectedClass Is Nothing) Then _
        Exit Sub

    mIgnoreChanges = True

    mPanel.SetTitle gSelectedClass.App.Name & ": " & gSelectedClass.Description

    uSetValues
    uSetControls

    uSetConditions

    mIgnoreChanges = False

End Sub

Private Sub uUpdateStyleList()
Dim pc As BControl

    If Not (mPanel.Find("style_name", pc)) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

    pc.SetText g_StyleRoster.GetStyleList

Dim pm As CTempMsg
Dim sz As String
Dim i As Long

    Set pm = New CTempMsg

    For i = 1 To g_StyleRoster.CountStyles()
        sz = g_StyleRoster.StyleAt(i).IconPath
        If Not g_Exists(sz) Then _
            sz = g_MakePath(App.Path) & "etc\icons\style.png"

        pm.Replace "image-file", sz
        pm.Replace "index", i
        pc.DoExCmd B_SET_ITEM, pm

    Next i

End Sub

Private Sub uUpdateSchemeList()
Dim pc As BControl

    If Not mPanel.Find("scheme_name", pc) Then _
        Exit Sub

    If (Not (mPanel.Find("style_name", pc))) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

Dim px As TStyle

    ' /* get the style object */

    Set px = g_StyleRoster.StyleAt(Val(pc.GetValue))

    ' /* get the scheme combo control */

    mPanel.Find "scheme_name", pc

    If (px Is Nothing) Then
        ' /* style not found so assume <Default> */
        g_Debug "TConfigureClassPanel.uUpdateSchemeList(): style not found, assuming <default>..."
        pc.SetText ""
        pc.SetEnabled False
        Exit Sub

    End If

Dim pm As CTempMsg
Dim i As Long

    If px.CountSchemes = 0 Then
        g_Debug "TConfigureClassPanel.uUpdateSchemeList(): style '" & px.Name & "' has no schemes", LEMON_LEVEL_WARNING
        pc.SetText ""

    Else
        pc.SetText px.GetSchemeList()

        ' /* set the icons */

        Set pm = New CTempMsg
        For i = 1 To px.CountSchemes                '// pc.DoExCmd(B_COUNT_ITEMS)
            pm.Replace "index", i
            pm.Replace "image-file", g_MakePath(px.Path) & px.SchemeAt(i) & ".icon"
            pc.DoExCmd B_SET_ITEM, pm

        Next i

        ' /* select the first item */
        pc.SetValue "1"

    End If

    ' /* enable/disable the combo based on the scheme count */

    If pc.IsEnabled Then _
        pc.SetEnabled (px.CountSchemes > 0)

End Sub

Public Sub SetFocus()

    If Not (mPanel Is Nothing) Then _
        g_WindowToFront mPanel.hWnd, True

End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static rc As RECT
Static rcOwner As RECT

    If (mPanel Is Nothing) Then _
        Exit Function

    Select Case uMsg

    Case WM_MOVE
        If hWnd = mhWndPanel Then
            GetWindowRect hWnd, rcOwner
            GetWindowRect mPanel.hWnd, rc
            g_MoveWindow mPanel.hWnd, rcOwner.Right - 4, _
                                      rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - (rc.Bottom - rc.Top)) / 2)

        End If

    Case WM_NCHITTEST
        If hWnd = mPanel.hWnd Then
            rc.Left = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)
            If rc.Left = HTCAPTION Then
                ReturnValue = 50
                MWndProcSink_WndProc = True

            End If
        End If

'    Case WM_NCACTIVATE
'        If (wParam = 0) And (hWnd = mPanel.hWnd) Then
'            ReturnValue = 0
'            MWndProcSink_WndProc = True
'
'        End If

    End Select

End Function

Private Sub uSafeEnable(ByVal Name As String, ByVal Enabled As Boolean)
Dim pc As BControl

    If mPanel.Find(Name, pc) Then
        If pc.IsEnabled <> Enabled Then _
            pc.SetEnabled Enabled

    End If

End Sub

Private Sub uSafeSetValue(ByVal Name As String, ByVal Value As String)
Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        pc.SetValue Value

End Sub

Private Sub uSafeSetText(ByVal Name As String, ByVal Text As String)
Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        pc.SetText Text

End Sub

Private Sub uSetValues()
Dim dw As Long

    With gSelectedClass

        uSafeSetValue "show_on_screen", .GetValue("show_on_screen")
        uSafeSetValue "ignore_input", .GetValue("ignore_input")

        If .GetValue("custom_priority") = "0" Then
            ' /* app decides */
            uSafeSetValue "priority", "1"

        Else
            Select Case Val(.GetValue("priority"))
            Case Is < 0
                uSafeSetValue "priority", "2"

            Case 0
                uSafeSetValue "priority", "3"

            Case Is > 0
                uSafeSetValue "priority", "4"
        
            End Select

        End If

        dw = Val(.GetValue("duplicates"))
        uSafeSetValue "duplicates", CStr(dw + 1)

        uSafeSetValue "use_style", .GetValue("use_style")
    
        dw = Val(.GetValue("duration"))
        uSafeSetValue "duration", CStr(dw + 1)
    
        uSafeSetValue "custom_timeout", .GetValue("custom_timeout")
    
        dw = Val(.GetValue("position"))
        uSafeSetValue "position", CStr(dw + 1)
    
        uUpdateStyleList
    
Dim b As Boolean
Dim sz As String
Dim i As Long
Dim j As Long
    
        b = (.GetValue("use_style") = "1")
        If b Then
            ' /* look it up */
            sz = .GetValue("style")
            i = g_StyleRoster.IndexOf(style_GetStyleName(sz))
            If i > 0 Then
                ' /* look up the scheme */
                j = g_StyleRoster.StyleAt(i).SchemeIndex(style_GetSchemeName(sz))
                Debug.Print "index of " & style_GetSchemeName(sz) & " : " & j
    
            Else
                g_Debug "TAppsPage.uUpdateSubPages(): style '" & sz & "' not found", LEMON_LEVEL_WARNING
    
            End If
    
        End If
    
        If (i = 0) Or (j = 0) Then
            ' /* style and/or scheme not found OR we're using the default */
            sz = g_ConfigGet("default_style")
            i = g_StyleRoster.IndexOf(style_GetStyleName(sz))
            If i > 0 Then _
                j = g_StyleRoster.StyleAt(i).SchemeIndex(style_GetSchemeName(sz))
    
        End If
    
        uSafeSetValue "style_name", CStr(i)
        uUpdateSchemeList
    
        uSafeSetValue "scheme_name", CStr(j)


        uSafeSetValue "use_custom_icon", .GetValue("use_custom_icon")
        uSafeSetText "custom_icon", .GetValue("custom_icon")

        uSafeSetValue "use_custom_sound", .GetValue("use_custom_sound")
        uSafeSetText "custom_sound", .GetValue("custom_sound")

        ' /* R2.31d3 */

        uSafeSetValue "use_custom_ack", .GetValue("use_custom_ack")
        uSafeSetText "custom_ack", .GetValue("custom_ack")

        uSafeSetValue "log_during_dnd", .GetValue("log_during_dnd")
        uSafeSetValue "block_spam", .GetValue("block_spam")

        ' /* R2.4.2 DR3 */

        uSafeSetValue "use_custom_actions", .GetValue("use_custom_actions")
        uSafeSetText "custom_actions", .GetValue("custom_actions")


        ' /* forwarding */

Dim pc As BControl
Dim pm As CTempMsg

        If mPanel.Find("lb>notify", pc) Then
            j = Val(pc.DoExCmd(B_COUNT_ITEMS))
            If j Then
                For i = 1 To j
                    Debug.Print CStr(i) & " >> " & prefskit_GetItem(pc, , i) & " " & .HasForward(prefskit_GetItem(pc, , i))

                    Set pm = New CTempMsg
                    pm.Add "index", i
                    pm.Add "checked", IIf(.HasForward(prefskit_GetItem(pc, , i)), 1&, 0&)
                    pc.DoExCmd B_SET_ITEM, pm

                Next i
            End If

        End If

        uSafeSetValue "override-normal", .GetValue("override-normal")
        uSafeSetValue "override-away", .GetValue("override-away")
        uSafeSetValue "override-busy", .GetValue("override-busy")

    End With

End Sub

Private Sub uSetControls()
Dim pc As BControl
Dim B2 As Boolean
Dim b As Boolean

    ' /* enable/disable controls */

    If mPanel.Find("show_on_screen", pc) Then
        b = (pc.GetValue = "1")

        uSafeEnable "position", b
        uSafeEnable "priority", b
        uSafeEnable "use_style", b

        ' /* only enable the combos if "show_on_screen" and "use custom style" are enabled */

        B2 = b
        If B2 Then
            If mPanel.Find("use_style", pc) Then _
                B2 = (pc.GetValue = "1")

        End If

        uSafeEnable "style_name", B2
        uSafeEnable "scheme_name", B2

        uSafeEnable "duration", b

        ' /* only enable the timeout slider if SN_ND_CUSTOM is set */

        B2 = b
        If B2 Then
            If mPanel.Find("duration", pc) Then _
                B2 = (pc.GetValue = CStr(SN_ND_CUSTOM + 1))

        End If
        
        uSafeEnable "custom_timeout", B2

        ' /* custom icon */

        uSafeEnable "use_custom_icon", b
        B2 = b
        If B2 Then
            If mPanel.Find("use_custom_icon", pc) Then _
                B2 = (pc.GetValue = "1")

        End If

        uSafeEnable "custom_icon", B2

        ' /* custom sound */

        uSafeEnable "use_custom_sound", b
        B2 = b
        If B2 Then
            If mPanel.Find("use_custom_sound", pc) Then _
                B2 = (pc.GetValue = "1")

        End If

        uSafeEnable "custom_sound", B2

        ' /* custom ack */

        uSafeEnable "use_custom_ack", b
        B2 = b
        If B2 Then
            If mPanel.Find("use_custom_ack", pc) Then _
                B2 = (pc.GetValue = "1")

        End If

        uSafeEnable "custom_ack", B2

    End If

End Sub

Private Function uSelectedItem(ByVal Name As String) As String
Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        uSelectedItem = prefskit_GetItem(pc)

End Function

Private Sub uSetConditions()
Dim pc As BControl

    If (Not mPanel.Find("when_list", pc)) Or (gSelectedClass Is Nothing) Or (mCurrentForward = 0) Then _
        Exit Sub

Dim pm As CTempMsg
Dim i As Long

    i = Val(gSelectedClass.ForwardList.ValueOf(mCurrentForwardName))
    Debug.Print "current forward: " & mCurrentForwardName & " flags=" & g_HexStr(i)

Dim n As Long

    For n = 0 To 2
        Set pm = New CTempMsg
        pm.Add "index", n + 1
        pm.Add "checked", IIf((i And 2 ^ n), 1&, 0&)
        pc.DoExCmd B_SET_ITEM, pm

    Next n

End Sub
