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
    Case "show_on_screen", "use_style", "custom_timeout", "use_custom_icon", "custom_icon", "use_custom_sound", "custom_sound"
        gSelectedClass.SetValue Control.GetName, Value

        If Control.GetName = "custom_timeout" Then _
            Control.SetText IIf(Value = "0", "Sticky", Value & " second" & IIf(Value = "1", "", "s"))

        If (Control.GetName = "use_style") And (Value = "1") Then _
            gSelectedClass.SetStyleAndScheme uSelectedItem("style_name"), uSelectedItem("scheme_name")


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

    End Select

    ' /* update control status */

    uSetControls

'    Select Case Control.GetName
'    Case "show_on_screen"
'        Me.Update
'
'    End Select

'Dim fUpdateConfig As Boolean
'Dim pc As BControl
'Dim pm As CTempMsg
'Dim sz As String
'Dim dw As Long
'
'    Select Case Control.GetName()
'
'    Case "duration_cycle"
'        ' /* duration cycle button */
'        If Not mPanel.Find("fs>gen.timeout", pc) Then
'            g_Debug "TConfigureClasSPanel.ControlChanged(): can't find timeout slider", LEMON_LEVEL_CRITICAL
'            Exit Sub
'
'        End If
'
'        ' /* value translates to: App Decides (1) Default (2) Custom (3) */
'
'        If Val(Value) = 3 Then
'            ' /* custom duration */
'            pc.SetEnabled True
'            gSelectedClass.SetDuration Val(pc.GetValue)
'
'        Else
'            ' /* app decides or default */
'            pc.SetEnabled False
'            gSelectedClass.SetDuration IIf(Val(Value) = 1, E_CLASS_DURATION_APP_DECIDES, E_CLASS_DURATION_DEFAULT)
'
'        End If
'
'        fUpdateConfig = True
'
'
'    Case "fs>gen.timeout"
'        ' /* custom timeout slider (while it's being dragged) */
'        Select Case Val(Value)
'
'        Case 0:     Control.SetText "Sticky"
'        Case 1:     Control.SetText "1 sec"
'        Case Else:  Control.SetText Value & " secs"
'
'        End Select
'
'        gSelectedClass.SetDuration Val(Value)
'        fUpdateConfig = True
'
'
'    Case "cyc>position"
'        ' /* position cycle button */
'        gSelectedClass.SetPosition (Val(Value) - 2)
'        fUpdateConfig = True
'
'
'    Case "ft>use_custom_style"
'        ' /* use custom style */
'
'        gSelectedClass.SetUseCustomStyle Val(Value)
'
'        If mPanel.Find("cb>style", pc) Then _
'            pc.SetEnabled Val(Value)
'
'        If mPanel.Find("cb>scheme", pc) Then _
'            pc.SetEnabled Val(Value)
'
'        fUpdateConfig = True
'
'
'
'
'    Case "ipb>aud.sound"
'        ' /* cycle button - translates to: App Decides (1), Custom (2), None (3) */
'
'        Select Case Val(Value)
'        Case 1
'            ' /* app decides */
''            gSelectedClass.RemoveFlag (E_CLASS_CUSTOM_SOUND Or E_CLASS_NO_SOUND)
'            dw = 0
'
'        Case 2
'            ' /* custom sound */
''            gSelectedClass.RemoveFlag E_CLASS_NO_SOUND
''            gSelectedClass.AddFlag E_CLASS_CUSTOM_SOUND
'
'        Case 3
'            ' /* force no sound */
''            gSelectedClass.RemoveFlag E_CLASS_CUSTOM_SOUND
''            gSelectedClass.AddFlag E_CLASS_NO_SOUND
'
'        End Select
'
'        ' /* enable/disable the file picker */
'
'        If mPanel.Find("fp>aud.sound", pc) Then _
'            pc.SetEnabled (Val(Value) = 2)
'
'        fUpdateConfig = True
'
'
'    Case "fp>aud.sound"
'        If Value <> gSelectedClass.CustomSound Then
''            gSelectedClass.SetCustomSound Value
'            sz = g_GetPath(Value)
'            If sz <> "" Then _
'                g_Prefs.last_sound_folder = sz
'
'            Debug.Print "** LAST SOUND FOLDER = " & sz & " ***"
'            fUpdateConfig = True
'
'        End If
'
'
'
'
'    ' /* ---------------------------------- advanced page ---------------------------------- */
'
'
'    Case "cyc>suppression"
'        ' /* duplicate suppression cycle button - translates to App Decides (1), Blocked (2), Not Blocked (3), Delayed (4) */
'
''        gSelectedClass.RemoveFlag (E_CLASS_ALLOW_DUPLICATES Or E_CLASS_DELAY_DUPLICATES Or E_CLASS_NO_DUPLICATES)
'
'        Select Case Val(Value)
'        Case 1
'            ' /* nothing to do! */
'
'        Case 2
''            gSelectedClass.AddFlag E_CLASS_NO_DUPLICATES
'
'        Case 3
''            gSelectedClass.AddFlag E_CLASS_ALLOW_DUPLICATES
'
'        Case 4
''            gSelectedClass.AddFlag E_CLASS_DELAY_DUPLICATES
'
'        End Select
'
'        fUpdateConfig = True
'
'
'    Case "ft>adv.priority"
'        ' /* 'priority notification' toggle */
'        If Val(Value) Then
'            gSelectedClass.SetPriority True
'
'        Else
'            gSelectedClass.SetPriority False
'
'        End If
'
'        fUpdateConfig = True
'
'
'    Case "ft>adv.notify_only"
'        ' /* 'notification only' toggle */
'        If Val(Value) Then
'            gSelectedClass.SetClickThru True
'
'        Else
'            gSelectedClass.SetClickThru False
'
'        End If
'
'        fUpdateConfig = True
'
'
'
'
'    Case "ipb>net"
'        dw = Val(Value)
'        If mPanel.Find("lb>net", pc) Then _
'            pc.SetEnabled (dw = 2)
'
''        mInfoLabel.SetText IIf(dw = 1, "Notifications will not be forwarded", IIf(dw = 2, "Notifications will be forwarded to the following:", "Notifications will be forwarded to the list of IP addresses defined in the network tab"))
'
'    Case "ft>forward"
'        ' /* the forward/don't forward toggle button on the 'forwarding' page */
'        If mPanel.Find("chk>hide", pc) Then _
'            pc.SetEnabled Val(Value)
'
'
'
'
'
'
'
'
'
'    ' /* advanced page */
'
'    Case "ft>adv.notify_only"
'        If Val(Value) <> 0 Then
'            gSelectedClass.SetClickThru True
'
'        Else
'            gSelectedClass.SetClickThru False
'
'        End If
'
''        Debug.Print g_HexStr(gSelectedClass.PersistentFlags)
'
'    End Select
'
'
'    If (fUpdateConfig) And (Not mIgnoreChanges) Then _
'        g_WriteConfig
'

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName

    Case "test"
        If Not (gSelectedClass Is Nothing) Then _
            gSelectedClass.TestNotification

'    Case "fb2>use_default_style"
'        If Not (gSelectedClass Is Nothing) Then
'            gSelectedClass.SetCustomStyle "", ""
'            Me.Update
'
'        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
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

    mhWndPanel = hWndPanel
    Set mPanel = New BPrefsPanel

    With mPanel
        .SetHandler Me
        .SetWindow 1                                            ' // so it doesn't appear during Go()
        .SetWidth 360

        ' /* general page */

'        Set pp = new_BPrefsPage("General", , Me)
        Set pp = new_BPrefsPage("Basic", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-basic.png"), Me)
        With pp
            .SetMargin 90
            .SetWidth 300

            ' /* Display */

            .Add new_BPrefsControl("fancytoggle2", "show_on_screen", "Show on screen?", , gSelectedClass.GetValue("show_on_screen"))
            .Add new_BPrefsControl("fancycycle", "priority", "App Decides|Low|Normal|High", "Priority:")
            .Add new_BPrefsControl("fancycycle", "position", "Default|Top Left|Top Right|Bottom Left|Bottom Right", "Position:")
            .Add new_BPrefsControl("fancycycle", "duration", "Default|App Decides|Custom", "Duration:")

            ' /* R2.2: custom timeout slider */
            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 60&
            pm.Add "freq", 10&
            pm.Add "label_size", 64&
            .Add new_BPrefsControl("fancyslider", "custom_timeout", , , , pm)

            ' /* Style */

            .Add new_BPrefsControl("fancytoggle2", "use_style", "Use a specific style?", "Style:", gSelectedClass.GetValue("use_style"))
            .Add new_BPrefsCombo("style_name", "", "", , 24)
            .Add new_BPrefsCombo("scheme_name", "", "", , 24)

            ' /* test button */

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

        End With

        .AddPage pp



        Set pp = new_BPrefsPage("Advanced", load_image_obj(g_MakePath(App.Path) & "etc\icons\class-advanced.png"), Me)
        With pp
            .SetMargin 90
            .SetWidth 300

            ' /* R2.3: custom icon */
            .Add new_BPrefsControl("fancytoggle2", "use_custom_icon", "Use custom icon?", "Icon:")
            .Add new_BPrefsControl("fancyfilepicker", "custom_icon", , , "")

            ' /* R2.3: custom sound */
            .Add new_BPrefsControl("fancytoggle2", "use_custom_sound", "Play custom sound?", "Sound:")
            .Add new_BPrefsControl("fancyfilepicker", "custom_sound", , , "")

            .Add new_BPrefsControl("fancytoggle2", "ft>adv.notify_only", "Display in view-only mode?", "Other:")

            ' /* Duplicates */
            .Add new_BPrefsControl("fancycycle", "duplicates", "Allow|Block|Delay", "Duplicates:")

            ' /* forwarding */
            Set pm = New CTempMsg
            pm.Add "checkboxes", 1
            .Add new_BPrefsControl("listbox", "lb>notify", g_StyleRoster.GetNonWindowStyleList(), "Forwarding:", "1", pm)
            mCurrentForward = 1

        End With
        .AddPage pp

'                .Add new_BPrefsControl("fancycycle", "fc>fwd.when", "Never|Always|Away|Do Not Disturb Set", "When:")
'            .Add new_BPrefsControl("label", "", "View-only mode makes the notification invisible to the mouse pointer, allowing you to access content beneath it.  View-only notifications cannot be interacted with.")

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

Private Function uGetAudioSetting() As String

'    If (gSelectedClass.GetFlags And E_CLASS_NO_SOUND) Then
'        uGetAudioSetting = "3"
'
'    ElseIf (gSelectedClass.GetFlags And E_CLASS_CUSTOM_SOUND) Then
'        uGetAudioSetting = "2"
'
'    Else
'        ' /* app decides */
'        uGetAudioSetting = "1"
'
'    End If

End Function

Public Sub ClassChanged()

    If (gSelectedClass Is Nothing) Then _
        Exit Sub

    mIgnoreChanges = True

    mPanel.SetTitle gSelectedClass.App.Name & ": " & gSelectedClass.Description

    uSetValues
    uSetControls


'Dim pc As BControl




'Dim pca As BControl
'Dim pcb As BControl
'Dim lDur As E_CLASS_DURATION
'
'    ' /* notification duration cycle button on general page */
'
'    If (mPanel.Find("duration_cycle", pca)) And (mPanel.Find("fs>gen.timeout", pcb)) Then
'        ' /* pca is the cycle button
'        '    pcb is the duration slider */
'
'        lDur = gSelectedClass.Duration
'
'        If lDur >= E_CLASS_DURATION_CUSTOM Then
'            pca.SetValue "3"
'            pcb.SetValue CStr(lDur)
'            pcb.SetEnabled True
'
'        ElseIf lDur = E_CLASS_DURATION_DEFAULT Then
'            pca.SetValue "2"
'            pcb.SetValue CStr(g_Prefs.default_duration)
'            pcb.SetEnabled False
'
'        ElseIf lDur = E_CLASS_DURATION_APP_DECIDES Then
'            pca.SetValue "1"
'            pcb.SetValue CStr(g_Prefs.default_duration)
'            pcb.SetEnabled False
'
''            pcb.SetValue CStr(IIf(gSelectedClass.Duration = E_DEFAULT_DURATION, g_Prefs.default_duration, gSelectedClass.Duration))
'
'        End If
'    End If
'
'    ' /* R2.2: notification only toggle */
'
'    If mPanel.Find("ft>adv.notify_only", pca) Then _
'        pca.SetValue IIf(gSelectedClass.IsClickThru, "1", "0")
'
'    ' /* duplicate suppression cycle button - translates to App Decides (1), Blocked (2), Not Blocked (3), Delayed (4) */
'
'    If mPanel.Find("cyc>suppression", pca) Then _
'        pca.SetValue uGetSuppressionSetting()

    mIgnoreChanges = False

End Sub

Private Function uGetSuppressionSetting() As String

'    If (gSelectedClass.GetFlags() And E_CLASS_NO_DUPLICATES) Then
'        uGetSuppressionSetting = "2"
'
'    ElseIf (gSelectedClass.GetFlags() And E_CLASS_DELAY_DUPLICATES) Then
'        uGetSuppressionSetting = "4"
'
'    ElseIf (gSelectedClass.GetFlags() And E_CLASS_ALLOW_DUPLICATES) Then
'        uGetSuppressionSetting = "3"
'
'    Else
'        uGetSuppressionSetting = "1"
'
'    End If

End Function

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
    '    uSafeSetValue "is_priority", .GetValue("is_priority")
        
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

    End With

End Sub

Private Sub uSetControls()
Dim pc As BControl
Dim b2 As Boolean
Dim b As Boolean

    ' /* enable/disable controls */

    If mPanel.Find("show_on_screen", pc) Then
        b = (pc.GetValue = "1")

        uSafeEnable "position", b
        uSafeEnable "priority", b
        uSafeEnable "use_style", b

        ' /* only enable the combos if "show_on_screen" and "use custom style" are enabled */

        b2 = b
        If b2 Then
            If mPanel.Find("use_style", pc) Then _
                b2 = (pc.GetValue = "1")

        End If

        uSafeEnable "style_name", b2
        uSafeEnable "scheme_name", b2

        uSafeEnable "duration", b

        ' /* only enable the timeout slider if E_DURATION_CUSTOM is set */

        b2 = b
        If b2 Then
            If mPanel.Find("duration", pc) Then _
                b2 = (pc.GetValue = CStr(E_DURATION_CUSTOM + 1))

        End If
        
        uSafeEnable "custom_timeout", b2

        ' /* custom icon */

        uSafeEnable "use_custom_icon", b
        b2 = b
        If b2 Then
            If mPanel.Find("use_custom_icon", pc) Then _
                b2 = (pc.GetValue = "1")

        End If

        uSafeEnable "custom_icon", b2

        ' /* custom sound */

        uSafeEnable "use_custom_sound", b
        b2 = b
        If b2 Then
            If mPanel.Find("use_custom_sound", pc) Then _
                b2 = (pc.GetValue = "1")

        End If

        uSafeEnable "custom_sound", b2

    End If

End Sub

Private Function uSelectedItem(ByVal Name As String) As String
Dim pc As BControl

    If mPanel.Find(Name, pc) Then _
        uSelectedItem = prefskit_GetItem(pc)

End Function
