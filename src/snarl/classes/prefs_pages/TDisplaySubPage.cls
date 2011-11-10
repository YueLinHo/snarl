VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDisplaySubPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TDisplaySubPage.cls
    '/
    '/  Description:    Provides the two sub-pages within the Display page
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Name As String

Dim mPage As BPrefsPage
Dim mStyle As TStyle
Dim mThemes As BTagList

Dim WithEvents theDefaultsPanel As TStyleDefaultsPanel
Attribute theDefaultsPanel.VB_VarHelpID = -1

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl

    If mPage.Find("default_duration", pc) Then _
        pc.SetValue g_ConfigGet("default_duration")


'    ' /* 2.4.2: global shadowing */
'
'Dim ppd As BPackedData
'
'    Set ppd = New BPackedData
'    ppd.SetTo g_ConfigGet("global_shadow_list")
'
'Dim sz As String
'Dim i As Long
'Dim j As Long
'
'    If mPage.Find("global_shadow_list", pc) Then
'        j = prefskit_GetItemCount(pc)
'        If j Then
'            For i = 1 To j
'                sz = Replace$(prefskit_GetItem(pc, "text", i), ": ", "/")
'                If ppd.ValueOf(sz) = "1" Then _
'                    prefskit_SetItem pc, i, "checked", 1&
'
'            Next i
'        End If
'
'        g_StyleRoster.SetNonWindowStyleIcons pc
'
'    End If

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg
Dim sz As String
Dim b As Boolean
Dim i As Long

    Select Case Name

    Case "def"

        ' /* ~~~~~~~~ appearance ~~~~~~~~ */

        uGetIconThemes

        With mPage

            ' /* general */

            .Add new_BPrefsControl("banner", "", "General")

            Set pm = New CTempMsg
            pm.Add "min", 1&
            pm.Add "max", 100&
            pm.Add "freq", 10&
            pm.Add "suffix", "%"
            .Add new_BPrefsControl("fancyslider", "global_opacity", "", "Opacity:", g_ConfigGet("global_opacity"), pm)

            .Add new_BPrefsControl("fancycycle", "style.overflow_limit", "5 lines|7 lines|15 lines", "Limit text to:", uOverflowLimit())
            .Add new_BPrefsControl("fancytoggle2", "use_style_icons", "Use style-provided icons (where possible) over Snarl icons?", "Other:", g_ConfigGet("use_style_icons"))
            .Add new_BPrefsControl("fancytoggle2", "auto_detect_url", "Automatically detect URLs in a notification?", "", g_ConfigGet("auto_detect_url"))
            
            .Add new_BPrefsControl("banner", "", "Style Defaults")
'            .Add new_BPrefsControl("label", "", "Certain styles (such as the built-in Snarl display)")
            .Add new_BPrefsControl("fancybutton2", "show_style_defaults", "Configure...")

            ' /* heads-up display */

            .Add new_BPrefsControl("banner", "", "Notification Overlay")
            .Add new_BPrefsControl("fancytoggle2", "show_timestamp", "Show timestamp while pointer is over notification?", "", g_ConfigGet("show_timestamp"))
            .Add new_BPrefsControl("fancycycle", "icon_theme", uGetIconThemeList(), "Emblems:", CStr(uIconThemeIndex()), , (mThemes.CountItems > 1))

            ' /* others */

'            .Add new_BPrefsControl("fancytoggle2", "no_callback_urls", "Block URLs as default callback?", "", g_ConfigGet("no_callback_urls"))
'            .Add new_BPrefsControl("banner", "", "Other")


        End With

    Case "lay"

        ' /* ~~~~~~~~ layout ~~~~~~~~ */

        mPage.Add new_BPrefsControl("banner", "", "General")

        b = uGetScreens(sz)
        If Not b Then _
            sz = "<Primary Screen>|"

        mPage.Add new_BPrefsControl("fancycycle", "default_screen", sz, "Screen:", g_ConfigGet("default_screen"), , b)
        mPage.Add new_BPrefsControl("fancycycle", "default_position", "Top Left#?0|Top Right#?1|Bottom Left#?2|Bottom Right#?3", "Position:", g_ConfigGet("default_position"))

        Set pm = New CTempMsg
        pm.Add "min", -20&
        pm.Add "max", 20&
        pm.Add "freq", 5&
        mPage.Add new_BPrefsControl("fancyslider", "margin_spacing", "", "Spacing:", g_ConfigGet("margin_spacing"), pm)

        ' /* dropshadow */

        mPage.Add new_BPrefsControl("banner", "", "Drop Shadow")

'        mPage.Add new_BPrefsControl("label", "", "Shadow:")

'        mPage.Add new_BPrefsControl("fancytoggle2", "ft>use_dropshadow", "Add dropshadow to notifications?", "", IIf(gPrefs.use_dropshadow, "1", "0"), pm)
    
        Set pm = New CTempMsg
        pm.Add "min", 0&
        pm.Add "max", 100&
        pm.Add "freq", 10&
        pm.Add "suffix", "%"
        mPage.Add new_BPrefsControl("fancyslider", "dropshadow_strength", "", "Strength:", g_ConfigGet("dropshadow_strength"), pm, gPrefs.use_dropshadow)
    
        Set pm = New CTempMsg
        pm.Add "min", 0&
        pm.Add "max", 20&
        pm.Add "freq", 2&
        mPage.Add new_BPrefsControl("fancyslider", "dropshadow_size", "", "Size:", g_ConfigGet("dropshadow_size"), pm, gPrefs.use_dropshadow)





    ' /* ¬¬¬¬¬¬¬¬¬¬ notification behaviour ¬¬¬¬¬¬¬¬¬¬ */

    Case "vis"

        With mPage

            .Add new_BPrefsControl("banner", "", "General")

            i = Val(g_ConfigGet("default_duration"))
            i = i \ 5
            i = i * 5

            Select Case i
                Case 0:     i = 1
                Case 5:     i = 2
                Case 10:    i = 3
                Case 30:    i = 4
                Case 60:    i = 5
                Case Else:  i = 3

            End Select

'        .Add new_BPrefsControl("fancycycle", "default_duration", "None (sticky)|5 seconds|10 seconds|30 seconds|1 minute", "Duration:", CStr(i))

            Set pm = New CTempMsg
            pm.Add "min", 0&
            pm.Add "max", 60&
            pm.Add "freq", 5&
            pm.Add "label_size", 50&
            .Add new_BPrefsControl("fancyslider", "default_duration", "", "Duration:", CStr(-1), pm)
'        .Add new_BPrefsControl("label", "", "The above settings can be overridden on a per-class basis by configuring the class in the 'Apps' tab.")


            ' /* sounds */

            .Add new_BPrefsControl("banner", "", "Sounds")
            .Add new_BPrefsControl("fancytoggle2", "enable_sounds", "Enable sounds?", "", g_ConfigGet("enable_sounds"))
            .Add new_BPrefsControl("fancytoggle2", "use_style_sounds", "Use style-provided sounds (if available)?", , g_ConfigGet("use_style_sounds"))
            .Add new_BPrefsControl("fancytoggle2", "prefer_style_sounds", "Style-provided sounds take priority over app-provided sounds?", , g_ConfigGet("prefer_style_sounds"))
'        .Add new_BPrefsControl("label", "", "If enabled Snarl will use the Style sounds even if an application specifies its own.")
            
            .Add new_BPrefsControl("banner", "", "Default Sounds")
            .Add new_BPrefsControl("fancyfilepicker", "default_notification_sound", g_ConfigGet("default_notification_sound"), "Normal:")
            .Add new_BPrefsControl("fancyfilepicker", "default_priority_sound", g_ConfigGet("default_priority_sound"), "Priority:")
            .Add new_BPrefsControl("label", "", "Leave blank for no default sound or specify a system sound by prefixing with '!'.  System sounds are stored HKEY_CURRENT_USER\AppEvents\EventLabels.")


        End With





    Case "thu"
        
'        ' /* Pass-Through settings */
'
'Const ITEM_HEIGHT = 30
'Const NUM_ITEMS = 10
'
'        With mPage
'            .SetMargin 0
'            .Add new_BPrefsControl("label", "", "Direct notification content to the following:")
'
'            Set pm = New CTempMsg
'            pm.Add "checkboxes", 1
'            pm.Add "item-height", CLng(ITEM_HEIGHT)
'            Set pc = new_BPrefsControl("listbox", "global_shadow_list", g_StyleRoster.GetNonWindowStyleList(), "", "", pm)
'            pc.SizeTo 0, (NUM_ITEMS * ITEM_HEIGHT) + 2
'            .Add pc
'
'        End With


    Case "adv"

        ' /* Sounds settings */

'        With mPage
''            .SetMargin 24
'
'            ' /* sounds */
'
'            .Add new_BPrefsControl("banner", "", "Sounds")
'            .Add new_BPrefsControl("fancytoggle2", "enable_sounds", "Enable sounds?", "", g_ConfigGet("enable_sounds"))
'            .Add new_BPrefsControl("fancytoggle2", "use_style_sounds", "Use style-provided sounds (if available)?", , g_ConfigGet("use_style_sounds"))
'            .Add new_BPrefsControl("fancytoggle2", "prefer_style_sounds", "Style-provided sounds take priority over app-provided sounds?", , g_ConfigGet("prefer_style_sounds"))
''        .Add new_BPrefsControl("label", "", "If enabled Snarl will use the Style sounds even if an application specifies its own.")
'
'            .Add new_BPrefsControl("banner", "", "Default Sounds")
'            .Add new_BPrefsControl("fancyfilepicker", "default_notification_sound", g_ConfigGet("default_notification_sound"), "Normal:")
'            .Add new_BPrefsControl("fancyfilepicker", "default_priority_sound", g_ConfigGet("default_priority_sound"), "Priority:")
'            .Add new_BPrefsControl("label", "", "Leave blank for no default sound or specify a system sound by prefixing with '!'.  System sounds are stored HKEY_CURRENT_USER\AppEvents\EventLabels.")
'
'        End With

    Case "sty"
    
        ' /* style defaults */

        With mPage
            .Add new_BPrefsControl("banner", "", "Notification Title")
            .Add new_BPrefsControl("fancyfontpicker", "style:title-font", gStyleDefaults.ValueOf("title-font"), "Font:", gStyleDefaults.ValueOf("title-font"))
            .Add new_BPrefsControl("fancycolourpicker", "style:title-colour", , "Colour:", gStyleDefaults.ValueOf("title-colour"))
'            .Add new_BPrefsControl("fancycycle", "style:title-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("title-weight"))
            Set pm = New CTempMsg
            With pm
                .Add "min", 0&
                .Add "max", 100&
                .Add "freq", 10&
                .Add "suffix", "%"

            End With
            .Add new_BPrefsControl("fancyslider", "style:title-opacity", "", "Opacity:", gStyleDefaults.ValueOf("title-opacity"), pm)

            .Add new_BPrefsControl("banner", "", "Notification Text")
            .Add new_BPrefsControl("fancyfontpicker", "style:text-font", gStyleDefaults.ValueOf("text-font"), "Font:", gStyleDefaults.ValueOf("text-font"))
            .Add new_BPrefsControl("fancycolourpicker", "style:text-colour", , "Colour:", gStyleDefaults.ValueOf("text-colour"))
'            .Add new_BPrefsControl("fancycycle", "style:text-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("text-weight"))
            Set pm = New CTempMsg
            With pm
                .Add "min", 0&
                .Add "max", 100&
                .Add "freq", 10&
                .Add "suffix", "%"

            End With
            .Add new_BPrefsControl("fancyslider", "style:text-opacity", "", "Opacity:", gStyleDefaults.ValueOf("text-opacity"), pm)

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("label", "", "These settings only apply to the default style, however other styles may also (and are encouraged to) support these settings.")

        End With

    End Select



        ' /* spacing */

'        mPage.Add new_BPrefsSeparator()
'        mPage.Add new_BPrefsControl("banner", "", "Other")
'        mPage.Add new_BPrefsCombo("c>icon_theme", "", "Icon Theme:", , 28)
'        uUpdateIconThemes

        ' /* font smoothing */

'        mPage.Add new_BPrefsControl("banner", "", "Font Smoothing")
'        mPage.Add new_BPrefsControl("fancycycle", "font_smoothing", "MelonType#?0|None#?1|Antialiased#?2|ClearType#?3|Windows Default#?4", "Mode:", CStr(gPrefs.font_smoothing + 1))
'
'        Set pm = New CTempMsg
'        pm.Add "min", 1&
'        pm.Add "max", 100&
'        pm.Add "freq", 10&
'        pm.Add "suffix", "%"
'        mPage.Add new_BPrefsControl("fancyslider", "melontype_contrast", "", "Contrast:", CStr(gPrefs.melontype_contrast), pm)


End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim sz As String

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

'    If g_SafeLeftStr(Control.GetName, 6) = "style:" Then
'        ' /* style default */
'        sz = g_SafeRightStr(Control.GetName, Len(Control.GetName) - 6)
'        gStyleDefaults.Update sz, Value
'        gStyleDefaults.Save
'        Exit Sub
'
'    End If


Dim fUpdate As Boolean
Dim pc As BControl
Dim pm As CTempMsg
Dim dw As Long
Dim i As Long

    Select Case Control.GetName()

    ' /* notification settings */

'    Case "font_smoothing"
'        gPrefs.font_smoothing = Val(Control.GetValue()) - 1
'        If mPage.Find("melontype_contrast", pc) Then _
'            pc.SetEnabled (Val(Value) = 1)
'
'        fUpdate = True


    Case "ft>use_dropshadow"
        gPrefs.use_dropshadow = (Val(Value) <> 0)
        If mPage.Find("dropshadow_strength", pc) Then _
            pc.SetEnabled (Val(Value) <> 0)

        If mPage.Find("dropshadow_size", pc) Then _
            pc.SetEnabled (Val(Value) <> 0)

        fUpdate = True


    Case "default_style"
        If (g_StyleRoster Is Nothing) Then _
            Exit Sub

'        mStyle.Info.Name = ""
'        g_StyleRoster.StyleAt Val(Value), mStyle
'        uUpdateSchemeList


    Case "default_position", "melontype_contrast", "global_opacity", "margin_spacing", "dropshadow_strength", "dropshadow_size", _
         "default_notification_sound", "default_priority_sound", "enable_sounds", "use_style_sounds", "prefer_style_sounds", _
         "use_style_icons", "show_timestamp", "auto_detect_url", "default_screen"

        g_ConfigSet Control.GetName, Value


    Case "default_scheme"
        ' /* 'scheme' combo on 'visuals' page */
        Set pm = New CTempMsg
        pm.Add "what", "text"

        If Control.Page.Find("default_style", pc) Then
            ' /* get the style name */
            sz = pc.DoExCmd(B_GET_ITEM, pm)
            If sz <> "" Then
                ' /* append the style name */
                sz = LCase$(sz & "/" & Control.DoExCmd(B_GET_ITEM, pm))
                ' /* set it */
                g_ConfigSet "default_style", sz
                Debug.Print "*** default scheme changed to '" & sz & "' ***"

            End If
        End If


    Case "default_duration"
        Select Case Val(Value)
        Case 0
            Control.SetText "Sticky"

        Case 1
            Control.SetText "1 sec"

        Case Else
            Control.SetText Value & " secs"

        End Select

        g_ConfigSet Control.GetName, Value


    Case "icon_theme"
        g_ConfigSet Control.GetName, mThemes.TagAt(Val(Value)).Value
        g_LoadIconTheme


    Case "style.overflow_limit"
        Select Case Value
        Case "1"
            g_ConfigSet Control.GetName, "5"

        Case "2"
            g_ConfigSet Control.GetName, "7"

        Case "3"
            g_ConfigSet Control.GetName, "15"

        End Select

    End Select


    If fUpdate Then _
        g_WriteConfig

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName = "show_style_defaults" Then
        Set theDefaultsPanel = New TStyleDefaultsPanel
        theDefaultsPanel.Go frmAbout.PanelhWnd

    End If

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
Dim ppd As BPackedData
Dim i As Long
Dim j As Long

    Select Case Control.GetName

    Case "global_shadow_list"
        If Notification = "checked" Then
            ' /* get item name */
            Set ppd = New BPackedData
            j = prefskit_GetItemCount(Control)
            If j Then
                For i = 1 To j
                    If prefskit_GetItem(Control, "checked", i) = "1" Then _
                        ppd.Add Replace$(prefskit_GetItem(Control, "text", i), ": ", "/"), "1"

                Next i

            End If

            g_ConfigSet "global_shadow_list", ppd.AsString

        End If

    End Select

End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 100

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

Private Sub uUpdateStyleList()
Dim pc As BControl

    If (Not (mPage.Find("default_style", pc))) Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

Dim sz As String
Dim b As Boolean

    pc.SetText g_StyleRoster.GetStyleList(b)
    pc.SetEnabled b
'    pc.SetValue "1"

End Sub

Private Sub uUpdateSchemeList()
Dim pc As BControl

    If (Not (mPage.Find("default_scheme", pc))) Or (mStyle.Name = "") Or (g_StyleRoster Is Nothing) Then _
        Exit Sub

    ' /* set it */

'    pc.SetText g_StyleRoster.GetSchemeList(mStyle)
'    pc.SetEnabled True
'    pc.SetValue "1"

End Sub




'Private Sub uUpdateIconThemes()
'
'    If (mPage Is Nothing) Then _
'        Exit Sub
'
'Dim pc As BControl
'Dim sz As String
'Dim i As Long
'Dim n As Long
'
'    If Not (mPage.Find("c>icon_theme", pc)) Then _
'        Exit Sub
'
'    sz = "(Default Theme)|"
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            If LCase$(gIconTheme(i).Name) = LCase$(gPrefs.icon_theme) Then _
'                n = i
'
'            sz = sz & gIconTheme(i).Name & "|"
'
'        Next i
'
'    End If
'
'    pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
'    pc.SetValue CStr(n + 1)
'
'    ' /* set the icons */
'
'Dim pm As CTempMsg
'
'    Set pm = New CTempMsg
'    pm.Add "index", 1
'    pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\snarl.png"
'    pc.DoExCmd B_SET_ITEM, pm
'
'
'    If gIconThemes Then
'        For i = 1 To gIconThemes
'            Set pm = New CTempMsg
'            pm.Add "index", i + 1
'            pm.Add "image-file", gIconTheme(i).IconFile
'            pc.DoExCmd B_SET_ITEM, pm
'
'        Next i
'
'    End If
'
'
'End Sub

Private Function uGetIconThemeList() As String

    If (mThemes Is Nothing) Then
        g_Debug "TDisplaySubPage.uGetIconThemeList(): list is null", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pt As BTagItem

    With mThemes
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            uGetIconThemeList = uGetIconThemeList & pt.Name & "|"

        Loop
        
    End With

'    uGetIconThemeList = g_SafeLeftStr(uGetIconThemeList, Len(uGetIconThemeList) - 1)
    g_Debug "TDisplaySubPage.uGetIconThemeList(): list is: " & uGetIconThemeList

End Function

Private Function uIconThemeIndex() As Long
Dim sz As String

    uIconThemeIndex = 1

    sz = g_ConfigGet("icon_theme")
    If sz = "" Then _
        Exit Function

    If (mThemes Is Nothing) Then
        g_Debug "TDisplaySubPage.uIconThemeIndex(): list is null", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pt As BTagItem
Dim i As Long

    With mThemes
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            i = i + 1
            If pt.Value = sz Then
                uIconThemeIndex = i
                Exit Function
            
            End If

        Loop

    End With

End Function

Private Function uTranslateIconTheme(ByVal Path As String) As String

    Path = g_RemoveExtension(g_FilenameFromPath(Path))

Dim s() As String

    s = Split(Path, "@")
    If UBound(s) = 1 Then
        uTranslateIconTheme = s(1) & " by " & s(0)

    Else
        uTranslateIconTheme = Path

    End If

End Function

Private Sub uGetIconThemes()
Dim sz As String

    g_Debug "TDisplaySubPage.uGetIconThemes()", LEMON_LEVEL_PROC_ENTER

    ' /* get icon themes */

    Set mThemes = new_BTagList()
    mThemes.Add new_BTagItem("Default", "")

    With New CFolderContent2
        If .SetTo(g_MakePath(App.Path) & "etc\icons") Then
            .Rewind
            Do While .GetNextFile(sz)
                If (g_IsFolder(sz)) And (g_GetExtension(sz, True) = "icons") Then _
                    mThemes.Add new_BTagItem(uTranslateIconTheme(sz), g_RemoveExtension(g_FilenameFromPath(sz)))

            Loop

        Else
            g_Debug "couldn't find \etc\icons", LEMON_LEVEL_CRITICAL

        End If

    End With

    If mThemes.CountItems < 2 Then _
        g_Debug "Only the default theme available", LEMON_LEVEL_WARNING

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Function uOverflowLimit() As String
Dim n As Long

    ' /* translate the number of lines into a cyclebutton value */

    n = g_SafeLong(g_ConfigGet("style.overflow_limit"))

    Select Case n
    Case 5
        uOverflowLimit = "1"

    Case 15
        uOverflowLimit = "3"

    Case Else
        uOverflowLimit = "2"

    End Select

End Function

Private Function uGetScreens(ByRef List As String) As Boolean
Dim sz As String
Dim pmi As T_MONITOR_INFO
Dim c As Long
Dim i As Long

    c = g_CountMonitors()
    If c > 0 Then
        uGetScreens = (c > 1)

        For i = 1 To c
            g_GetMonitorInfo i, pmi
            sz = sz & "Screen " & CStr(i) & IIf(pmi.IsPrimary, " (Primary)", "") & "|"

        Next i

        List = sz

    End If

End Function

Private Sub theDefaultsPanel_Closed()

    Set theDefaultsPanel = Nothing

End Sub
