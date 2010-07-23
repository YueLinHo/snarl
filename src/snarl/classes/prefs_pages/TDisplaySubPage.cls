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

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
Dim pc As BControl

    If mPage.Find("default_duration", pc) Then _
        pc.SetValue g_ConfigGet("default_duration")

End Sub

Private Sub KPrefsPage_Attached()
Dim pm As CTempMsg
Dim i As Long

    Select Case Name

    ' /* ¬¬¬¬¬¬¬¬¬¬ appearance & layout ¬¬¬¬¬¬¬¬¬¬ */

    Case "def"

        ' /* ~~~~~~~~ inside the notification ~~~~~~~~ */

        mPage.Add new_BPrefsControl("banner", "", "Inside")

        ' /* opacity */

        Set pm = New CTempMsg
        pm.Add "min", 1&
        pm.Add "max", 100&
        pm.Add "freq", 10&
        pm.Add "suffix", "%"
        mPage.Add new_BPrefsControl("fancyslider", "global_opacity", "", "Opacity:", g_ConfigGet("global_opacity"), pm)

        ' /* icon theme */

        mPage.Add new_BPrefsControl("fancytoggle2", "use_style_icons", "Use style-provided icons (where possible) over Snarl icons?", "Icons:", g_ConfigGet("use_style_icons"))

        ' /* timestamp */

        mPage.Add new_BPrefsControl("fancytoggle2", "show_timestamp", "Show timestamp while pointer is over notification?", "", g_ConfigGet("show_timestamp"))

        ' /* ~~~~~~~~ outside the notification ~~~~~~~~ */


        mPage.Add new_BPrefsControl("banner", "", "Outside")

        mPage.Add new_BPrefsControl("fancycycle", "default_position", "Top Left#?0|Top Right#?1|Bottom Left#?2|Bottom Right#?3", "Position:", g_ConfigGet("default_position"))

        Set pm = New CTempMsg
        pm.Add "min", -20&
        pm.Add "max", 20&
        pm.Add "freq", 5&
        mPage.Add new_BPrefsControl("fancyslider", "margin_spacing", "", "Spacing:", g_ConfigGet("margin_spacing"), pm)

        ' /* dropshadow */

        mPage.Add new_BPrefsControl("label", "", "Shadow:")

'        mPage.Add new_BPrefsControl("fancytoggle2", "ft>use_dropshadow", "Add dropshadow to notifications?", "", IIf(g_Prefs.use_dropshadow, "1", "0"), pm)
    
        Set pm = New CTempMsg
        pm.Add "min", 0&
        pm.Add "max", 100&
        pm.Add "freq", 10&
        pm.Add "suffix", "%"
        mPage.Add new_BPrefsControl("fancyslider", "dropshadow_strength", "", "Strength:", g_ConfigGet("dropshadow_strength"), pm, g_Prefs.use_dropshadow)
    
        Set pm = New CTempMsg
        pm.Add "min", 0&
        pm.Add "max", 20&
        pm.Add "freq", 2&
        mPage.Add new_BPrefsControl("fancyslider", "dropshadow_size", "", "Size:", g_ConfigGet("dropshadow_size"), pm, g_Prefs.use_dropshadow)





    ' /* ¬¬¬¬¬¬¬¬¬¬ notification behaviour ¬¬¬¬¬¬¬¬¬¬ */

    Case "vis"

        mPage.Add new_BPrefsControl("banner", "", "General")

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

'        mPage.Add new_BPrefsControl("fancycycle", "default_duration", "None (sticky)|5 seconds|10 seconds|30 seconds|1 minute", "Duration:", CStr(i))

        Set pm = New CTempMsg
        pm.Add "min", 0&
        pm.Add "max", 60&
        pm.Add "freq", 5&
        pm.Add "label_size", 50&
        mPage.Add new_BPrefsControl("fancyslider", "default_duration", "", "Duration:", CStr(-1), pm)
'        mPage.Add new_BPrefsControl("label", "", "The above settings can be overridden on a per-class basis by configuring the class in the 'Apps' tab.")

        ' /* text overflow */

'        mPage.Add new_BPrefsControl("label", "", "", "Text Limiting:")
'        mPage.Add new_BPrefsControl("fancycycle", "text_overflow", "Any|3 Lines|5 Lines|7 Lines|9 Lines", "", "1")

        ' /* sounds */

        mPage.Add new_BPrefsControl("banner", "", "Sounds")
        mPage.Add new_BPrefsControl("fancytoggle2", "enable_sounds", "Enable sounds?", "", g_ConfigGet("enable_sounds"))
        mPage.Add new_BPrefsControl("fancytoggle2", "use_style_sounds", "Use style-provided sounds (if available)?", , g_ConfigGet("use_style_sounds"))
        mPage.Add new_BPrefsControl("fancytoggle2", "prefer_style_sounds", "Style-provided sounds take priority over app-provided sounds?", , g_ConfigGet("prefer_style_sounds"))
'        mPage.Add new_BPrefsControl("label", "", "If enabled Snarl will use the Style sounds even if an application specifies its own.")

        mPage.Add new_BPrefsControl("label", "", "Default sounds:", "")
        mPage.Add new_BPrefsControl("fancyfilepicker", "default_notification_sound", g_ConfigGet("default_notification_sound"), "Normal:")
        mPage.Add new_BPrefsControl("fancyfilepicker", "default_priority_sound", g_ConfigGet("default_priority_sound"), "Priority:")
        mPage.Add new_BPrefsControl("label", "", "Leave blank for no default sound or specify a Windows sound by prefixing with '!'.  Sounds are stored HKEY_CURRENT_USER\AppEvents\EventLabels.")


        ' /* spacing */

'        mPage.Add new_BPrefsSeparator()
'        mPage.Add new_BPrefsControl("banner", "", "Other")
'        mPage.Add new_BPrefsCombo("c>icon_theme", "", "Icon Theme:", , 28)
'        uUpdateIconThemes

        ' /* font smoothing */

'        mPage.Add new_BPrefsControl("banner", "", "Font Smoothing")
'        mPage.Add new_BPrefsControl("fancycycle", "font_smoothing", "MelonType#?0|None#?1|Antialiased#?2|ClearType#?3|Windows Default#?4", "Mode:", CStr(g_Prefs.font_smoothing + 1))
'
'        Set pm = New CTempMsg
'        pm.Add "min", 1&
'        pm.Add "max", 100&
'        pm.Add "freq", 10&
'        pm.Add "suffix", "%"
'        mPage.Add new_BPrefsControl("fancyslider", "melontype_contrast", "", "Contrast:", CStr(g_Prefs.melontype_contrast), pm)



    Case "cd"

'        ' /* class defaults */
'
'        i = g_Prefs.default_duration
'        i = i \ 5
'        i = i * 5
'
'        Select Case i
'        Case 0
'            i = 1
'
'        Case 5
'            i = 2
'
'        Case 10
'            i = 3
'
'        Case 30
'            i = 4
'
'        Case 60
'            i = 5
'
'        Case Else
'            i = 3
'
'        End Select
'
''        mPage.Add new_BPrefsControl("fancycycle", "default_duration", "None (sticky)|5 seconds|10 seconds|30 seconds|1 minute", "Duration:", CStr(i))
'
'        mPage.Add new_BPrefsControl("banner", "", "General Settings")
'
'        Set pm = New CTempMsg
'        pm.Add "min", 0&
'        pm.Add "max", 60&
'        pm.Add "freq", 5&
'        pm.Add "label_size", 50&
'        mPage.Add new_BPrefsControl("fancyslider", "default_duration", "", "Duration:", CStr(-1), pm)
'
'        mPage.Add new_BPrefsControl("fancycycle", "default_position", "Top Left#?0|Top Right#?1|Bottom Left#?2|Bottom Right#?3", "Position:", CStr(g_Prefs.default_position + 1))


    End Select

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim fUpdate As Boolean
Dim pc As BControl
Dim pm As CTempMsg
Dim sz As String
Dim dw As Long
Dim i As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

    ' /* notification settings */

    Case "font_smoothing"
        g_Prefs.font_smoothing = Val(Control.GetValue()) - 1
        If mPage.Find("melontype_contrast", pc) Then _
            pc.SetEnabled (Val(Value) = 1)

        fUpdate = True


    Case "ft>use_dropshadow"
        g_Prefs.use_dropshadow = (Val(Value) <> 0)
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
         "use_style_icons", "show_timestamp"

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

    End Select


    If fUpdate Then _
        g_WriteConfig

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
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
'            If LCase$(gIconTheme(i).Name) = LCase$(g_Prefs.icon_theme) Then _
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



