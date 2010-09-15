VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           extension.cls
    '/
    '/  Description:    Extension main object
    '/
    '/  Created:        3-Apr-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* our constants */

Const EXT_NAME = "Clock+"
Const EXT_DATE = "15-Sep-2010"
Const EXT_REL = "2.02"

Const CLASS_HOUR_CHIME = "hrch"
Const CLASS_COUNTDOWN_COMPLETE = "ccom"
'Const CLASS_REMINDERS = "outr"

    ' /*
    '
    '   This class provides information about the Extension itself.  It _must_ implement
    '   both MWndProcSink and MVersionInfo otherwise it won't get loaded by Snarl.  It's
    '   worth noting the the MWndProcSink interface here is not used as a 'normal' window
    '   procedure; it's simply a useful interface that Snarl can send messages to.
    '
    '   See the individual methods and properties for more detailed information.
    '
    '   New for Snarl V40 you can also implement MSimpleEnum which can be used to retrieve more
    '   information about the extension.  Only MSimpleEnum.FindItem() is currently used - all
    '   other functions should return zero or False for now - with the following queries:
    '
    '       'path' - return current extension path (usually App.Path)
    '       'copyright' - return copyright information
    '       'url' - return fully formatted URL to website (e.g. 'http://wwww.fullphat.net/')
    '
    '
    '   IMPORTANT: This is a fully functional Snarl extension - if you're planning on
    '   writing your own extension you should use the template extension in the Snarl SDK.
    '
    ' */

Dim mToken As Long

Dim mLastErr As String
Dim mConfigFile As CConfFile

Dim mPanel As BPrefsPanel           ' // our prefs panel
Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)

    ' /* hourly chime */
Dim mChimeId As Long
Dim mLastHour As Integer

    ' /* timer */
Dim mTimerId As Long
Dim mTimer As Date
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Implements KPrefsPanel              ' // we're a panel handler
Implements KPrefsPage               ' // and a page handler
Implements ISnarlExtension

Private Sub Class_Initialize()

    l3OpenLog "%APPDATA%\" & EXT_NAME & ".log"
    g_Debug EXT_NAME & " running from '" & g_MakePath(App.Path) & "'"
    g_Debug ""

End Sub

Private Sub Class_Terminate()

'    lemonCloseLog EXT_NAME

End Sub

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage
Dim pm As CTempMsg
Dim sz As String
Dim pc As BControl


    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle EXT_NAME & " Preferences"
            .SetWidth 380
            .SetWindow 1                    ' // so calling Go() doesn't show the window

            Set pPage = new_BPrefsPage("General", load_image_obj(g_MakePath(App.Path) & "general.png"), Me)

            ' /* General */

            With pPage
                .SetMargin 90
                .Add new_BPrefsControl("banner", "", "Startup")
            
'            Set pm = New CTempMsg
'            pm.Add "inset_by", CLng(Fix(.GetMargin() / 2))

                .Add new_BPrefsControl("fancytoggle2", "show_time_on_load", "Show current time?", , IIf(gConfig.show_time_on_load, "1", "0"))
                .Add new_BPrefsControl("fancytoggle2", "show_cookie_on_load", "Display random cookie?", , IIf(gConfig.show_cookie_on_load, "1", "0"))

                .Add new_BPrefsControl("banner", "", "Cookies")
                .Add new_BPrefsControl("fancytoggle2", "show_cookie", "Include random cookie in notification?", , IIf(gConfig.show_cookie, "1", "0"))

                sz = g_GetPath(gConfig.cookie_file)
                If sz <> "" Then
                    Set pm = New CTempMsg
                    pm.Add "path", sz

                End If

                .Add new_BPrefsControl("fancyfilepicker", "cookie_file", gConfig.cookie_file, "Cookie file:", , pm)

                .Add new_BPrefsControl("banner", "", "About " & EXT_NAME)
                .Add new_BPrefsControl("label", "", "V" & CStr(App.Major) & "." & CStr(App.Revision), "Version:")
                .Add new_BPrefsControl("label", "", App.LegalCopyright)
                .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net", "Website:")

            End With

            .AddPage pPage

            ' /* Timer */

            Set pPage = new_BPrefsPage("Timer", load_image_obj(g_MakePath(App.Path) & "timer.png"), Me)
            With pPage
                .SetMargin 90
                .Add new_BPrefsControl("banner", "", "Duration")
                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 23&
                pm.Add "freq", 2&
                .Add new_BPrefsControl("fancyslider", "timer_hours", , "Hours:", CStr(gConfig.timer_hours), pm)

                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 59&
                pm.Add "freq", 5&
                .Add new_BPrefsControl("fancyslider", "timer_mins", , "Minutes:", CStr(gConfig.timer_mins), pm)

                .Add new_BPrefsControl("fancytoggle", "start_stop_timer", "Paused|Running")
                .Add new_BPrefsControl("label", "timer_current", _
                                       Format$(gConfig.timer_hours, "00") & ":" & Format$(gConfig.timer_mins, "00") & ":00", _
                                       "Remaining:")


                .Add new_BPrefsControl("banner", "", "Options")
'            .Add new_BPrefsControl("fancytoggle2", "custom_timer_icon", "Use custom icon?", , IIf(gConfig.custom_timer_icon, "1", "0"))
                .Add new_BPrefsControl("fancyfilepicker", "timer_icon", gConfig.timer_icon, "Icon:", , pm)

                .Add new_BPrefsControl("fancyedit", "timer_message", gConfig.timer_message, "Message:", , pm)
                .Add new_BPrefsControl("label", "", "(Leave either field blank to use the default)")

                .Add new_BPrefsControl("fancytoggle2", "timer_10sec_warning", "Show notification during last 10 seconds?", "Other:", IIf(gConfig.timer_10sec_warning, "1", "0"))
                .Add new_BPrefsControl("fancytoggle2", "show_meter", "Meter-friendly?", , IIf(gConfig.show_meter, "1", "0"))
                .Add new_BPrefsControl("label", "", "Enabling this option causes " & EXT_NAME & " to include text in the notification which some styles can display graphically.")

            End With

            .AddPage pPage

            ' /* formats */

            Set pPage = new_BPrefsPage("Formats", load_image_obj(g_MakePath(App.Path) & "format.png"), Me)
            With pPage
                .SetMargin 90
                .Add new_BPrefsControl("banner", "", "General Time Format")

                Set pm = New CTempMsg
                pm.Add "inset_by", CLng(Fix(.GetMargin() / 2))

'                .Add new_BPrefsControl("fancytoggle2", "show_as_words", "Show as text?", , IIf(gConfig.show_as_words, "1", "0"), pm)
                .Add new_BPrefsControl("fancycycle", "format_type", "Numeric|Text|YYYYMMDDHHMMSS", , CStr(gConfig.format_type + 1))

                .Add new_BPrefsControl("banner", "", "Numeric Time Format")
                .Add new_BPrefsControl("fancytoggle2", "use_12_hour_clock", "Use 12 hour clock?", , IIf(gConfig.use_12_hour_clock, "1", "0"), pm)
                .Add new_BPrefsControl("fancytoggle2", "two_digit_hour", "Two digit hour?", , IIf(gConfig.two_digit_hour, "1", "0"), pm)
                .Add new_BPrefsControl("fancytoggle2", "show_seconds", "Show seconds?", , IIf(gConfig.show_seconds, "1", "0"), pm)
                .Add new_BPrefsControl("fancytoggle2", "show_am_pm", "Show AM/PM?", , IIf(gConfig.show_am_pm, "1", "0"), pm)
    
    '            .Add new_BPrefsControl("fancytoggle2", "custom_am_pm", "Custom AM/PM values?", , IIf(1, "1", "0"), pm, Not gConfig.show_as_words)

                .Add new_BPrefsControl("banner", "", "Custom Mnemonics")
                Set pc = new_BPrefsControl("fancyedit", "am_val", gConfig.am_val, "AM:")
                pc.SizeTo 48, pc.Height
                .Add pc
    
                Set pc = new_BPrefsControl("fancyedit", "pm_val", gConfig.pm_val, "PM:")
                pc.SizeTo 48, pc.Height
                .Add pc

                .Add new_BPrefsControl("separator", "")
    '            .Add new_BPrefsControl("label", "", "These settings affect the time format used in any notifications this extension displays")

                Set pc = new_BPrefsControl("fancybutton2", "test", "Test Settings")
                pc.SizeTo 128, 28
                .Add pc
    
            End With

            .AddPage pPage

            ' /* create it */

            .Go
            g_SetWindowIconToAppResourceIcon .hWnd

        End With

    End If

    ISnarlExtension_GetConfigWindow = mPanel.hWnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "full phat products"
        .Copyright = App.LegalCopyright
        .Date = EXT_DATE
        .Description = App.FileDescription
        .Flags = SNARL_EXTN_IS_CONFIGURABLE Or SNARL_EXTN_WANTS_PULSE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = EXT_NAME
        .Path = App.Path
        .Release = EXT_REL
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long
Dim szPath As String

    On Error Resume Next

    mLastErr = ""

    ' /* does snarl's /etc/ path exist? */

    If Not snext_getuserpath(szPath) Then
        mLastErr = "Couldn't get Snarl user path"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    ' /* defaults */

    With gConfig
        .cookie_file = g_MakePath(App.Path) & "cookies.txt"
        .format_type = 1                    ' // text
        .show_cookie = True
        .show_time_on_load = True
        .use_12_hour_clock = True
        .show_am_pm = True
        .show_seconds = False
        .two_digit_hour = False
        .am_val = "am"
        .pm_val = "pm"

        .timer_icon = ""
        .timer_10sec_warning = True
        .timer_message = ""

        .timer_hours = 0
        .timer_mins = 10
        .timer_running = False

        .show_meter = True
        .show_cookie_on_load = True

    End With

    ' /* load config: if our config exists we load it, otherwise, if
    '    the old SnarlClock config exists, we load that instead */

Dim sz As String

    sz = g_MakePath(szPath) & "etc\.snarlclock2"
    If Not g_Exists(sz) Then
        sz = g_MakePath(szPath) & "extensions\snarlclock\snarlclock.conf"
        If Not g_Exists(sz) Then _
            sz = ""

    End If

Dim szData As String

'    MsgBox "config: " & sz

    If sz <> "" Then
        With New CConfFile
            If .SetTo(sz) Then
                If .FindEntry("cookie_file", szData) Then _
                    gConfig.cookie_file = szData

'                If .FindEntry("reminder_duration", szData) Then _
                    gConfig.reminder_duration = Val(szData)
    
                If .FindEntry("format_type", szData) Then
                    gConfig.format_type = Val(szData)

                Else
                    ' /* set from previous version */
                    If .FindEntry("show_as_words", szData) Then _
                        gConfig.format_type = Val(szData)

                End If

                If .FindEntry("show_cookie", szData) Then _
                    gConfig.show_cookie = (szData = "1")

                If .FindEntry("show_time_on_load", szData) Then _
                    gConfig.show_time_on_load = (szData = "1")

                If .FindEntry("use_12_hour_clock", szData) Then _
                    gConfig.use_12_hour_clock = (szData = "1")

                If .FindEntry("show_am_pm", szData) Then _
                    gConfig.show_am_pm = (szData = "1")
            
                If .FindEntry("show_seconds", szData) Then _
                    gConfig.show_seconds = (szData = "1")
            
                If .FindEntry("two_digit_hour", szData) Then _
                    gConfig.two_digit_hour = (szData = "1")

                If .FindEntry("pm_val", szData) Then _
                    gConfig.pm_val = szData

                If .FindEntry("am_val", szData) Then _
                    gConfig.am_val = szData

                If .FindEntry("debug_mode", szData) Then _
                    gConfig.debug_enabled = (szData = "1")

                If .FindEntry("timer_icon", szData) Then _
                    gConfig.timer_icon = szData

                If .FindEntry("timer_message", szData) Then _
                    gConfig.timer_message = szData

                ' /* 2.0 */

                If .FindEntry("show_meter", szData) Then _
                    gConfig.show_meter = (szData = "1")

            End If

        End With
    End If

    ' /* validate config */

'    If gConfig.reminder_duration < 0 Then
'        gConfig.reminder_duration = 0
'
'    ElseIf gConfig.reminder_duration > 60 Then
'        gConfig.reminder_duration = 60
'
'    End If

    g_LoadCookies

    ' /* ok */

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
Static dTemp As Date

    ' /* hourly reminder */

    If (Minute(Now()) = 0) And (Second(Now()) = 0) And (Hour(Now()) <> mLastHour) Then
        If sn41IsVisible(mChimeId) Then _
            sn41Hide mChimeId

        mChimeId = sn41EZNotify(mToken, _
                                CLASS_HOUR_CHIME, _
                                "Hourly reminder", _
                                IIf(gConfig.format_type = 2, "", "It's now ") & g_GetTimeString(gConfig, True), _
                                -1, _
                                g_MakePath(App.Path) & "hourly.png")

        mLastHour = Hour(Now)

    End If

End Sub

Private Sub ISnarlExtension_Start()

    g_Debug "ISnarlExtension_Start"

    ' /* register */
    mToken = sn41RegisterApp("application/vnd.k23-" & EXT_NAME, EXT_NAME, g_MakePath(App.Path) & "icon.png")

    If mToken = 0 Then _
        Exit Sub

    sn41AddClass mToken, CLASS_HOUR_CHIME, "Hourly chime"
    sn41AddClass mToken, CLASS_COUNTDOWN_COMPLETE, "Countdown complete"
'    sn41AddClass mToken, CLASS_REMINDERS, "Outlook reminders"

    If gConfig.show_time_on_load Then _
        sn41EZNotify mToken, "", EXT_NAME, "It's now " & g_GetTimeString(gConfig, False), -1, g_MakePath(App.Path) & "icon.png"

Dim sz As String

    If (gConfig.show_cookie_on_load) And (g_GetRandomCookie(sz)) Then _
        sn41EZNotify mToken, "", "", sz, -1, g_MakePath(App.Path) & "cookie.png"

End Sub

Private Sub ISnarlExtension_Stop()

    sn41UnregisterApp mToken

End Sub

Private Sub ISnarlExtension_TidyUp()

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    EnableWindow mhWndOwner, -1
    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
Dim pc As BControl

    If mPanel.Find("timer_hours", pc) Then _
        pc.SetEnabled Not gConfig.timer_running

    If mPanel.Find("timer_mins", pc) Then _
        pc.SetEnabled Not gConfig.timer_running

    uUpdateRemaining

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl
Dim nm As Integer
Dim nh As Integer

'    g_Debug "[" & mpanel.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()

    ' /* general */

    Case "show_time_on_load"
        gConfig.show_time_on_load = (Value = "1")
        g_WriteConfig

    Case "show_cookie"
        gConfig.show_cookie = (Value = "1")
        g_WriteConfig

    Case "cookie_file"
        gConfig.cookie_file = Value
        g_WriteConfig

    Case "show_cookie_on_load"
        gConfig.show_cookie_on_load = (Value = "1")
        g_WriteConfig

    ' /* timer */

    Case "start_stop_timer"
        gConfig.timer_running = (Val(Value) <> 0)

        ' /* enabled/disable the sliders */

        If mPanel.Find("timer_hours", pc) Then _
            pc.SetEnabled Not gConfig.timer_running

        If mPanel.Find("timer_mins", pc) Then _
            pc.SetEnabled Not gConfig.timer_running

        If gConfig.timer_running Then

            ' /* resumed/started timer so calculate minutes */

            If mPanel.Find("timer_hours", pc) Then _
                nh = Val(pc.GetValue())

            If mPanel.Find("timer_mins", pc) Then _
                nm = Val(pc.GetValue())

            mTimer = TimeSerial(nh, nm, Second(mTimer))
            If sn41IsVisible(mTimerId) Then _
                sn41Hide mTimerId

            mTimerId = 0
            Set theTimer = new_BTimer(1000)

        Else
            ' /* stop the timer */
            Set theTimer = Nothing

        End If

    Case "timer_hours"
        uUpdateRemaining

    Case "timer_mins"
        uUpdateRemaining

    Case "timer_10sec_warning"
        gConfig.timer_10sec_warning = (Value = "1")
        g_WriteConfig
    
    Case "show_meter"
        gConfig.show_meter = (Value = "1")
        g_WriteConfig
    
    Case "timer_icon"
        gConfig.timer_icon = Value
        g_WriteConfig

    Case "timer_message"
        gConfig.timer_message = Value
        g_WriteConfig

    ' /* formats */

    Case "format_type"
        gConfig.format_type = Val(Value) - 1
        g_WriteConfig
        uEnableTimeSettings (gConfig.format_type = 1)

    Case "use_12_hour_clock"
        gConfig.use_12_hour_clock = (Value = "1")
        g_WriteConfig

    Case "show_am_pm"
        gConfig.show_am_pm = (Value = "1")
        g_WriteConfig

    Case "show_seconds"
        gConfig.show_seconds = (Value = "1")
        g_WriteConfig

    Case "two_digit_hour"
        gConfig.two_digit_hour = (Value = "1")
        g_WriteConfig

    Case "am_val"
        gConfig.am_val = Value

    Case "pm_val"
        gConfig.pm_val = Value

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

'    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "test"
        sn41EZNotify 0, "", "Time format preview", "It's now " & g_GetTimeString(gConfig, False), -1, g_MakePath(App.Path) & "icon.png"

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub uUpdateRemaining()
Dim pc As BControl
Dim nm As Integer
Dim nh As Integer

    ' /* get current hour value */
    If mPanel.Find("timer_hours", pc) Then _
        nh = Val(pc.GetValue)

    ' /* get current minute value */
    If mPanel.Find("timer_mins", pc) Then _
        nm = Val(pc.GetValue)

    ' /* set the label */
    If mPanel.Find("timer_current", pc) Then _
        pc.SetText Format$(nh, "00") & ":" & Format$(nm, "00") & ":00"

    ' /* if the timer isn't running then disable the switch if 0:0 is specified */
    If Not gConfig.timer_running Then
        If mPanel.Find("start_stop_timer", pc) Then _
            pc.SetEnabled Not ((nh = 0) And (nm = 0))

        mTimer = TimeSerial(nh, nm, 0)

    End If

End Sub

Private Sub uEnableTimeSettings(ByVal Enabled As Boolean)
'Dim pc As BControl
'
'    If mPage.Find("use_12_hour_clock", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("two_digit_hour", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("show_seconds", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("show_am_pm", pc) Then _
'        pc.SetEnabled Enabled
'
'    If mPage.Find("custom_am_pm", pc) Then _
'        pc.SetEnabled Enabled


End Sub

Private Sub theTimer_Pulse()
Static i As Long

    ' /* doing a countdown */

    mTimer = DateAdd("s", -1, mTimer)

    If (Hour(mTimer) = 0) And (Minute(mTimer) = 0) Then
        Select Case Second(mTimer)
        Case 0
            ' /* we're done */
            If sn41IsVisible(mTimerId) Then _
                sn41Hide mTimerId               ' // hide 10 second countdown notification

            i = sn41EZNotify(mToken, _
                            CLASS_COUNTDOWN_COMPLETE, _
                            IIf(gConfig.timer_message = "", "Countdown done!", gConfig.timer_message), _
                            "", _
                            -1, _
                            IIf(gConfig.timer_icon = "", g_MakePath(App.Path) & "icon.png", gConfig.timer_icon))

            Set theTimer = Nothing

        Case 10
            ' /* show the T-10 countdown notification? */

            If gConfig.timer_10sec_warning Then _
                mTimerId = sn41EZNotify(mToken, _
                                        CLASS_COUNTDOWN_COMPLETE, _
                                        Second(mTimer) & " second" & IIf(Second(mTimer) = 1, "", "s") & " to go", _
                                        IIf(gConfig.show_meter, "100", ""), _
                                        -1, _
                                        IIf(gConfig.timer_icon = "", g_MakePath(App.Path) & "timer.png", gConfig.timer_icon))


        Case Is < 10
            If gConfig.timer_10sec_warning Then _
                sn41EZUpdate mTimerId, _
                             CStr(Second(mTimer)) & " second" & IIf(Second(mTimer) = 1, "", "s") & " to go", _
                             IIf(gConfig.show_meter, CStr(Second(mTimer) * 10), "")

        End Select

    End If

Static pc As BControl

    If Not (mPanel Is Nothing) Then

        ' /* update the panel controls */

'    g_Debug "uUpdateTimerControls(): " & CStr(mTimer)
        If (Hour(mTimer) = 0) And (Minute(mTimer) = 0) And (Second(mTimer) = 0) Then
            ' /* countdown completed: update toggle */
            If mPanel.Find("start_stop_timer", pc) Then
                pc.SetValue "0"
                pc.SetEnabled False

            End If
        End If

        ' /* update minute slider */

        If mPanel.Find("timer_mins", pc) Then _
            pc.SetValue CStr(Minute(mTimer))

        ' /* update hours slider */

        If mPanel.Find("timer_hours", pc) Then _
            pc.SetValue CStr(Hour(mTimer))

        ' /* label */

        If mPanel.Find("timer_current", pc) Then _
            pc.SetText Format$(mTimer, "hh:mm:ss")

    End If

End Sub
