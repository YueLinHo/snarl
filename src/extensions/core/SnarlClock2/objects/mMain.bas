Attribute VB_Name = "mMain"
Option Explicit

Dim mCookie() As String
Dim mCookies As Long

Public Type T_CONFIG
    show_time_on_load As Boolean
'    show_as_words As Boolean        ' // hourly reminder in words not numbers
    show_cookie As Boolean          ' // show cookie
    cookie_file As String           ' // cookie file
    show_cookie_on_load As Boolean  ' // 2.0

    ' /* persistent timer settings */

    timer_icon As String
    timer_message As String
    custom_timer_icon As Boolean
    timer_10sec_warning As Boolean
    show_meter As Boolean           ' // 2.0 send % in text during 10 second countdown

    ' /* time format settings */

    format_type As Long             ' // added in Alpha 2 - 0=numeric, 1=text, 2=yyyymmddhhmmss

    use_12_hour_clock As Boolean
    show_am_pm As Boolean
    show_seconds As Boolean
    two_digit_hour As Boolean
    am_val As String
    pm_val As String

    debug_enabled As Boolean

    ' /* not persistent */
    timer_hours As Integer
    timer_mins As Integer
    timer_running As Boolean

End Type

Public gConfig As T_CONFIG

'Public Function g_Notify(ByVal Preview As Boolean) As Long
'Dim szTitle As String
'Dim pc As T_CONFIG
'Dim hr As Long
'
'    If Preview Then
'        szTitle = "Time format preview"
'
'    Else
'        szTitle = "Hourly reminder"
'        LSet pc = gConfig
'
'    End If
'
'    LSet pc = gConfig
'
''    g_Notify = snShowMessageEx(CLASS_HOUR_CHIME, _
'                               szTitle, _
'                               "It's now " & g_GetTimeString(pc, Not Preview), _
'                               10, _
'                               g_MakePath(App.Path) & "icon.png")
'
'End Function

Public Function g_GetTimeString(ByRef Config As T_CONFIG, ByVal HoursOnly As Boolean) As String
Dim txt As String
Dim pt As CTime
Dim i As Integer

    ' /* construct the time format */

    g_Debug "uShowMsg()", LEMON_LEVEL_PROC

    Set pt = New CTime
    pt.SetTo Now()

    If Config.format_type = 1 Then

        ' /* format is either:
        '       hh o'clock
        '       mm minute(s) past hh o'clock
        '       mm minute(s) to hh o'clock
        '       quarter past/half past/quarter to hh o'clock
        ' */

        If Config.use_12_hour_clock Then
            Select Case pt.Minutes()
            Case 0
                ' /* on the hour */
                txt = LCase$(pt.HourText(False)) & " o'clock"

            Case 15
                txt = "quarter past " & LCase$(pt.HourText(False))

            Case 30
                txt = "half past " & LCase$(pt.HourText(False))
            
            Case 45
                txt = "quarter to " & LCase$(pt.NumberToText(pt.Hour12 + 1))

            Case Is > 30
                txt = LCase$(pt.NumberToText(60 - pt.Minutes())) & " minute" & IIf(pt.Minutes = 59, "", "s") & " to " & LCase$(pt.NumberToText(pt.Hour12 + 1)) & " o'clock"

            Case Else
                txt = LCase$(pt.MinuteText()) & " minute" & IIf(pt.Minutes = 1, "", "s") & " past " & LCase$(pt.HourText(False)) & " o'clock"
            
            End Select

        Else
            txt = pt.HourText(True) & " " & pt.MinuteText()

            If Config.show_seconds Then _
                txt = txt & " and " & pt.SecondsText & " second" & IIf(pt.Seconds = 1, "", "s")

            txt = LCase$(txt)

        End If

    ElseIf Config.format_type = 2 Then
        ' /* YYYYMMDDHHMMSS */
        txt = Format$(pt.Year, "0000") & Format$(pt.Month, "00") & Format$(pt.Day, "00") & _
              Format$(pt.Hour, "00") & Format$(pt.Minutes, "00") & Format$(pt.Seconds, "00")

    Else
        ' /* numerically */

        If Config.use_12_hour_clock Then
            i = pt.Hour12

        Else
            i = pt.Hour

        End If

        txt = txt & IIf(Config.two_digit_hour, Format$(i, "00"), CStr(i))

        If Not HoursOnly Then
            txt = txt & ":" & Format$(pt.Minutes(), "00")

            If Config.show_seconds Then _
                txt = txt & ":" & Format$(pt.Seconds(), "00")

        End If

        ' /* am/pm setting only applies for 12 hour clock */

        If (Config.use_12_hour_clock) And (Config.show_am_pm) Then
            If pt.IsAfternoon() Then
                txt = txt & Config.pm_val

            Else
                txt = txt & Config.am_val

            End If
        End If


'    If HoursOnly Then
'        If TwentyFourHourClock Then
'            ' /* 24 hour clock */
'            AsText = HourText() & " hundred hours"
'
'        Else
'            ' /* 12 hour clock */
'
'            If Me.Hour12 = 12 Then
'                AsText = "Noon"
'
'            ElseIf Me.Hour12 = 0 Then
'                AsText = "Midnight"
'
'            Else
'                AsText = HourText() & " o'clock"
'
'            End If
'
'        End If
'
'    Else
'        If TwentyFourHourClock Then
'            AsText = HourText()
'
'        Else
'
'
'        End If
'
'    End If



    End If

'        Else
'
'            ' /* for the hourly reminder, show_seconds means we display as hh:mm:ss instead of h o'clock... */
'
'            If m_Config.show_seconds Then
'
'                ' /* display as h[h]:mm:ss [ampm] */
'
'                i = IIf(m_Config.use_12_hour_clock, .Hour12, .Hour)
'
'                txt = IIf(m_Config.two_digit_hour, Format$(i, "00"), CStr(i))
'                txt = txt & ":00:00"          ' // we _know_ it's 00:00 as it's an hourly reminder...
'
'                If (m_Config.show_am_pm) And (m_Config.use_12_hour_clock) Then
'                    ' /* only show am/pm value if we're using 12 hour format */
'
'                    If .IsAfternoon Then
'                        txt = txt & " " & m_Config.pm_val
'
'                    Else
'                        txt = txt & " " & m_Config.am_val
'
'                    End If
'
'                End If
'            Else
'                txt = CStr(.Hour12) & " o'clock"
'
'            End If
'
'        End If
'
'    End With

Dim a As Long

    If (Config.show_cookie) And (mCookies > 0) Then
        Randomize Timer
        a = Rnd * mCookies
        txt = txt & vbCrLf & vbCrLf & mCookie(a)

    ElseIf Config.show_cookie Then
        g_Debug "uGetTimeString(): show_cookie option set but no cookies available", LEMON_LEVEL_WARNING

    End If

    g_GetTimeString = txt

End Function

Public Sub g_LoadCookies()
Dim n As Integer
Dim sz As String

    On Error Resume Next

    mCookies = 0
    ReDim mCookie(0)

    n = FreeFile()
    Err.Clear
    Open gConfig.cookie_file For Input As #n
    If Err.Number = 0 Then
        Do While Not EOF(n)
            Line Input #n, sz
            If sz <> "" Then
                If Left$(sz, 1) <> "#" Then
                    mCookies = mCookies + 1
                    ReDim Preserve mCookie(mCookies)
                    mCookie(mCookies) = sz

                End If
            End If
        Loop
        Close #n
        g_Debug "g_LoadCookies(): " & CStr(mCookies) & " loaded"

    Else
        g_Debug "g_LoadCookies(): couldn't open '" & gConfig.cookie_file & "'", LEMON_LEVEL_CRITICAL

    End If

End Sub

Public Sub g_WriteConfig()
Dim szPath As String

    If Not snext_getuserpath(szPath) Then _
        Exit Sub

    With New CConfFile
        .Add "cookie_file", gConfig.cookie_file
'        .Add "show_as_words", IIf(gConfig.show_as_words, "1", "0")
        .Add "format_type", CStr(gConfig.format_type)
        .Add "show_cookie", IIf(gConfig.show_cookie, "1", "0")
        .Add "show_time_on_load", IIf(gConfig.show_time_on_load, "1", "0")
        .Add "use_12_hour_clock", IIf(gConfig.use_12_hour_clock, "1", "0")
        .Add "show_am_pm", IIf(gConfig.show_am_pm, "1", "0")
        .Add "show_seconds", IIf(gConfig.show_seconds, "1", "0")
        .Add "two_digit_hour", IIf(gConfig.two_digit_hour, "1", "0")
        .Add "am_val", gConfig.am_val
        .Add "pm_val", gConfig.pm_val
        .Add "debug_mode", IIf(gConfig.debug_enabled, "1", "0")
        .Add "timer_10sec_warning", IIf(gConfig.timer_10sec_warning, "1", "0")
        .Add "timer_message", gConfig.timer_message
        .Add "timer_icon", gConfig.timer_icon
        ' /* 2.0 additions */
        .Add "show_meter", IIf(gConfig.show_meter, "1", "0")
        .Add "show_cookie_on_load", IIf(gConfig.show_cookie_on_load, "1", "0")

        g_Debug "g_WriteConfig(): " & .Save(g_MakePath(szPath) & "etc\.snarlclock2")

    End With

End Sub

Public Function g_GetRandomCookie(ByRef Cookie As String) As Boolean

    If mCookies = 0 Then _
        Exit Function

    Randomize Timer

    Cookie = mCookie(Rnd * mCookies)
    g_GetRandomCookie = (Cookie <> "")

End Function
