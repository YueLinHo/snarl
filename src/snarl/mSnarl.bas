Attribute VB_Name = "mSnarl"
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const WINDOW_CLASS = "w>Snarl"

Public Const WM_SNARL_INIT = WM_USER + 1
Public Const WM_SNARL_QUIT = WM_USER + 2
Public Const WM_SNARL_TRAY_ICON = WM_USER + 3
'Public Const WM_SNARL_NOTIFY_RUNNING = WM_USER + 4

Public Const WM_REMOTENOTIFY = WM_USER + 9              ' // frmAbout: remote notifications
Public Const WM_INSTALL_SNARL = WM_USER + 12            ' // frmAbout: snarl update available

Public Const TIMER_UPDATES = 32

Public Const SNARL_CLASS_WELCOME = "Welcome Message"
Public Const SNARL_CLASS_APP_UNREG = "Application Unregistered"
Public Const SNARL_CLASS_APP_REG = "Application Registered"
Public Const SNARL_CLASS_JSON = "Incoming JSON Notifications"
Public Const SNARL_CLASS_ANON_NET = "Anonymous Remote Notifications"
Public Const SNARL_CLASS_ANON = "Anonymous Local Notifications"

Public Type T_NOTIFICATION_INFO
    pid As Long
    Title As String
    Text As String
    Timeout As E_NOTIFICATION_DURATION
    IconPath As String
    hWndReply As Long
    uReplyMsg As Long
    SndFile As String
    StyleToUse As String
    DefaultAck As String
    Position As E_START_POSITIONS
    Token As Long
    Priority As Long                    ' // V41: <0 = low, 0 = normal, >0 = high

End Type

Public Type T_SNARL_STYLE_ENGINE_INFO
    Name As String
    Version As Long
    Revision As Long
    Date As String
    Path As String                  ' // path to the engine's dll
'    Copyright As String
    Description As String
    Obj As IStyleEngine
    Flags As Long                   ' // bit 31 set means internal

End Type

'Public Enum E_APP_FLAGS
'    E_APP_HAS_ABOUT = SNARL_APP_HAS_ABOUT
'    E_APP_HAS_PREFS = SNARL_APP_HAS_PREFS
'    E_APP_IS_WINDOWLESS = &H8000&
'
'End Enum

Public Type T_SNARL_APP
    Name As String
    hWnd As Long
    uMsg As Long
    pid As Long                 ' // V38 (for V39)
    Icon As String              ' // R1.6 - path to application icon (if empty we use window icon)
    LargeIcon As String         ' // V38 (private for now) - path to large icon
    Token As Long               ' // V41
    Signature As String         ' // V41 - MIME string
    Flags As SNARL41_APP_FLAGS  ' // V41

End Type


Public Type T_SNARL_ICON_THEME
    Name As String
    Path As String
    IconFile As String

End Type

Public gIconTheme() As T_SNARL_ICON_THEME
Public gIconThemes As Long

Private Const SPI_GETFONTSMOOTHING = 74
Private Const SPI_GETFONTSMOOTHINGTYPE = 8202
Private Const FE_FONTSMOOTHINGSTANDARD = 1
Private Const FE_FONTSMOOTHINGCLEARTYPE = 2


Public Const HWND_SNARL = &H534E524C Or &H80000000

Public Const MSG_SHOW_PREFS = WM_USER + 80
Public Const MSG_QUIT = WM_USER + 81

'Public bm_Back As mfxBitmap
'Public bm_Shadow As mfxBitmap
Public bm_Close As MImage

Public Enum E_START_POSITIONS
    ' /* IMPORTANT!! These have now changed under V41 */
    E_START_DEFAULT_POS = 0
    E_START_TOP_LEFT
    E_START_TOP_RIGHT
    E_START_BOTTOM_LEFT
    E_START_BOTTOM_RIGHT

End Enum

    ' /* these only apply if the new E_CLASS_CUSTOM_DURATION is not set */

Public Enum E_NOTIFICATION_DURATION
    ' /* IMPORTANT!! These have now changed under V41 */
    E_DURATION_DEFAULT = 0
    E_DURATION_APP_DECIDES
    E_DURATION_CUSTOM           ' // "custom_timeout" contains value in seconds

End Enum

'Public Const GOFFSET = 4

Public g_IsRunning  As Boolean
Public g_IsQuitting As Boolean

Public g_ExtnRoster As TExtensionRoster
Public g_StyleRoster As TStyleRoster
Public g_AppRoster As TApplicationRoster
Public g_NotificationRoster As TNotificationRoster

Public Enum E_FONTSMOOTHING
    E_MELONTYPE
    E_NONE
    E_ANTIALIAS
    E_CLEARTYPE
    E_WINDOWS_DEFAULT

End Enum


Public Type G_NOTIFICATION_CONTENT
    Title As String
    Text As String
    Icon As mfxBitmap
    Timeout As Long
    Ack As String

    Timestamp As Date

End Type


Public Type T_CONFIG
    ' /* non-configurable elements */
'    STEP_SIZE As Double

    ' /* user-configurable elements */
'    default_position As E_START_POSITIONS
'    default_opacity As Long
    show_msg_on_start As Boolean
    run_on_logon As Boolean
'    default_dye_colour As Long      ' // rgba

    ' /* R2.0 (V38.13) */
'    default_style As String

    ' /* R2.0 (V38.32) */
    sticky_snarls As Boolean
    log_only As Boolean
    default_duration As Long

    ' /* R2.04 (V38.82) */
    font_smoothing As E_FONTSMOOTHING
    melontype_contrast As Long

    ' /* R2.1 (V39) */
'    listen_for_json As Boolean
'    listen_for_snarl As Boolean
    suppress_delay As Long          ' // in ms
    hotkey_prefs As Long            ' // MAKELONG(mods,key)
'    hotkey_test_message As Long     ' // MAKELONG(mods,key)
    notify_on_first_register As Boolean
    global_opacity As Long
    last_sound_folder As String
    show_tray_icon As Boolean
    ignore_new_classes As Boolean   ' // new alert classes are always disabled by default

    ' /* R2.2 */
    use_hotkey As Boolean
    do_not_disturb_ As Boolean      ' // not persitent: user-controlled DND setting
    dnd_count As Long               ' // set using WM_MANAGE_SNARL

    idle_timeout As Long     ' // number of seconds of idle before enabling sticky mode
    margin_spacing As Long          ' // px between notification and screen edge
'    icon_theme As String            ' // name of icon theme to use (if blank, 'default_theme' is used)
    use_dropshadow As Boolean
    dropshadow_strength As Long
    dropshadow_size As Long
    auto_update As Boolean
    last_update_check As Date

    enable_sounds As Boolean
    use_style_sounds As Boolean
    prefer_style_sounds As Boolean
    default_notification_sound As String
    default_priority_sound As String

    AgreeBetaUsage As Boolean

End Type

Public g_Prefs As T_CONFIG
'Public g_TempPrefs As T_CONFIG
'Dim m_GlobalId As Long

Dim mSettings As ConfigFile         ' // V40.25 - new way of managing persistent settings
Dim mConfig As ConfigSection        ' // V40.25 - the actual config section
Dim mDefaults As BPackedData        ' // V40.25 - new way of managing persistent settings
Dim mConfigLocked As Boolean
Dim mWriteConfigOnUnlock As Boolean

'Public g_Settings As ConfigFile
Dim m_Alerts As ConfigSection

Public g_IgnoreLock As Long         ' // if >0 don't alert when app registers - overrides class setting


Public gSelectedClass As TAlert

Public gDebugMode As Boolean

'Public gAwayModeEnabled As Boolean
Public gIsIdle As Boolean

Public Type G_REMOTE_COMPUTER
    IsHostName As Boolean
    HostNameOrIp As String

End Type

Public gRemoteComputers As ConfigSection

Public gForwarder() As TForwarder
Public gForwarders As Long
Public gForwarderUID As Long

Public ghWndMain As Long
Public gUpdateFilename As String            ' // name of the update file to download

Dim myUpdateCheck As TAutoUpdate

Public gLastErr As SNARL_STATUS_41      ' // V41 api

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub CoFreeUnusedLibrariesEx Lib "ole32" (ByVal dwUnloadDelay As Long, ByVal dwReserved As Long)
'Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Private Declare Sub CoFreeLibrary Lib "ole32.dll" (ByVal hInst As Long)

'Public gClassToken As Long

Public Sub Main()
Dim sz As String
Dim l As Long

    ' /* first off, get the command line */

    sz = Command$

    ' /* get comctl.dll loaded up for our XP manifest... */

    g_InitComCtl

    ' /* if '-quit' was specified in the command line look for an already
    '    running instance of snarl and tell it to quit */

    If InStr(sz, "-quit") Then
        l = FindWindow(WINDOW_CLASS, "Snarl")
        If l <> 0 Then _
            SendMessage l, MSG_QUIT, 0, ByVal 0&    ' // tell the running instance to quit...
                    ' // better to use WM_CLOSE??

        Exit Sub

    End If

    App.TaskVisible = False

    ' /* check we're not already running */

    l = FindWindow(WINDOW_CLASS, "Snarl")
    If l <> 0 Then
        PostMessage l, MSG_SHOW_PREFS, 0, ByVal 0&    ' // tell the running instance to show its gui...
        Exit Sub

    End If

    ' /* V38.133 - enable debug mode if switch present */
    ' /* V40.18 - or if either CTRL key is held down */

    gDebugMode = (InStr(sz, "-debug") > 0) Or (g_IsPressed(VK_LCONTROL)) Or (g_IsPressed(VK_RCONTROL))

    If gDebugMode Then
        ' /* start logging */
        l3OpenLog "%APPDATA%\snarl.log"
        g_Debug "** Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ") **"
        g_Debug "** " & App.LegalCopyright
        g_Debug ""

    End If

    ' /* need to check exec.library first */

    If Not melonGetExec() Then
        MsgBox "Snarl requires exec.library V46 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo noexec

    End If

    ' /* NOTE! We no longer need V46 graphics but older extensions will rely on it */

    If open_library("graphics.library", 46) <> M_OK Then
        MsgBox "Snarl requires graphics.library V46 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo nographics

    End If

    If open_library("libnitro1_2", 41) <> M_OK Then
        MsgBox "Snarl requires Nitro R1.2 V41 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo nonitro

    End If

    If open_library("openmenulite.library", 46) <> M_OK Then
        MsgBox "Snarl requires openmenulite.library V46 or greater", vbCritical, "Snarl Initialisation Failed"
        GoTo noopenmenu

    End If

    ' /* check storage kit */

    If Not melonCheckKit("storage") Then
        MsgBox "Snarl needs the Storage Kit!", vbCritical, "Snarl Initialisation Failed"
        GoTo nostoragekit

    End If

Dim pName As String

    ' /* resources */

    If Not melonCheckLibOrResource("icon_resource", 0, pName) Then
        MsgBox "icon.resource is damaged or not installed", vbCritical, "Snarl Early Startup Error"
        GoTo nostoragekit

    Else
        g_Debug "Main(): got " & pName

    End If

    If Not melonCheckLibOrResource("web_resource", 0, pName) Then
        MsgBox "web.resource is damaged or not installed", vbCritical, "Snarl Early Startup Error"
        GoTo nostoragekit

    Else
        g_Debug "Main(): got " & pName

    End If

    If Not melonCheckLibOrResource("misc_resource", 0, pName) Then
        MsgBox "misc.resource is damaged or not installed", vbCritical, "Snarl Early Startup Error"
        GoTo nostoragekit

    Else
        g_Debug "Main(): got " & pName

    End If



    ' /* -------------- end of early startup ------------- */

    ' /* the first thing we should do now is create the message handling window */

    If Not EZRegisterClass(WINDOW_CLASS) Then
        g_Debug "main(): couldn't register window class", LEMON_LEVEL_CRITICAL
        GoTo nostoragekit

    End If

    ghWndMain = EZAddWindow(WINDOW_CLASS, New TMainWindow, "Snarl", 0, 0)
    If ghWndMain = 0 Then
        g_Debug "main(): couldn't create window", LEMON_LEVEL_CRITICAL
        EZUnregisterClass WINDOW_CLASS
        GoTo nostoragekit

    End If

    ' /* V41: set our system version and revision as properties */

    SetProp ghWndMain, "_version", App.Major
    SetProp ghWndMain, "_revision", App.Revision

    ' /* now we have the handling window, we can set our console */

'    lemonSetConsole ghWndMain

    load_image g_MakePath(App.Path) & "etc\icons\close.png", bm_Close

    If Not g_IsValidImage(bm_Close) Then
        With New mfxView
            .SizeTo 24, 24
            .EnableSmoothing False
            .SetHighColour rgba(255, 0, 0)
            .FillRect .Bounds
            .SetHighColour rgba(0, 0, 0)
            .StrokeRect .Bounds
            Set bm_Close = .ConvertToBitmap()

        End With
    End If

    gForwarderUID = &HE0

    ' /* get settings */

    If Not g_ConfigInit() Then

        MsgBox "As this is the first time Snarl has been run, I need to test the drawing" & vbCrLf & _
               "speed of your PC.  This test will only take a few seconds, if that.", _
               vbOKOnly Or vbInformation, "Snarl Graphics Test"

        l = GetTickCount()
        g_Debug "main(): GFXCHK: starting graphics test (tick_count=" & CStr(l) & ")"

Dim pInfo As T_NOTIFICATION_INFO

        With pInfo
            .Title = "Snarl Graphics Test"
            .Text = "Test Message"
            .Timeout = 1
            .StyleToUse = ""            ' // the scheme

        End With

        With New CSnarlWindow
            .Create New TAlert, pInfo, New TInternalStyle, 0, ""
            '"Snarl Graphics Test", "Test message", 1, "", 0, 0, 0, "", New TInternalStyle, "", 0
            .Quit

        End With

Dim dStep As Double

        l = GetTickCount() - l
        g_Debug "main(): GFXCHK: completed graphics test (tick_count=" & CStr(GetTickCount()) & ")"

        dStep = l / 422#

        If dStep < 1# Then _
            dStep = 1#

        g_Debug "main(): GFXCHK: delta=" & l & " actual=" & CStr(l / 422#) & " step=" & dStep

        MsgBox "Graphics test complete.  Snarl has calculated a factor of " & Format$(dStep, "0.00") & vbCrLf & _
               "to use when displaying messages.", vbOKOnly Or vbInformation, "Test Complete"

        g_ConfigSet "step_size", CStr(dStep)

    End If

    ' /* main start */

    SendMessage ghWndMain, WM_SNARL_INIT, 0, ByVal 0&

    Load frmAbout           ' // keeps us open...

    ' /* start notification roster */

    g_Debug "Main(): Starting notifications roster..."
    Set g_NotificationRoster = New TNotificationRoster
    melonLibInit g_NotificationRoster
    melonLibOpen g_NotificationRoster

    ' /* start app roster */

    g_Debug "Main(): Starting app roster..."
    Set g_AppRoster = New TApplicationRoster
    melonLibInit g_AppRoster
    melonLibOpen g_AppRoster

    g_Debug "Main(): Setting auto-run state..."
    g_SetAutoRun g_Prefs.run_on_logon

    ' /* get style packs */

    g_Debug "Main(): Starting style roster..."
    Set g_StyleRoster = New TStyleRoster
    melonLibInit g_StyleRoster
    melonLibOpen g_StyleRoster

    ' /* get icon themes */

    g_GetIconThemes

    ' /* tell everyone we're open for business */

    g_Debug "Notifying ready to run..."
    g_SetRunning True, True


    ' /* get extensions */

    g_Debug "Main(): Starting extensions roster..."
    Set g_ExtnRoster = New TExtensionRoster
    melonLibInit g_ExtnRoster
    melonLibOpen g_ExtnRoster

#If BETA_REL = 1 Then

Dim pBetaPanel As TBetaPanel

    If (Not g_Prefs.AgreeBetaUsage) Then
        Set pBetaPanel = New TBetaPanel
        pBetaPanel.Go

    End If

#End If

    If g_Prefs.auto_update Then
        g_Debug "Main(): Doing auto-update check..."
        Set myUpdateCheck = New TAutoUpdate
        If myUpdateCheck.Check(False) Then _
            g_Debug "Main(): auto-update check initiated"

    Else
        g_Debug "Main(): Auto-update is disabled"

    End If

    g_Debug "Main(): startup complete"
    With New BMsgLooper
        .Run

    End With

    Set myUpdateCheck = Nothing         ' // this will abort the request if it's still running...

Dim uSnarlGlobal As Long
Dim t As Long

    g_Debug "main(): stopping..."

#If BETA_REL = 1 Then
    Set pBetaPanel = Nothing
#End If

    t = GetTickCount()
    g_IsQuitting = True

    ' /* broadcast SNARL_QUIT */

    uSnarlGlobal = snGetGlobalMsg()
    If uSnarlGlobal > 0 Then
        g_Debug "main(): broadcasting SNARL_QUIT..."
'        SendMessageTimeout HWND_BROADCAST, um, SNARL_QUIT, ByVal 0&, SMTO_NORMAL, 10, i
        PostMessage HWND_BROADCAST, uSnarlGlobal, SNARL_QUIT, ByVal 0&

    Else
        g_Debug "main(): snGetGlobalMsg() returned zero", LEMON_LEVEL_WARNING

    End If

    SendMessage ghWndMain, WM_SNARL_QUIT, 0, ByVal 0&
    Unload frmAbout

    ' /* stop various rosters - order *is* important (should be reverse of startup) */

    g_Debug "main(): stopping extension roster..."
    melonLibClose g_ExtnRoster
    melonLibUninit g_ExtnRoster

    g_Debug "main(): stopping style roster..."
    melonLibClose g_StyleRoster
    melonLibUninit g_StyleRoster

    g_Debug "main(): stopping application roster..."
    melonLibClose g_AppRoster
    melonLibUninit g_AppRoster

    g_Debug "main(): stopping notification roster..."
    melonLibClose g_NotificationRoster
    melonLibUninit g_NotificationRoster

    EZRemoveWindow ghWndMain
    EZUnregisterClass WINDOW_CLASS

    ' /* done */

    t = GetTickCount() - t
    g_Debug "main(): took " & t & " ms to complete closedown"

nostoragekit:
    close_library "openmenulite.library"

noopenmenu:
    close_library "Nitro R1.2"

nonitro:
    close_library "graphics.library"

nographics:

noexec:

'    lemonUnregister

End Sub


'Public Sub g_Quit()
'Dim um As Long
'Dim i As Long
'
'    ' /* broadcast SNARL_QUIT */
'
'    um = snGetGlobalMsg()
'    If um > 0 Then
'        sosOutput "g_Quit(): broadcasting SNARL_QUIT..."
''        SendMessageTimeout HWND_BROADCAST, um, SNARL_QUIT, ByVal 0&, SMTO_NORMAL, 10, i
'        PostMessage HWND_BROADCAST, um, SNARL_QUIT, ByVal 0&
'
'    Else
'        sosOutput "g_Quit(): snGetGlobalMsg() returned zero", LEMON_LEVEL_WARNING
'
'    End If
'
'    ' /* stop various rosters - order *is* important
'    '    (and should be reverse of startup) */
'
'    g_Debug "g_Quit(): stopping extension roster..."
'    melonLibClose g_ExtnRoster
'    melonLibUninit g_ExtnRoster
'
'    g_Debug "g_Quit(): stopping style roster..."
'    melonLibClose g_StyleRoster
'    melonLibUninit g_StyleRoster
'
'    g_Debug "g_Quit(): stopping application roster..."
'    melonLibClose g_AppRoster
'    melonLibUninit g_AppRoster
'
'    g_Debug "g_Quit(): stopping notification roster..."
'    melonLibClose g_NotificationRoster
'    melonLibUninit g_NotificationRoster
'
'
'
''    If mNotifications Then
''        sosOutput "g_Quit(): zapping " & CStr(mNotifications) & " snarl(s)"
''        For i = mNotifications To 1 Step -1
''            If Not (mNotification(i).Window Is Nothing) Then _
''                mNotification(i).Window.Zap
''
''        Next i
''    End If
''
''    ReDim mNotification(0)
''    mNotifications = 0
'
''    If g_Applets Then
''        For i = g_Applets To 1 Step -1
''            If (IsWindow(g_Applet(i).hWnd) <> 0) And (g_Applet(i).uMsg <> 0) Then _
''                SendMessage g_Applet(i).hWnd, g_Applet(i).uMsg, SNARL_QUIT, ByVal 0&
''
''        Next i
''    End If
'
'    close_library "Nitro R1.2"
'    close_library "graphics.library"
'    close_library "openmenulite.library"
'
'    sosOutput "g_Quit(): done", LEMON_LEVEL_PROC
'
'End Sub



'Public Function globalSetTimeout(ByVal msgId As Long, ByVal Timeout As Long) As M_RESULT
'Dim i As Long
'
'    i = uFindMsg(msgId)
'    If i Then
'        g_Debug "globalSetTimeout(): id " & msgId & " found at index " & i, LEMON_LEVEL_INFO
'        mNotification(i).Window.SetTimeout Timeout
'        globalSetTimeout = M_OK
'
'    Else
'        g_Debug "globalSetTimeout(): id " & msgId & " not found", LEMON_LEVEL_WARNING
'        globalSetTimeout = M_NOT_FOUND
'
'    End If
'
'End Function

Public Function g_ConfigInit() As Boolean

    On Error Resume Next

    With g_Prefs
'        .default_position = E_START_BOTTOM_RIGHT
        .show_msg_on_start = True
        .run_on_logon = True
'        .step_size = 1#
'        .default_dye_colour = 0     ' // none

        .default_duration = 10
        .log_only = False
        .sticky_snarls = False
'        .default_style = "iphoney/standard"

        .font_smoothing = E_MELONTYPE
        .melontype_contrast = 10
        .suppress_delay = 2000
        .hotkey_prefs = vbKeyF10
'        .hotkey_test_message = MAKELONG(MOD_CONTROL Or MOD_SHIFT, vbKeyF10)
        .notify_on_first_register = False
        .global_opacity = 100

'        If g_GetSystemFolder(CSIDL_PERSONAL, sz) Then _
            .last_sound_folder = sz

        .show_tray_icon = True
        .use_hotkey = True
        .do_not_disturb_ = False
        .idle_timeout = 300          ' // == 5 minutes
        .margin_spacing = 0
        .use_dropshadow = True
        .dropshadow_strength = 88           ' // as a % - translated into a 0-255 opacity value
        .dropshadow_size = 10
'        .icon_theme = ""

        .auto_update = True

        .enable_sounds = True
        .use_style_sounds = True
        .prefer_style_sounds = False
        .default_notification_sound = ""
        .default_priority_sound = ""

    End With


    ' /* defaults */

    Set mDefaults = New BPackedData
    With mDefaults
        .Add "default_position", CStr(E_START_BOTTOM_RIGHT)
        .Add "show_msg_on_start", "1"
        .Add "run_on_logon", "1"

        ' /* R2.0 (V38.13) */
        .Add "default_style", "iphoney/standard"    ' // as "<style>[/<scheme>]

        ' /* R2.0 (V38.32) */
        .Add "sticky_snarls", "0"
        .Add "log_only", "0"
        .Add "default_duration", "10"

        ' /* R2.04 (V38.82) - no longer used */
        .Add "font_smoothing", CStr(E_MELONTYPE)
        .Add "melontype_contrast", "10"

        ' /* R2.1 (V39) */
        .Add "listen_for_json", "0"
        .Add "listen_for_snarl", "0"
        .Add "duplicates_quantum", "2000"
        .Add "hotkey_prefs", CStr(vbKeyF10)
        .Add "notify_on_first_register", "0"
        .Add "global_opacity", "100"
        .Add "last_sound_folder", g_GetSystemFolderStr(CSIDL_PERSONAL)
        .Add "show_tray_icon", "1"
        .Add "ignore_new_classes", "0"      ' // new alert classes are always disabled by default

        ' /* R2.2 */
        .Add "use_hotkey", "1"
        .Add "do_not_disturb", "0"
        .Add "idle_timeout", "300"          ' // i.e. 5 minutes
        .Add "margin_spacing", "0"
        .Add "use_dropshadow", "1"
        .Add "dropshadow_strength", "88"    ' // is a %
        .Add "dropshadow_size", "10"
        .Add "icon_theme", ""
        
        .Add "auto_update", "1"
        .Add "enable_sounds", "1"
        .Add "use_style_sounds", "1"
        .Add "prefer_style_sounds", "0"
        .Add "default_normal_sound", ""
        .Add "default_priority_sound", ""
        .Add "use_style_icons", "1"

    End With

    ' /* attempt to load the config file */

Dim sz As String

    If Not g_GetUserFolderPath(sz) Then
        g_Debug "g_ConfigInit(): %APP_DATA% path not found", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    sz = g_MakePath(sz) & "etc"

Dim i As Long

    Set mSettings = New ConfigFile
    With mSettings
        .File = sz & "\config41.snarl"
        .Load

        i = .FindSection("general")
        If i = 0 Then
            Set mConfig = .AddSectionObj("general")
            .Save

        Else
            Set mConfig = .SectionAt(i)

        End If

        i = .FindSection("remote_computers")
        If i = 0 Then
            Set gRemoteComputers = .AddSectionObj("remote_computers")
            .Save

        Else
            Set gRemoteComputers = .SectionAt(i)

        End If

    End With

    g_ConfigInit = (Val(g_ConfigGet("step_size")) > 0)

End Function


'Public Function g_ReadConfig() As Boolean
'Dim sz As String
'Dim dw As Long
'Dim i As Long
'Dim f As Boolean
'
'    On Error Resume Next
'
'    ' /* defaults */
'
'    With g_Prefs
'        .default_position = E_START_BOTTOM_RIGHT
'        .show_msg_on_start = True
'        .run_on_logon = True
''        .step_size = 1#
''        .default_dye_colour = 0     ' // none
'
'        .default_duration = 10
'        .log_only = False
'        .sticky_snarls = False
'        .default_style = "iphoney/standard"
'
'        .font_smoothing = E_MELONTYPE
'        .melontype_contrast = 10
'        .suppress_delay = 2000
'        .hotkey_prefs = vbKeyF10
''        .hotkey_test_message = MAKELONG(MOD_CONTROL Or MOD_SHIFT, vbKeyF10)
'        .notify_on_first_register = False
'        .global_opacity = 100
'
'        If g_GetSystemFolder(CSIDL_PERSONAL, sz) Then _
'            .last_sound_folder = sz
'
'        .show_tray_icon = True
'        .use_hotkey = True
'        .do_not_disturb_ = False
'        .idle_timeout = 300          ' // == 5 minutes
'        .margin_spacing = 0
'        .use_dropshadow = True
'        .dropshadow_strength = 88           ' // as a % - translated into a 0-255 opacity value
'        .dropshadow_size = 10
'        .icon_theme = ""
'
'        .auto_update = True
'
'        .enable_sounds = True
'        .use_style_sounds = True
'        .prefer_style_sounds = False
'        .default_notification_sound = ""
'        .default_priority_sound = ""
'
'    End With
'
'    ' /* do we have a V37 config? */
'
'    g_Debug "g_ReadConfig(): loading config..."
'
'    Set g_Settings = New ConfigFile
'    With g_Settings
'        If g_GetSystemFolder(CSIDL_APPDATA, sz) Then
'            ' /* V38 - look for settings in %appdata% first */
'            .File = g_MakePath(sz) & "full phat\snarl\etc\v41.snarl"
'            f = .Load()
'
'            If Not f Then _
'                g_Debug "g_ReadConfig(): not found in '" & .File & "'"
'
'        End If
'
'        If Not f Then
'            ' /* if this fails, look in %pwd% */
'            .File = g_MakePath(App.Path) & "etc\v41.snarl"
'            f = .Load()
'            If Not f Then _
'                g_Debug "g_ReadConfig(): not found in '" & .File & "'"
'
'        End If
'
'        If f Then
'            i = .FindSection("snarl")
'            If i = 0 Then
'                g_Debug "g_ReadConfig(): pre-V37 config detected", LEMON_LEVEL_WARNING
'
'                ' /* load existing settings */
'
'                With New CConfFile
'                    If .SetTo(g_MakePath(App.Path) & "etc\.snarl") Then
'
'                        val(g_configget("step_size")) = .GetValueWithDefault("gfxdelta", "0")
'
''                        If .FindEntry("snarlopacity", sz) Then
''                            dw = Val(sz)
''                            If (dw < 1) Or (dw > 255) Then
''                                g_Debug "g_ReadConfig(): 'snarlopacity' was set to invalid value '" & sz & "'", LEMON_LEVEL_WARNING
''                                dw = 200
''
''                            End If
''                            g_Prefs.default_opacity = dw
''
''                        End If
'
'                        If .FindEntry("startinfo", sz) Then _
'                            g_Prefs.show_msg_on_start = .StrToBool(sz)
'
'                        If .FindEntry("startonlogon", sz) Then _
'                            g_Prefs.run_on_logon = .StrToBool(sz)
'
'                        ' /* XxXxXxXxXxX DON'T ADD HERE!  THIS IS FOR PRE-V37 CONFIGS XxXxXxXxXxX */
'
'                    Else
'                        g_Debug "g_ReadConfig(): failed to load pre-V37 config", LEMON_LEVEL_CRITICAL
'
'                    End If
'
'                End With
'            Else
'                ' /* read settings from V37 config */
'
'                g_Debug "g_ReadConfig(): using V37 config", LEMON_LEVEL_INFO
'                val(g_configget("step_size")) = Val(.SectionAt(i).GetValueWithDefault("gfxdelta", "0"))
'                g_Prefs.default_position = Val(.SectionAt(i).GetValueWithDefault("default_position", CStr(E_START_POSITIONS.E_START_BOTTOM_RIGHT)))
''                g_Prefs.default_opacity = Val(.SectionAt(i).GetValueWithDefault("snarlopacity", 255))
'                g_Prefs.show_msg_on_start = .StrToBool(.SectionAt(i).GetValueWithDefault("startinfo", "1"))
'                g_Prefs.run_on_logon = .StrToBool(.SectionAt(i).GetValueWithDefault("startonlogon", "1"))
'                g_Prefs.default_style = .SectionAt(i).GetValueWithDefault("default_style", "iphoney/standard")
'                g_Prefs.log_only = .StrToBool(.SectionAt(i).GetValueWithDefault("log_only", "0"))
'                g_Prefs.sticky_snarls = .StrToBool(.SectionAt(i).GetValueWithDefault("sticky_snarls", "0"))
'                g_Prefs.default_duration = Val(.SectionAt(i).GetValueWithDefault("default_duration", "10"))
'
'                With g_Prefs
'                    If .default_duration < 0 Then
'                        .default_duration = 0
'
'                    ElseIf .default_duration > 60 Then
'                        .default_duration = 60
'
'                    End If
'
'                    If (.default_position < 0) Or (.default_position > 3) Then _
'                        .default_position = E_START_BOTTOM_RIGHT
'
'                End With
'
'                ' /* R2.04 */
'
'                g_Prefs.font_smoothing = Val(.SectionAt(i).GetValueWithDefault("font_smoothing", "0"))
'                g_Prefs.melontype_contrast = Val(.SectionAt(i).GetValueWithDefault("melontype_contrast", "10"))
'
'                With g_Prefs
'                    If (.font_smoothing < E_MELONTYPE) Or (.font_smoothing > E_WINDOWS_DEFAULT) Then _
'                        .font_smoothing = E_MELONTYPE
'
'                    If (.melontype_contrast < 0) Or (.melontype_contrast > 100) Then _
'                        .melontype_contrast = 10
'
'                    ' /* R2.06 */
'                    If .default_style = "" Then _
'                        .default_style = "glass/smoke"
'
'                End With
'
'                ' /* R2.1 */
'
'                g_Prefs.listen_for_json = .StrToBool(.SectionAt(i).GetValueWithDefault("listen_for_json", "0"))
'                g_Prefs.listen_for_snarl = .StrToBool(.SectionAt(i).GetValueWithDefault("listen_for_snarl", "0"))
'                g_Prefs.suppress_delay = Val(.SectionAt(i).GetValueWithDefault("suppress_delay", "2000"))
'
'                With g_Prefs
'                    If .suppress_delay < 500 Then _
'                        .suppress_delay = 500
'
'                End With
'
'                If .SectionAt(i).Find("hotkey_prefs", sz) Then _
'                    g_Prefs.hotkey_prefs = Val(sz)
'
'                g_Prefs.notify_on_first_register = .StrToBool(.SectionAt(i).GetValueWithDefault("notify_on_first_register", "0"))
'
'                If .SectionAt(i).Find("global_opacity", sz) Then _
'                    g_Prefs.global_opacity = Val(sz)
'
'                ' /* 2.1b4 */
'
'                With g_Prefs
'                    If .global_opacity > 100 Then
'                        .global_opacity = 100
'
'                    ElseIf .global_opacity < 1 Then
'                        .global_opacity = 1
'
'                    End If
'
'                End With
'
'                If .SectionAt(i).Find("last_sound_folder", sz) Then _
'                    g_Prefs.last_sound_folder = sz
'
'                g_Prefs.show_tray_icon = .StrToBool(.SectionAt(i).GetValueWithDefault("show_tray_icon", "1"))
'                g_Prefs.ignore_new_classes = .StrToBool(.SectionAt(i).GetValueWithDefault("ignore_new_classes", "0"))
'
'                ' /* 2.2 */
'
'                g_Prefs.use_hotkey = .StrToBool(.SectionAt(i).GetValueWithDefault("use_hotkey", "1"))
'                g_Prefs.do_not_disturb_ = .StrToBool(.SectionAt(i).GetValueWithDefault("do_not_disturb", "0"))
'
'                If .SectionAt(i).Find("idle_timeout", sz) Then _
'                    g_Prefs.idle_timeout = Val(sz)
'
'                If .SectionAt(i).Find("margin_spacing", sz) Then _
'                    g_Prefs.margin_spacing = Val(sz)
'
'                g_Prefs.use_dropshadow = .StrToBool(.SectionAt(i).GetValueWithDefault("use_dropshadow", "1"))
'
'                If .SectionAt(i).Find("dropshadow_strength", sz) Then _
'                    g_Prefs.dropshadow_strength = Val(sz)
'
'                With g_Prefs
'                    If .dropshadow_strength > 100 Then
'                        .dropshadow_strength = 100
'
'                    ElseIf .dropshadow_strength < 0 Then
'                        .dropshadow_strength = 0
'
'                    End If
'
'                End With
'
'                If .SectionAt(i).Find("dropshadow_size", sz) Then _
'                    g_Prefs.dropshadow_size = Val(sz)
'
'                If .SectionAt(i).Find("icon_theme", sz) Then _
'                    g_Prefs.icon_theme = sz
'
'                ' /* note the name mismatch - the config file entry changes depending on the
'                '    beta this release actually is */
'
'                If .SectionAt(i).Find("AgreeR22BetaUsage", sz) Then _
'                    g_Prefs.AgreeBetaUsage = (sz = "1")
'
'                If .SectionAt(i).Find("auto_update", sz) Then _
'                    g_Prefs.auto_update = (sz = "1")
'
'                If .SectionAt(i).Find("enable_sounds", sz) Then _
'                    g_Prefs.enable_sounds = (sz = "1")
'
'                If .SectionAt(i).Find("use_style_sounds", sz) Then _
'                    g_Prefs.use_style_sounds = (sz = "1")
'
'                If .SectionAt(i).Find("prefer_style_sounds", sz) Then _
'                    g_Prefs.prefer_style_sounds = (sz = "1")
'
'                If .SectionAt(i).Find("default_notification_sound", sz) Then _
'                    g_Prefs.default_notification_sound = sz
'
'                If .SectionAt(i).Find("default_priority_sound", sz) Then _
'                    g_Prefs.default_priority_sound = sz
'
'
'                g_ReadConfig = True
'
'            End If
'
'        Else
'            g_Debug "g_ReadConfig(): config not found", LEMON_LEVEL_WARNING
'
'        End If
'
'        If i = 0 Then _
'            .AddSection "snarl"
'
'
'        i = .FindSection("remote_computers")
'
'        If i Then
'            Set gRemoteComputers = .SectionAt(i)
'
'        Else
'            i = .AddSection("remote_computers")
'            Set gRemoteComputers = .SectionAt(i)
'
'        End If
'
'    End With
'
'    ' /* validate settings */
'
''    dw = Val(g_Prefs.default_opacity)
''    If (dw < 1) Or (dw > 255) Then
''        g_Debug "g_ReadConfig(): 'snarlopacity' was set to invalid value '" & sz & "'", LEMON_LEVEL_WARNING
''        dw = 200
''
''    End If
''    g_Prefs.default_opacity = dw
'
''    If val(g_configget("step_size")) < 1# Then _
''        val(g_configget("step_size")) = 1
'
'    g_ReadConfig = (val(g_configget("step_size")) > 0)
'
'    ' /* write updated file out */
'
'    g_WriteConfig
'
'End Function

'Public Sub g_WriteConfig()
'
'    Debug.Print "g_WriteConfig: " & gNoWriteConfig
'
'    If gNoWriteConfig Then _
'        Exit Sub
'
'    ' /* V38 - always save to %appdata% */
'
'Dim sz As String
'
'    If Not g_GetSystemFolder(CSIDL_APPDATA, sz) Then
'        g_Debug "g_WriteConfig(): couldn't retrieve %appdata% location", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If
'
'    sz = g_MakePath(sz) & "full phat\snarl\etc"
'    create_directory_path sz
'
'Dim i As Long
'
'    With g_Settings
'        ' /* V38 - force the filename */
'        .File = g_MakePath(sz) & ".snarl"
'
'        i = .FindSection("snarl")
'        If i Then
'            With .SectionAt(i)
'                .Update "gfxdelta", CStr(g_Prefs.STEP_SIZE)
'                .Update "default_position", CStr(g_Prefs.default_position)
'                .Update "startinfo", g_Settings.BoolToStr(g_Prefs.show_msg_on_start)
'                .Update "startonlogon", g_Settings.BoolToStr(g_Prefs.run_on_logon)
'                .Update "default_style", g_Prefs.default_style
'
'                .Update "log_only", g_Settings.BoolToStr(g_Prefs.log_only)
'                .Update "sticky_snarls", g_Settings.BoolToStr(g_Prefs.sticky_snarls)
'                .Update "default_duration", CStr(g_Prefs.default_duration)
'
'                ' /* R2.04 */
'                .Update "font_smoothing", CStr(g_Prefs.font_smoothing)
'                .Update "melontype_contrast", CStr(g_Prefs.melontype_contrast)
'
'                ' /* R2.1 */
'                .Update "listen_for_json", g_Settings.BoolToStr(g_Prefs.listen_for_json)
'                .Update "listen_for_snarl", g_Settings.BoolToStr(g_Prefs.listen_for_snarl)
'                .Update "suppress_delay", CStr(g_Prefs.suppress_delay)
'                .Update "hotkey_prefs", CStr(g_Prefs.hotkey_prefs)
'                .Update "notify_on_first_register", g_Settings.BoolToStr(g_Prefs.notify_on_first_register)
'                ' /* R2.1d4 */
'                .Update "global_opacity", CStr(g_Prefs.global_opacity)
'                .Update "last_sound_folder", g_Prefs.last_sound_folder
'                .Update "show_tray_icon", g_Settings.BoolToStr(g_Prefs.show_tray_icon)
'                .Update "ignore_new_classes", g_Settings.BoolToStr(g_Prefs.ignore_new_classes)
'                ' /* R2.2 */
'                .Update "use_hotkey", g_Settings.BoolToStr(g_Prefs.use_hotkey)
'                .Update "do_not_disturb", g_Settings.BoolToStr(g_Prefs.do_not_disturb_)
'                .Update "idle_timeout", CStr(g_Prefs.idle_timeout)
'                .Update "margin_spacing", CStr(g_Prefs.margin_spacing)
'                .Update "use_dropshadow", g_Settings.BoolToStr(g_Prefs.use_dropshadow)
'                .Update "dropshadow_strength", CStr(g_Prefs.dropshadow_strength)
'                .Update "dropshadow_size", CStr(g_Prefs.dropshadow_size)
'                .Update "icon_theme", g_Prefs.icon_theme
'                .Update "auto_update", g_Settings.BoolToStr(g_Prefs.auto_update)
'
'
'                .Update "AgreeR22BetaUsage", g_Settings.BoolToStr(g_Prefs.AgreeBetaUsage)
'
'            End With
'
'            If Not (g_AppRoster Is Nothing) Then _
'                g_AppRoster.WriteConfig
'
'
''            With gRemoteConfig
''                .MakeEmpty
''                If gRemoteCount Then
''                    For i = 1 To gRemoteCount
''                        .Add "ip", gRemoteComputer(i).HostNameOrIp
''
''                    Next i
''                End If
''
''            End With
'
'        Else
'            g_Debug "g_WriteConfig(): [snarl] section missing - can't write contents", LEMON_LEVEL_CRITICAL
'
'        End If
'
'        .Save
'
'    End With
'
'End Sub

Public Function g_ConfigGet(ByVal Name As String) As String

    ' /* pre-set with default */

    If Not (mDefaults Is Nothing) Then _
        g_ConfigGet = mDefaults.ValueOf(Name)

Dim sz As String

    If Not (mConfig Is Nothing) Then
        If mConfig.Find(Name, sz) Then _
            g_ConfigGet = sz

    End If

End Function

Public Sub g_ConfigSet(ByVal Name As String, ByVal Value As String)

    If (mConfig Is Nothing) Then _
        Exit Sub

    mConfig.Update Name, Value
    g_WriteConfig

End Sub

Public Sub g_WriteConfig()

'    Debug.Print "g_WriteConfig: " & gNoWriteConfig

    If mConfigLocked Then
        g_Debug "g_WriteConfig(): config is locked - request queued"
        mWriteConfigOnUnlock = True
        Exit Sub

    End If

    g_Debug "g_WriteConfig(): writing to " & mSettings.File & "..."
    mSettings.Save

    ' /* V38 - always save to %appdata% */

'Dim sz As String
'
'    If Not g_GetSystemFolder(CSIDL_APPDATA, sz) Then
'        g_Debug "g_WriteConfig(): couldn't retrieve %appdata% location", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If

'    sz = g_MakePath(sz) & "full phat\snarl\etc"
'    create_directory_path sz
'
'Dim i As Long
'
'    With g_Settings
'        ' /* V38 - force the filename */
'        .File = g_MakePath(sz) & "v41.snarl"
'
'        i = .FindSection("snarl")
'        If i Then
'            With .SectionAt(i)
'                .Update "gfxdelta", CStr(g_Prefs.STEP_SIZE)
'                .Update "default_position", CStr(g_Prefs.default_position)
'                .Update "startinfo", g_Settings.BoolToStr(g_Prefs.show_msg_on_start)
'                .Update "startonlogon", g_Settings.BoolToStr(g_Prefs.run_on_logon)
'                .Update "default_style", g_Prefs.default_style
'
'                .Update "log_only", g_Settings.BoolToStr(g_Prefs.log_only)
'                .Update "sticky_snarls", g_Settings.BoolToStr(g_Prefs.sticky_snarls)
'                .Update "default_duration", CStr(g_Prefs.default_duration)
'
'                ' /* R2.04 */
'                .Update "font_smoothing", CStr(g_Prefs.font_smoothing)
'                .Update "melontype_contrast", CStr(g_Prefs.melontype_contrast)
'
'                ' /* R2.1 */
'                .Update "listen_for_json", g_Settings.BoolToStr(g_Prefs.listen_for_json)
'                .Update "listen_for_snarl", g_Settings.BoolToStr(g_Prefs.listen_for_snarl)
'                .Update "suppress_delay", CStr(g_Prefs.suppress_delay)
'                .Update "hotkey_prefs", CStr(g_Prefs.hotkey_prefs)
'                .Update "notify_on_first_register", g_Settings.BoolToStr(g_Prefs.notify_on_first_register)
'                ' /* R2.1d4 */
'                .Update "global_opacity", CStr(g_Prefs.global_opacity)
'                .Update "last_sound_folder", g_Prefs.last_sound_folder
'                .Update "show_tray_icon", g_Settings.BoolToStr(g_Prefs.show_tray_icon)
'                .Update "ignore_new_classes", g_Settings.BoolToStr(g_Prefs.ignore_new_classes)
'                ' /* R2.2 */
'                .Update "use_hotkey", g_Settings.BoolToStr(g_Prefs.use_hotkey)
'                .Update "do_not_disturb", g_Settings.BoolToStr(g_Prefs.do_not_disturb_)
'                .Update "idle_timeout", CStr(g_Prefs.idle_timeout)
'                .Update "margin_spacing", CStr(g_Prefs.margin_spacing)
'                .Update "use_dropshadow", g_Settings.BoolToStr(g_Prefs.use_dropshadow)
'                .Update "dropshadow_strength", CStr(g_Prefs.dropshadow_strength)
'                .Update "dropshadow_size", CStr(g_Prefs.dropshadow_size)
'                .Update "icon_theme", g_Prefs.icon_theme
'                .Update "auto_update", g_Settings.BoolToStr(g_Prefs.auto_update)
'
'
'                .Update "AgreeR22BetaUsage", g_Settings.BoolToStr(g_Prefs.AgreeBetaUsage)
'
'            End With
'
'            If Not (g_AppRoster Is Nothing) Then _
'                g_AppRoster.WriteConfig
'
'
''            With gRemoteConfig
''                .MakeEmpty
''                If gRemoteCount Then
''                    For i = 1 To gRemoteCount
''                        .Add "ip", gRemoteComputer(i).HostNameOrIp
''
''                    Next i
''                End If
''
''            End With

End Sub

Public Function g_Version() As String

    '& IIf(App.Revision <> 0, "." & App.Revision, "")
'    g_Version = App.Major & "." & App.Minor & IIf(App.Comments <> "", " " & App.Comments, "") & " (Build " & CStr(App.Revision) & ")"

    g_Version = App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")"

End Function

Public Sub g_SetRunning(ByVal IsRunning As Boolean, Optional ByVal Startup As Boolean)
Dim snarl_msg As Long
Dim i As Long
Dim rv As Long

    If g_IsRunning = IsRunning Then _
        Exit Sub

    ' /* this just updates the tray icon menu contents */

    If Not IsRunning Then
        ' /* tell our applications we've stopped */

        If Not (g_AppRoster Is Nothing) Then _
            g_AppRoster.SendAll SNARL_QUIT

    Else
        ' /* started running so set master flag now */

        g_IsRunning = True

    End If

    ' /* special case (sic) - if 'Startup' is set, show our welcome message now */

'    If Startup Then _
        i = snShowMessageEx("Welcome Message", "Welcome to Snarl!", _
                            "Snarl " & g_Version() & vbCrLf & App.LegalCopyright & vbCrLf & "http://www.fullphat.net", _
                            10, "http://www.google.com/ig/images/weather/chance_of_rain.gif")

    If Startup Then _
        i = snShowMessageEx("Welcome Message", "Welcome to Snarl!", _
                            "Snarl " & g_Version() & vbCrLf & App.LegalCopyright & vbCrLf & "http://www.fullphat.net", _
                            10, g_MakePath(App.Path) & "etc\icons\snarl.png")


'    Debug.Print "**: " & i
'    If i <> 0 Then
'        gfSetSnarlElement i, SNARL_ATTRIBUTE_ACK, "http://www.fullphat.net/"
'
'    End If

    ' /* send started/stopped broadcast */

    snarl_msg = snGetGlobalMsg()
    If snarl_msg > 0 Then
        g_Debug "g_SetRunning(): Broadcasting " & IIf(IsRunning, "SNARL_LAUNCHED", "SNARL_QUIT") & "..."
        SendMessageTimeout HWND_BROADCAST, snarl_msg, IIf(IsRunning, SNARL_LAUNCHED, SNARL_QUIT), ByVal 0&, SMTO_ABORTIFHUNG, 10, i

    Else
        g_Debug "g_SetRunning(): couldn't get snarl global message", LEMON_LEVEL_WARNING

    End If

    ' /* tell the extensions */

    If Not (g_ExtnRoster Is Nothing) Then _
        g_ExtnRoster.SendSnarlState IsRunning

    ' /* stopped running so set master flag here */

    If Not IsRunning Then _
        g_IsRunning = False

End Sub

Public Sub g_SetAutoRun(ByVal AutoRun As Boolean)

    If AutoRun Then
        add_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"

    Else
        rem_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"

    End If

    g_Prefs.run_on_logon = AutoRun

End Sub

Public Function gfRegisterAlert(ByVal AppName As String, ByVal Class As String, ByVal Flags As Long) As M_RESULT
Dim pa As TApp

    g_Debug "gfRegisterAlert('" & AppName & "' '" & Class & "' #" & g_HexStr(Flags) & ")", LEMON_LEVEL_PROC

    ' /* find the app */

    If Not g_AppRoster.Find(AppName, pa) Then
        g_Debug "gfRegisterAlert(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfRegisterAlert = M_NOT_FOUND
        Exit Function

    End If

    ' /* try to add it */

    gfRegisterAlert = pa.AddAlert(Class, "")

End Function

Public Function gfAddClass(ByVal pid As Long, ByVal Class As String, ByVal Flags As Long, ByVal Description As String) As M_RESULT
Dim pa As TApp

    g_Debug "gfAddClass('" & CStr(pid) & "' '" & Class & "' #" & g_HexStr(Flags) & ")", LEMON_LEVEL_PROC

    ' /* find the app */

    If Not g_AppRoster.FindByPid(pid, pa) Then
        g_Debug "gfAddClass(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfAddClass = M_NOT_FOUND
        Exit Function

    End If

    ' /* try to add it */

    gfAddClass = pa.AddAlert(Class, Description)

End Function

'Public Function globalEnableAlert(ByVal AppName As String, ByVal AlertName As String, ByVal Enabled As Boolean) As Boolean
'Dim i As Long
'Dim j As Long
'
'    i = globalFindAppByName(AppName)
'    If i Then
'        With g_Applet(i).Alerts
'            j = .IndexOf(AlertName)
'            If j Then _
'                .Update AlertName, IIf(Enabled, "1", "0")
'
'        End With
'
'    End If
'
'End Function

Public Function g_UTF8(ByVal str As String) As String

    g_UTF8 = trim(toUnicodeUTF8(g_utoa(str)))

End Function

Public Function g_GetUserFolderPath(ByRef Path As String) As Boolean
Dim sz As String

    If Not g_GetSystemFolder(CSIDL_APPDATA, sz) Then _
        Exit Function

    Path = g_MakePath(sz) & "full phat\snarl"
    g_GetUserFolderPath = True

End Function

Public Function g_GetUserFolder(ByRef Folder As storage_kit.Node, Optional ByVal AllUsers As Boolean = False) As Boolean
Dim sz As String

    If Not g_GetSystemFolder(IIf(AllUsers, CSIDL_COMMONAPPDATA, CSIDL_APPDATA), sz) Then _
        Exit Function

    Set Folder = New Node
    g_GetUserFolder = Folder.SetTo(g_MakePath(sz) & "full phat\snarl")

End Function

Public Function g_GetSystemFolderNode(ByVal Path As CSIDL_VALUES, ByRef Folder As storage_kit.Node) As Boolean
Dim sz As String

    If Not g_GetSystemFolder(Path, sz) Then _
        Exit Function

    Set Folder = New storage_kit.Node
    g_GetSystemFolderNode = Folder.SetTo(g_MakePath(sz))

End Function


Public Function gSetUpFontSmoothing(ByRef aView As mfxView, ByVal TextColour As Long, ByVal SmoothingColour As Long) As MFX_DRAWSTRING_FLAGS
Dim dw As Long

    aView.SetHighColour TextColour

    Select Case g_Prefs.font_smoothing
    Case E_MELONTYPE
        If SmoothingColour = 0 Then
            ' /* calculate it (only really works for dark colours at present) */
            aView.SetLowColour rgba(get_red(TextColour), _
                                    get_green(TextColour), _
                                    get_blue(TextColour), _
                                    (g_Prefs.melontype_contrast / 100) * 255)

        Else
            aView.SetLowColour SmoothingColour

        End If

        gSetUpFontSmoothing = MFX_SIMPLE_OUTLINE

    Case E_NONE
        aView.TextMode = MFX_TEXT_PLAIN

    Case E_ANTIALIAS
        aView.TextMode = MFX_TEXT_ANTIALIAS

    Case E_CLEARTYPE
        aView.TextMode = MFX_TEXT_CLEARTYPE

    Case E_WINDOWS_DEFAULT
        SystemParametersInfo SPI_GETFONTSMOOTHING, 0, dw, 0
        If dw = 0 Then
            ' /* none */
            aView.TextMode = MFX_TEXT_PLAIN

        Else
            ' /* enabled - but which type? */
            aView.TextMode = MFX_TEXT_ANTIALIAS     ' // assume antialias...
            If g_IsWinXPOrBetter() Then
                dw = 0
                SystemParametersInfo SPI_GETFONTSMOOTHINGTYPE, 0, dw, 0

                If dw = FE_FONTSMOOTHINGCLEARTYPE Then _
                    aView.TextMode = MFX_TEXT_CLEARTYPE

'                FE_FONTSMOOTHINGSTANDARD and

            End If

        End If

    End Select

End Function

Public Function gfSetAlertDefault(ByVal pid As Long, ByVal Class As String, ByVal Element As SNARL_ATTRIBUTES, ByVal Value As String) As M_RESULT
Dim pa As TApp
Dim pc As TAlert

    g_Debug "gfSetAlertDefault('" & pid & "' '" & Class & "' #" & CStr(Element) & " '" & Value & "')", LEMON_LEVEL_PROC

    If (g_AppRoster Is Nothing) Then
        g_Debug "gfSetAlertDefault(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_ABORTED
        Exit Function

    End If

    ' /* find the app */

    If Not g_AppRoster.FindByPid(pid, pa) Then
        g_Debug "gfSetAlertDefault(): App '" & pid & "' not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_NOT_FOUND
        Exit Function

    End If

    ' /* check the class  */

    If Not pa.FindAlert(Class, pc) Then
        g_Debug "gfSetAlertDefault(): alert class '" & Class & "' not found", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_NOT_FOUND
        Exit Function

    End If

    ' /* change the value */

    gfSetAlertDefault = M_OK

    Select Case Element

    Case SNARL_ATTRIBUTE_TITLE
        pc.DefaultTitle = Value

    Case SNARL_ATTRIBUTE_TEXT
        pc.DefaultText = Value

    Case SNARL_ATTRIBUTE_TIMEOUT
        pc.DefaultTimeout = Val(Value)

    Case SNARL_ATTRIBUTE_SOUND
        pc.DefaultSound = Value

    Case SNARL_ATTRIBUTE_ICON
        pc.DefaultIcon = Value

    Case SNARL_ATTRIBUTE_ACK
        pc.DefaultAck = Value

    Case Else
        g_Debug "gfSetAlertDefault(): unknown element '" & Element & "'", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_INVALID_ARGS

    End Select

End Function

Public Sub g_AddIgnoreLock()

    g_IgnoreLock = g_IgnoreLock + 1

End Sub

Public Sub g_RemIgnoreLock()

    g_IgnoreLock = g_IgnoreLock - 1

End Sub

Public Sub g_WriteToLog(ByVal Title As String, ByVal Text As String)
Dim sz As String
Dim n As Integer

    On Error Resume Next

    If Not g_GetUserFolderPath(sz) Then _
        Exit Sub

    n = FreeFile()
    Open g_MakePath(sz) & "snarl_log.txt" For Append As #n
    If Err.Number <> 0 Then _
        Exit Sub

    Print #n, CStr(Now()) & vbTab & Replace$(Title, vbCrLf, "/n") & vbTab & Replace$(Text, vbCrLf, "/n")
    Close #n

End Sub

Public Sub g_AddRemoteComputer(ByVal IPAddress As String)

    If (IPAddress = "") Or (gRemoteComputers Is Nothing) Then _
        Exit Sub

    ' /* should check for duplicates... */

Dim sz As String

    With gRemoteComputers
        If Not .Find(IPAddress, sz) Then
            .Add IPAddress, IPAddress
            frmAbout.bUpdateRemoteComputerList
            g_Debug "g_AddRemoteComputer(): added '" & IPAddress & "'"
            g_WriteConfig

        Else
            g_Debug "g_AddRemoteComputer(): '" & IPAddress & "' already in list"

        End If
    End With

End Sub

Public Function g_GetRemoteComputers() As String
Dim pe As ConfigEntry
Dim sz As String

    If (gRemoteComputers Is Nothing) Then _
        Exit Function

    With gRemoteComputers
        .Rewind
        Do While .GetNextEntry(pe)
            sz = sz & pe.Value & "|"

        Loop

    End With

    g_GetRemoteComputers = g_SafeLeftStr(sz, Len(sz) - 1)

End Function

Public Function g_GetRemoteComputersMenu() As OMMenu

    If (gRemoteComputers Is Nothing) Then _
        Exit Function

Dim pMenu As OMMenu
Dim pe As ConfigEntry

    Set pMenu = New OMMenu
    
    With gRemoteComputers
        .Rewind
        Do While .GetNextEntry(pe)
            pMenu.AddItem pMenu.CreateItem("ip>" & pe.Value, pe.Value)

        Loop

    End With
    
    If pMenu.CountItems > 0 Then _
        Set g_GetRemoteComputersMenu = pMenu

End Function

Public Sub g_AddDNDLock()

    g_Prefs.dnd_count = g_Prefs.dnd_count + 1

End Sub

Public Sub g_RemDNDLock()

    g_Prefs.dnd_count = g_Prefs.dnd_count - 1

End Sub

Public Function g_IsDNDModeEnabled() As Boolean

    ' /* DND mode is considered enabled if:
    '       1. the user has enabled it (g_Prefs.do_not_disturb is True)
    '       2. an app has enabled it (g_Prefs.dnd_count > 0)
    ' */

    If (g_Prefs.do_not_disturb_) Or (g_Prefs.dnd_count > 0) Then _
        g_IsDNDModeEnabled = True

End Function

Public Sub g_GetIconThemes()
Dim pn As storage_kit.Node

    ReDim gIconTheme(0)
    gIconThemes = 0

    If g_GetUserFolder(pn) Then _
        uGetIconThemes pn

    If g_GetUserFolder(pn, True) Then _
        uGetIconThemes pn

End Sub

Private Sub uGetIconThemes(ByRef Folder As storage_kit.Node)

    If Not (Folder.SetTo(g_MakePath(Folder.File) & "themes")) Then _
        Exit Sub

    If Not (Folder.IsFolder) Then _
        Exit Sub

Dim i As Long
Dim c As Long

    With Folder
        .ReadContents
        c = .CountNodes
        If c Then
            For i = 1 To c
                If .NodeAt(i).IsFolder Then _
                    uGetIconTheme .NodeAt(i)

            Next i
        End If
    End With

End Sub

Private Sub uGetIconTheme(ByRef Folder As storage_kit.Node)
Dim pn As storage_kit.Node

    Set pn = New storage_kit.Node
    If Not (pn.SetTo(g_MakePath(Folder.File) & "icons")) Then _
        Exit Sub

    If Not (pn.IsFolder) Then _
        Exit Sub

    gIconThemes = gIconThemes + 1
    ReDim Preserve gIconTheme(gIconThemes)
    With gIconTheme(gIconThemes)
        .Name = Folder.Filename
        .Path = pn.File
        .IconFile = g_MakePath(Folder.File) & "theme.png"

    End With

End Sub

Public Function g_GetIconThemePath(ByVal Name As String, ByRef Path As String) As Boolean
Dim i As Long

    If gIconThemes Then
        For i = 1 To gIconThemes
            If LCase$(gIconTheme(i).Name) = LCase$(Name) Then
                Path = g_MakePath(gIconTheme(i).Path)
                g_GetIconThemePath = True

            End If
        Next i
    End If

End Function

Public Function g_IsValidImage(ByRef Image As MImage) As Boolean

    If (Image Is Nothing) Then _
        Exit Function

    g_IsValidImage = ((Image.Width > 0) And (Image.Height > 0))

End Function

Public Function g_DoSchemePreview(ByVal Name As String, ByVal Scheme As String, ByVal Flags As Long, ByVal Timeout As Long, ByVal Percent As Integer) As M_RESULT

    g_DoSchemePreview = M_FAILED
    If (g_NotificationRoster Is Nothing) Or (g_StyleRoster Is Nothing) Then _
        Exit Function

Dim pStyle As TStyle

    g_DoSchemePreview = M_INVALID_ARGS
    If Not g_StyleRoster.Find(Name, pStyle) Then _
        Exit Function

    If Scheme = "" Then
        Scheme = "<Default>"

    Else
        ' /* supplied scheme must exist */
        If pStyle.SchemeIndex(Scheme) = 0 Then _
            Exit Function

    End If

Dim szText As String

    If (Percent > 0) And (Percent <= 100) Then
        szText = CStr(Percent)

    Else
        If Scheme <> "<Default>" Then _
            szText = " using the " & Scheme & " scheme"


        If (Flags And 1) Then
            szText = "This is a preview of a priority notification using the " & pStyle.Name & " style" & szText

        Else
            szText = "This is a preview of the " & pStyle.Name & " style" & szText

        End If

    End If

Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        .Title = "Scheme Preview"
        .Text = szText
        .Timeout = Timeout
        .IconPath = IIf(pStyle.IconPath = "", g_MakePath(App.Path) & "etc\icons\style_preview.png", pStyle.IconPath)
        .StyleToUse = pStyle.Name & "/" & LCase$(Scheme)
        .Position = E_START_DEFAULT_POS
        .Priority = (Flags And 1)

    End With

    If g_NotificationRoster.Add(New TAlert, pInfo) <> 0 Then
        g_DoSchemePreview = M_OK

    Else
        g_DoSchemePreview = M_FAILED

    End If

End Function

Public Function g_GetSafeTempIconPath() As String
Dim sz As String
Dim c As Long

    sz = String$(MAX_PATH + 1, 0)
    GetTempPath MAX_PATH, sz
    sz = g_TrimStr(sz)
    If sz = "" Then _
        Exit Function

    sz = g_MakePath(sz)

    c = 1
    Do While g_Exists(sz & "snarl-icon" & CStr(c))
        c = c + 1

    Loop

    g_GetSafeTempIconPath = sz & "snarl-icon" & CStr(c)

End Function

Public Sub g_LockConfig()

    mConfigLocked = True

End Sub

Public Sub g_UnlockConfig()

    mConfigLocked = False

    If mWriteConfigOnUnlock Then
        mWriteConfigOnUnlock = False
        g_WriteConfig

    End If

End Sub

'Public Function new_Class(ByVal Priority As Boolean) As TAlert
'
'    Set new_Class = New TAlert
'    new_Class.bSpecialInit "_id", "_desc", Priority
'
'End Function

Public Function g_IsSticky() As Boolean

    g_IsSticky = (g_Prefs.sticky_snarls) Or (gIsIdle)

End Function

Public Function g_GetStylePath(ByVal StyleToUse As String) As String

    g_GetStylePath = g_MakePath(App.Path) & "etc\default_theme\"
    If (g_StyleRoster Is Nothing) Then _
        Exit Function

    If StyleToUse = "" Then _
        StyleToUse = g_ConfigGet("default_style")

    StyleToUse = style_GetStyleName(StyleToUse)
    If StyleToUse = "" Then _
        Exit Function                               ' // indicates a problem!

Dim pStyle As TStyle

    If g_StyleRoster.Find(style_GetStyleName(StyleToUse), pStyle) Then _
        g_GetStylePath = pStyle.Path

End Function

Public Function g_RemoveForwarder(ByVal UID As Long)

    Debug.Print "STUB: g_RemoveForwarder(" & CStr(UID) & ")"

End Function
