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
'    StyleToUse As String
    StyleName As String                 ' // Split
    SchemeName As String                ' // Split
    DefaultAck As String
    Position As E_START_POSITIONS
    Token As Long
    Priority As Long                    ' // V41: <0 = low, 0 = normal, >0 = high
    Value As String                     ' // V41: freeform value which will negate the need to use the Text field
                                        '         thoughts are the value can encapsulate the format it's sent in
                                        '         e.g. 45%, 2.3466, $5.00, etc. it's up to the style to determine
                                        '         how/if it's displayed
    DateStamp As Date                   ' // V41: when it was added to the Notification Roster
    Icon As mfxBitmap                   ' // V41 (R2.31): note it's an mfxBitmap, not an MImage!
    Sender As String
    Class As String

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


Public Type T_SNARL_ADMIN
    HideIcon As Boolean                 ' // hides the tray icon (over-rules undoc'd setting in .snarl file)
    InhibitPrefs As Boolean             ' // completely blocks access to prefs panel
    InhibitQuit As Boolean              ' // can't quit Snarl using menu
    InhibitMenu As Boolean              ' // right-click tray icon does nothing
    TreatSettingsAsReadOnly As Boolean  ' // don't write settings

End Type

Public gSysAdmin As T_SNARL_ADMIN


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
    ' /* MSG_SHOW_PREFS expanded in R2.3 to use wParam and lParam, as follows:
    '       wParam                                                          lParam
    '          0        Display Preferences Panel                           0
    '          1        Install style or extension                          atom of registered string containing style or extension
    '          2        Configure style or extension                        ("")
    ' */


    ' /* MSG_QUIT is deprecated in favour of using standard WM_CLOSE.  Handling of MSG_QUIT is retained in R2.3 */
Public Const MSG_QUIT = WM_USER + 81

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

    ' /* used by the missed notifications panel - should be consolidated with other similar structs */

'Public Type G_NOTIFICATION_CONTENT
'    Title As String
'    Text As String
'    Icon As mfxBitmap
'    Timeout As Long
'    Ack As String
'    Timestamp As Date
'    Sender As String
'    Class As String
'    ' /* R2.31 */
'    hWndReply As Long
'    uMsgReply As Long
'
'End Type


Public Type T_CONFIG

'    run_on_logon As Boolean
    font_smoothing As E_FONTSMOOTHING
'    suppress_delay As Long          ' // in ms
'    hotkey_prefs As Long            ' // MAKELONG(mods,key)
'    last_sound_folder As String
'    use_hotkey As Boolean
    UserDnD As Boolean      ' // not persitent: user-controlled DND setting
    SysDnDCount As Long             ' // set using WM_MANAGE_SNARL
    MissedCountOnDnD As Long
    use_dropshadow As Boolean
    last_update_check As Date
    AgreeBetaUsage As Boolean

    ' /* R2.31 */
    SnarlConfigPath As String           ' // path (UNC or other) to Snarl config folder (should contain /etc and other folders)
    SnarlConfigFile As String           ' // not persistent; just a handy copy

End Type

Public gPrefs As T_CONFIG

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
'Public gIsIdle As Boolean

Public gLastNotification As Date    ' // V41.47 - last notification timestamp

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

'Public gLastErr As SNARL_STATUS_41      ' // V41 api

Public gCurrentLowPriorityId As Long

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub CoFreeUnusedLibrariesEx Lib "ole32" (ByVal dwUnloadDelay As Long, ByVal dwReserved As Long)
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long

'Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Private Declare Sub CoFreeLibrary Lib "ole32.dll" (ByVal hInst As Long)

Public Sub Main()
Dim szArg() As String
Dim l As Long

'Dim iv As MVersionInfo
'
'    Set iv = web_resource.globals
'    MsgBox iv.Version & "." & iv.Revision

    ' /* get comctl.dll loaded up for our XP manifest... */

    g_InitComCtl

    ' /* is Snarl already running? */

    l = FindWindow(WINDOW_CLASS, "Snarl")


    szArg = Split(Command$, " ")
    If UBound(szArg) > -1 Then
        ' /* command specified, but which? */
        Select Case LCase$(szArg(0))
        Case "-quit"
            If l <> 0 Then _
                SendMessage l, WM_CLOSE, 0, ByVal 0&

            Exit Sub

        Case "-debug"
            gDebugMode = True


        Case "-install"
            ' /* must have one further arg: style engine or extension to install */

            If (UBound(szArg()) = 1) And (l <> 0) Then
                PostMessage l, MSG_SHOW_PREFS, 1, ByVal RegisterClipboardFormat(szArg(1))
                Exit Sub

            End If

        Case "-configure"
            ' /* must have one further arg: style or extension to install */

            If (UBound(szArg()) = 1) And (l <> 0) Then
                PostMessage l, MSG_SHOW_PREFS, 2, ByVal RegisterClipboardFormat(szArg(1))
                Exit Sub

            End If

        End Select

    End If

    If l <> 0 Then
        ' /* Snarl is already running (and no useful command-line arg specified) */
        PostMessage l, MSG_SHOW_PREFS, 0, ByVal 0&    ' // tell the running instance to show its ui...
        Exit Sub

    End If

    App.TaskVisible = False

    ' /* V38.133 - enable debug mode if switch present */
    ' /* V40.18 - or if either CTRL key is held down */

    gDebugMode = gDebugMode Or (g_IsPressed(VK_LCONTROL)) Or (g_IsPressed(VK_RCONTROL))

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

    ' /* R2.31: set our flags */

Dim dwFlags As Long

    If gDebugMode Then _
        dwFlags = dwFlags Or &H80000000
    

    SetProp ghWndMain, "_flags", dwFlags

    ' /* R2.31: init last notification timestamp */
    
    gLastNotification = Now()


    ' /* load up some required bits and bobs */

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

    ' /* R2.31 - pre-set our config folder */

    gPrefs.SnarlConfigPath = App.Path                ' // fail-safe

Dim sz As String

    If g_GetUserFolderPath(sz) Then
        gPrefs.SnarlConfigPath = sz                  ' // standard location

    Else
        g_Debug "uSetConfigFile(): %APP_DATA% path not found", LEMON_LEVEL_CRITICAL

    End If

    ' /* R2.31 - look for a local sysconfig.ssl and get its target  */

Dim pSysConfig As CConfFile

    Set pSysConfig = New CConfFile
    With pSysConfig
        If .SetTo(g_MakePath(App.Path) & "sysconfig.ssl", True) Then
            g_Debug "Main: local sysconfig.ssl exists, querying..."
            sz = g_RemoveQuotes(.ValueOf("target"))

            ' /* R2.31 - quick check on folder structure */

            If g_IsFolder(g_MakePath(sz) & "etc") Then
                gPrefs.SnarlConfigPath = sz

            Else
                g_Debug "Main: config path '" & sz & "' is invalid", LEMON_LEVEL_CRITICAL

            End If
        End If
    End With

    gPrefs.SnarlConfigPath = g_MakePath(gPrefs.SnarlConfigPath)
    g_Debug "Main: config path is '" & gPrefs.SnarlConfigPath & "'"

    ' /* fix up the .snarl path */

    gPrefs.SnarlConfigFile = gPrefs.SnarlConfigPath & "etc\config41.snarl"
    g_Debug "Main: .snarl path is '" & gPrefs.SnarlConfigFile & "'"

    ' /* R2.31: register our config path as a global atom and store the atom as a window property */

    SetProp ghWndMain, "_config_path", RegisterWindowMessage(gPrefs.SnarlConfigPath)

    ' /* do we have a snarl.admin file to load? */

Dim szName As String
Dim szData As String

    Set pSysConfig = New CConfFile
    If pSysConfig.SetTo(gPrefs.SnarlConfigPath & "etc\snarl.admin", True) Then

        g_Debug "Main: loaded admin settings from '" & gPrefs.SnarlConfigPath & "snarl.admin" & "'"

        With pSysConfig
            .Rewind
            Do While .GetEntry(szName, szData)
                g_Debug "Main: '" & szName & "'='" & szData & "'"

            Loop

        End With

        With gSysAdmin
            .HideIcon = (pSysConfig.ValueOf("HideIcon") = "1")
            .InhibitPrefs = (pSysConfig.ValueOf("InhibitPrefs") = "1")
            .TreatSettingsAsReadOnly = (pSysConfig.ValueOf("TreatSettingsAsReadOnly") = "1")
            .InhibitMenu = (pSysConfig.ValueOf("InhibitMenu") = "1")
            .InhibitQuit = (pSysConfig.ValueOf("InhibitQuit") = "1")

        End With

    Else
        g_Debug "Main: no admin settings file"

    End If

    ' /* get settings */

    If Not g_ConfigInit() Then
        g_Debug "main(): new/clean installation..."
        g_ConfigSet "step_size", "1"

    End If

'    ' /* R2.4: managed style settings */
'
'    Set gStyleSettings = New CConfFile
'    gStyleSettings.SetTo gPrefs.SnarlConfigPath & "etc\.stylesettings"


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
    g_SetAutoRun2

    ' /* get style packs */

    g_Debug "Main(): Starting style roster..."
    Set g_StyleRoster = New TStyleRoster
    melonLibInit g_StyleRoster
    melonLibOpen g_StyleRoster

    ' /* get icon themes */

    g_GetIconThemes

    ' /* set master running flag */

    g_SetRunning True, False

    ' /* display welcome message */

    If g_ConfigGet("show_msg_on_start") = "1" Then _
        snShowMessageEx "Welcome Message", "Welcome to Snarl!", _
                        "Snarl " & g_Version() & vbCrLf & App.LegalCopyright & vbCrLf & "http://www.fullphat.net", _
                        -1, g_MakePath(App.Path) & "etc\icons\snarl.png"

    ' /* get extensions */

    g_Debug "Main(): Starting extensions roster..."
    Set g_ExtnRoster = New TExtensionRoster
    melonLibInit g_ExtnRoster
    melonLibOpen g_ExtnRoster

#If BETA_REL = 1 Then

Dim pBetaPanel As TBetaPanel

    If (Not gPrefs.AgreeBetaUsage) Then
        Set pBetaPanel = New TBetaPanel
        pBetaPanel.Go

    End If

#End If

    If g_ConfigGet("auto_update") = "1" Then
        g_Debug "Main(): Doing auto-update check..."
        Set myUpdateCheck = New TAutoUpdate
        If myUpdateCheck.Check(False) Then _
            g_Debug "Main(): auto-update check initiated"

    Else
        g_Debug "Main(): Auto-update is disabled"

    End If

    ' /* done */

    frmAbout.bReadyToRun

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
        PostMessage HWND_BROADCAST, uSnarlGlobal, SNARL_QUIT, ByVal CLng(App.Major)

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

End Sub

Public Function g_ConfigInit() As Boolean

    On Error Resume Next

    With gPrefs
'        .run_on_logon = True
        .font_smoothing = E_MELONTYPE
'        .suppress_delay = 2000
'        If g_GetSystemFolder(CSIDL_PERSONAL, sz) Then _
            .last_sound_folder = sz

        .UserDnD = False
        .use_dropshadow = True

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
        .Add "ignore_new_classes", "0"      ' // new alert classes are always enabled by default

        ' /* R2.2 */
        .Add "use_hotkey", "1"
        .Add "do_not_disturb", "0"
        .Add "idle_timeout", "300"          ' // i.e. 5 minutes
        .Add "margin_spacing", "0"
        .Add "use_dropshadow", "1"
        .Add "dropshadow_strength", "88"    ' // is a %
        .Add "dropshadow_size", "10"
        .Add "icon_theme", ""

        ' /* R2.3 */
        .Add "auto_update", "1"
        .Add "enable_sounds", "1"
        .Add "use_style_sounds", "1"
        .Add "prefer_style_sounds", "0"
        .Add "default_normal_sound", ""
        .Add "default_priority_sound", ""
        .Add "use_style_icons", "1"
        .Add "auto_sticky_on_screensaver", "1"
        .Add "show_timestamp", "0"

        ' /* R2.4: style-usable settings are prefixed with 'style.' */
        
        .Add "style.overflow_limit", "7"

    End With

    ' /* attempt to load the config file */

'    MsgBox gSysAdmin.RemoteSnarlFile

Dim i As Long

    Set mSettings = New ConfigFile
    With mSettings
        .File = gPrefs.SnarlConfigFile
        .Load

        i = .FindSection("general")
        If i = 0 Then
            Set mConfig = .AddSectionObj("general")
            If Not gSysAdmin.TreatSettingsAsReadOnly Then _
                .Save

        Else
            Set mConfig = .SectionAt(i)

        End If

        i = .FindSection("remote_computers")
        If i = 0 Then
            Set gRemoteComputers = .AddSectionObj("remote_computers")
            If Not gSysAdmin.TreatSettingsAsReadOnly Then _
                .Save

        Else
            Set gRemoteComputers = .SectionAt(i)

        End If

    End With

    g_ConfigInit = (Val(g_ConfigGet("step_size")) > 0)

End Function

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

    ' /* R2.31: only if admin says we can... */

    If gSysAdmin.TreatSettingsAsReadOnly Then
        g_Debug "g_WriteConfig(): configuration as been set as read only by system administrator", LEMON_LEVEL_WARNING

    Else
        g_Debug "g_WriteConfig(): writing to " & mSettings.File & "..."
        mSettings.Save

    End If

End Sub

Public Function g_Version() As String

    '& IIf(App.Revision <> 0, "." & App.Revision, "")
'    g_Version = App.Major & "." & App.Minor & IIf(App.Comments <> "", " " & App.Comments, "") & " (Build " & CStr(App.Revision) & ")"

    g_Version = App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")"

End Function

Public Sub g_SetRunning(ByVal IsRunning As Boolean, Optional ByVal Broadcast As Boolean = True)

    If g_IsRunning = IsRunning Then _
        Exit Sub

Dim dw As Long

    If IsRunning Then
        ' /* set master flag *first* */
        g_IsRunning = True

        ' /* tell the extensions */
        If Not (g_ExtnRoster Is Nothing) Then _
            g_ExtnRoster.SendSnarlState True

        ' /* tell our applications we're starting */
        If Not (g_AppRoster Is Nothing) Then _
            g_AppRoster.SendAll SNARL_LAUNCHED

'        If Broadcast Then
'            ' /* send started broadcast */
'            g_Debug "g_SetRunning(): Broadcasting SNARL_LAUNCHED..."
'            SendMessageTimeout HWND_BROADCAST, snGetGlobalMsg(), SNARL_LAUNCHED, ByVal CLng(App.Major), SMTO_ABORTIFHUNG, 500, dw
'
'        End If

    Else

'        If Broadcast Then
'            ' /* send stopped broadcast */
'            g_Debug "g_SetRunning(): Broadcasting SNARL_QUIT..."
'            SendMessageTimeout HWND_BROADCAST, snGetGlobalMsg(), SNARL_QUIT, ByVal CLng(App.Major), SMTO_ABORTIFHUNG, 500, dw
'
'        End If

        ' /* close all notifications */
        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.CloseMultiple 0

        ' /* tell the extensions */
        If Not (g_ExtnRoster Is Nothing) Then _
            g_ExtnRoster.SendSnarlState False

        ' /* tell our applications we've stopped */
        If Not (g_AppRoster Is Nothing) Then _
            g_AppRoster.SendAll SNARL_QUIT

        ' /* set master flag last */
        g_IsRunning = False

    End If

End Sub

Public Sub g_SetAutoRun2()
Dim bAutoRun As Boolean

    bAutoRun = CBool(g_ConfigGet("run_on_logon"))

    If bAutoRun Then
        add_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"

    Else
        rem_registry_startup_item "Snarl", g_MakePath(App.Path) & LCase$(App.EXEName) & ".exe"

    End If

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

    Select Case gPrefs.font_smoothing
    Case E_MELONTYPE
        If SmoothingColour = 0 Then
            ' /* calculate it (only really works for dark colours at present) */
            aView.SetLowColour rgba(get_red(TextColour), _
                                    get_green(TextColour), _
                                    get_blue(TextColour), _
                                    (Val(g_ConfigGet("melontype_contrast")) / 100) * 255)

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

    gPrefs.SysDnDCount = gPrefs.SysDnDCount + 1

    If (g_IsDNDModeEnabled) And (Not (g_NotificationRoster Is Nothing)) Then _
        gPrefs.MissedCountOnDnD = g_NotificationRoster.CountMissed

End Sub

Public Sub g_RemDNDLock()

    gPrefs.SysDnDCount = gPrefs.SysDnDCount - 1
    g_CheckMissed

End Sub

Public Sub g_CheckMissed()
Dim iMissed As Long
Dim iToken As Long

    If (g_NotificationRoster Is Nothing) Then _
        Exit Sub

    ' /* if DnD mode is now disabled and there were some missed notifications, show notification
    '    explaining this */

    iMissed = g_NotificationRoster.CountMissed - gPrefs.MissedCountOnDnD

    If (Not g_IsDNDModeEnabled()) And (iMissed > 0) Then
        iToken = snShowMessageEx("", "While you were away...", _
                                    "You missed " & CStr(iMissed) & " notification" & IIf(iMissed > 1, "s", ""), _
                                    -1, _
                                    g_MakePath(App.Path) & "etc\icons\snarl.png")

        If iToken > 0 Then _
            snChangeAttribute iToken, SNARL_ATTRIBUTE_ACK, "!show_missed_panel"

    End If

End Sub

Public Function g_IsDNDModeEnabled() As Boolean

    ' /* DND mode is considered enabled if:
    '       1. the user has enabled it (gPrefs.do_not_disturb is True)
    '       2. an app has enabled it (gPrefs.SysDnDCount > 0)
    ' */

    If (gPrefs.UserDnD) Or (gPrefs.SysDnDCount > 0) Then _
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
        .StyleName = pStyle.Name
        .SchemeName = LCase$(Scheme)
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

    g_IsSticky = (g_ConfigGet("sticky_snarls") = "1") 'Or (gIsIdle)

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

Public Function g_SettingsPath() As String

    If Not (mSettings Is Nothing) Then _
        g_SettingsPath = mSettings.File

End Function

Public Sub g_ProcessAck(ByVal Ack As String)
Dim arg() As String
Dim argC As Long

    g_Debug "g_ProcessAck(): ACK is '" & Ack & "'", LEMON_LEVEL_INFO

    If Left$(Ack, 1) = "!" Then
        ' /* bang command */

        arg = Split(Right$(Ack, Len(Ack) - 1), " ")
        argC = UBound(arg) + 1

        Debug.Print "g_ProcessAck(): " & Ack & " == " & argC

        Select Case LCase$(arg(0))
        Case "show_missed_panel"
            If Not (g_NotificationRoster Is Nothing) Then _
                g_NotificationRoster.ShowMissedPanel

        Case "cfg"
            If argC > 0 Then _
                frmAbout.DoAppConfigBySignature arg(1)

        End Select

    Else
        ' /* treat as launchable */
        ShellExecute frmAbout.hWnd, vbNullString, Ack, vbNullString, vbNullString, SW_SHOW

    End If

End Sub

Public Sub gSetLastError(ByVal Error As SNARL_STATUS_41)

    SetProp ghWndMain, "last_error", Error

End Sub


'        MsgBox "As this is the first time Snarl has been run, I need to test the drawing" & vbCrLf & _
'               "speed of your PC.  This test will only take a few seconds, if that.", _
'               vbOKOnly Or vbInformation, "Snarl Graphics Test"
'
'        l = GetTickCount()
'        g_Debug "main(): GFXCHK: starting graphics test (tick_count=" & CStr(l) & ")"
'
'Dim pInfo As T_NOTIFICATION_INFO
'
'        With pInfo
'            .Title = "Snarl Graphics Test"
'            .Text = "Test Message"
'            .Timeout = 1
'            .StyleToUse = ""            ' // the scheme
'
'        End With
'
'        With New CSnarlWindow
'            .Create New TAlert, pInfo, New TInternalStyle, 0, ""
'            '"Snarl Graphics Test", "Test message", 1, "", 0, 0, 0, "", New TInternalStyle, "", 0
'            .Quit
'
'        End With
'
'Dim dStep As Double
'
'        l = GetTickCount() - l
'        g_Debug "main(): GFXCHK: completed graphics test (tick_count=" & CStr(GetTickCount()) & ")"
'
'        dStep = l / 422#
'
'        If dStep < 1# Then _
'            dStep = 1#
'
'        g_Debug "main(): GFXCHK: delta=" & l & " actual=" & CStr(l / 422#) & " step=" & dStep
'
'        MsgBox "Graphics test complete.  Snarl has calculated a factor of " & Format$(dStep, "0.00") & vbCrLf & _
'               "to use when displaying messages.", vbOKOnly Or vbInformation, "Test Complete"

'Public Function g_StyleConfigGet(ByVal Name As String, Optional ByVal Default As String) As String
'
'    ' /* pre-set with default */
'
'    g_StyleConfigGet = mDefaults.ValueOf(Name)
'
'Dim sz As String
'
'    If Not (mConfig Is Nothing) Then
'        If mConfig.Find(Name, sz) Then _
'            g_ConfigGet = sz
'
'    End If
'
'End Function
'
''Public Sub g_ConfigSet(ByVal Name As String, ByVal Value As String)
''
''    If (mConfig Is Nothing) Then _
''        Exit Sub
''
''    mConfig.Update Name, Value
''    g_WriteConfig
''
''End Sub

