Attribute VB_Name = "mSnarl"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mSnarl.bas
    '/
    '/  Description:    Global functions and declarations
    '/
    '/  © 2004-2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* these are used by the deprecated SNARL_GET_VERSION and for GNTP responses */
Public Const APP_VER = 2
Public Const APP_SUB_VER = 4

Public Const GNTP_DEFAULT_PORT = 23053
Public Const SNP_DEFAULT_PORT = 9887
Public Const JSON_DEFAULT_PORT = 9889

Public Declare Sub CoFreeUnusedLibrariesEx Lib "ole32" (ByVal dwUnloadDelay As Long, ByVal dwReserved As Long)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Sub ShellAbout Lib "SHELL32.DLL" Alias "ShellAboutA" (ByVal hWndOwner As Long, ByVal lpszAppName As String, ByVal lpszMoreInfo As String, ByVal hIcon As Long)
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Const WINDOW_CLASS = "w>Snarl"

Public Const WM_SNARL_INIT = WM_USER + 1
Public Const WM_SNARL_QUIT = WM_USER + 2
Public Const WM_SNARL_TRAY_ICON = WM_USER + 3
'Public Const WM_SNARL_NOTIFY_RUNNING = WM_USER + 4

'Public Const WM_REMOTENOTIFY = WM_USER + 9              ' // frmAbout: remote notifications
Public Const WM_INSTALL_SNARL = WM_USER + 12            ' // frmAbout: snarl update available

Public Const TIMER_UPDATES = 32

    ' /* Snarl app class id's */

Public Const SNARL_CLASS_GENERAL = "_WLC"
Public Const SNARL_CLASS_APP_UNREG = "_APU"
Public Const SNARL_CLASS_APP_REG = "_APR"
Public Const SNARL_CLASS_JSON = "_ANJ"
Public Const SNARL_CLASS_ANON_NET = "_ANN"
Public Const SNARL_CLASS_ANON = "_ANL"
'Public Const SNARL_CLASS_LOW_PRIORITY = "_LOW"
'Public Const SNARL_CLASS_SYSTEM = "_SYS"

    ' /* internal notification flags */

Public Enum E_NOTIFICATION_FLAGS
    NF_REMOTE = &H80000000
    NF_SECURE = &H40000000
    NF_IS_GNTP = &H20000000           ' // R2.4.2: GNTP-based notification

    NF_MERGE = &H1000

    ' /* bottom 8 bits used for api version (V42 onwards) */
    NF_API_MASK = &HFF&

End Enum

    ' /* master notification structure, as used by the notification roster */

Public Type T_NOTIFICATION_INFO
    Pid As Long
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
    DefaultAck As String                ' // Known as "callback" from R2.4 DR7
    Position As E_START_POSITIONS
    Token As Long
    ' /* V41 */
    Priority As Long                    ' // V41: <0 = low, 0 = normal, >0 = high
    Value As String                     ' // V41: freeform value which will negate the need to use the Text field
                                        '         thoughts are the value can encapsulate the format it's sent in
                                        '         e.g. 45%, 2.3466, $5.00, etc. it's up to the style to determine
                                        '         how/if it's displayed
    DateStamp As Date                   ' // V41: when it was added to the Notification Roster
'    Icon As mfxBitmap                   ' // V41 (R2.31): note it's an mfxBitmap, not an MImage!
'    Sender As String
'    Class As String
    ' /* V42 */
    Flags As SNARL41_NOTIFICATION_FLAGS ' // V41 (R2.4): new flags
    OriginalContent As String           ' // V41 (R2.4): as passed from external source
    LastUpdated As Date                 ' // time last changed
    Socket As CSocket                   ' // reply socket (SNP2.0 native only)
    IntFlags As E_NOTIFICATION_FLAGS    ' // internal notification flags
    RemoteHost As String                ' // sender (as string) for remote connections that do not have reply sockets
    ClassObj As TAlert                  ' // object
    CustomUID As String                 ' // R2.4 DR7: custom UID (set during <notify>)
    Actions As BTagList                 ' // R2.4 DR7: should have been here all along
    APIVersion As Long                  ' // R2.4.1: will be 42 for V42, 0 for everything prior to it

End Type

'Public Type T_NOTIFICATION_EXTRA
'    IsRemoteApp As Boolean
'    ReplySocket As CSocket
'    RemoteAddr As String                ' // ip address
'    APIVersion As Long
'
'End Type


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

    ' /* internal registered application structure */

Public Type T_SNARL_APP
    Name As String
    hWnd As Long
    uMsg As Long
    Pid As Long                 ' // V38 (for V39)
    Icon As String              ' // R1.6 - path to application icon (if empty we use window icon)
    LargeIcon As String         ' // V38 (private for now) - path to large icon
    Token As Long               ' // V41
    Signature As String         ' // V41 - MIME string
    Flags As SNARLAPP_FLAGS     ' // V41
    Password As String          ' // V42 - non-persistent (so the app can generate a new one each time)
    IsRemote As Boolean         ' // V42 - remotely registered
    Timestamp As Date           ' // V42.21: set when added

End Type

    ' /* internal Snarl admin structure */

Public Type T_SNARL_ADMIN
    HideIcon As Boolean                 ' // hides the tray icon (over-rules undoc'd setting in .snarl file)
    InhibitPrefs As Boolean             ' // completely blocks access to prefs panel
    InhibitQuit As Boolean              ' // can't quit Snarl using menu
    InhibitMenu As Boolean              ' // right-click tray icon does nothing
    TreatSettingsAsReadOnly As Boolean  ' // don't write settings

End Type

Public gSysAdmin As T_SNARL_ADMIN
Public gExtDetailsToken As Long
Public gStyleEngineDetailsToken As Long

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

Public bm_Close As MImage
Public bm_Menu As MImage
Public bm_Actions As MImage
Public bm_HasActions As MImage
Public bm_Remote As MImage
Public bm_Secure As MImage
Public bm_IsSticky As MImage
Public bm_Priority As MImage

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
'    E_DURATION_DEFAULT = 0
    E_DURATION_APP_DECIDES = 1
    E_DURATION_CUSTOM           ' // "custom_timeout" contains value in seconds

End Enum

    ' /* master controls */
Public g_IsRunning  As Boolean
Public g_IsQuitting As Boolean

    ' /* rosters */
Public g_ExtnRoster As TExtensionRoster
Public g_StyleRoster As TStyleRoster
Public g_AppRoster As TApplicationRoster
Public g_NotificationRoster As TNotificationRoster

'Public Enum E_FONTSMOOTHING
'    E_MELONTYPE
'    E_NONE
'    E_ANTIALIAS
'    E_CLEARTYPE
'    E_WINDOWS_DEFAULT
'
'End Enum

Private Const SNARL_XXX_GLOBAL_MSG = "SnarlGlobalEvent"

Public Type T_CONFIG

'    run_on_logon As Boolean
'    font_smoothing As E_FONTSMOOTHING
'    suppress_delay As Long          ' // in ms
'    hotkey_prefs As Long            ' // MAKELONG(mods,key)
'    last_sound_folder As String
'    use_hotkey As Boolean
'    UserDnD As Boolean      ' // not persitent: user-controlled DND setting
'    SysDnDCount As Long             ' // set using WM_MANAGE_SNARL
'    MissedCountOnDnD As Long
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

'Public g_IgnoreLock As Long         ' // if >0 don't alert when app registers - overrides class setting
Public gSelectedClass As TAlert
Public gDebugMode As Boolean
'Public mAwayCount As Long           ' // R2.4 DR8: renamed and reimplemented

Public gLastNotification As Date    ' // V41.47 - last notification timestamp

'Public Type G_REMOTE_COMPUTER
'    IsHostName As Boolean
'    HostNameOrIp As String
'
'End Type

Public ghWndMain As Long
Public gUpdateFilename As String        ' // name of the update file to download

Dim mUpdateCheck As TAutoUpdate
Dim mBetaUpdateCheck As TAutoUpdate

'Public gCurrentLowPriority As T_NOTIFICATION_INFO       ' // only one can be on-screen at any one time
Public gSnarlToken As Long              ' // when Snarl registers with itself
Public gSnarlPassword As String         ' // created on the fly

'Dim mDoNotDisturbLock As Long           ' // >0 means enabled, <=0 means disabled
'Public gNotificationMenuOpen As Boolean

    ' /* R2.4 DR8 */

Public Enum SP_PRESENCE_FLAGS
    ' /* Away flags occupy bottom 16 bits */
    SP_AWAY_USER_IDLE = 1
    SP_AWAY_COMPUTER_LOCKED = 2
    SP_AWAY_SCREENSAVER_ACTIVE = 8
    SP_AWAY_MASK = &HFFFF&

    ' /* DnD flags occupy top 16 bits */
    SP_DND_FULLSCREEN_APP = &H10000
    SP_DND_USER = &H20000                       ' // from the tray icon menu
    SP_DND_EXTERNAL = &H40000                   ' // for future use
    SP_DND_MASK = &HFFFF0000

End Enum

Dim mPresFlags As SP_PRESENCE_FLAGS

Public Enum E_PRESENCE_ACTIONS
    PA_DO_DEFAULT = 0
    PA_LOG_AS_MISSED = 1
    PA_MAKE_STICKY = 2
    PA_DO_NOTHING = 3
    PA_DISPLAY_NORMAL = 4
    PA_DISPLAY_URGENT = 5
    PA_FORWARD = 6

End Enum


    ' /* requesters */
Dim mReq() As TRequester
Dim mReqs As Long
Public gRequestId As Long

Public Sub Main()
Dim szArg() As String
Dim l As Long

'    g_ProcessAck "!snarl manage snarlware"
'    Exit Sub

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
        l3OpenLog "%APPDATA%\full phat\snarl\snarl.log"
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
    load_image g_MakePath(App.Path) & "etc\icons\menu.png", bm_Menu
    load_image g_MakePath(App.Path) & "etc\icons\actions.png", bm_Actions
    load_image g_MakePath(App.Path) & "etc\icons\has_actions.png", bm_HasActions
    load_image g_MakePath(App.Path) & "etc\icons\remote_app.png", bm_Remote
    load_image g_MakePath(App.Path) & "etc\icons\secure.png", bm_Secure
    load_image g_MakePath(App.Path) & "etc\icons\is_sticky.png", bm_IsSticky
    load_image g_MakePath(App.Path) & "etc\icons\is_priority.png", bm_Priority

    If Not g_IsValidImage(bm_Close) Then
        Set bm_Close = g_CreateBadge("X")
'        With New mfxView
'            .SizeTo 24, 24
'            .EnableSmoothing False
'            .SetHighColour rgba(255, 0, 0)
'            .FillRect .Bounds
'            .SetHighColour rgba(0, 0, 0)
'            .StrokeRect .Bounds
'            Set bm_Close = .ConvertToBitmap()
'
'        End With
    End If

    ' /* intialize the IP forwarding subsystem */

    g_ForwardInit

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

Dim i As Long

    ' /* display welcome message */

    If (g_ConfigGet("show_msg_on_start") = "1") Or (gDebugMode) Then
        i = g_PrivateNotify(SNARL_CLASS_GENERAL, "Welcome to Snarl!", _
                            "Snarl " & g_Version() & vbCrLf & App.LegalCopyright & vbCrLf & "http://www.fullphat.net" & IIf(gDebugMode, vbCrLf & vbCrLf & "Debug mode enabled", ""), , _
                            g_MakePath(App.Path) & "etc\icons\snarl.png")

        If i Then
            g_QuickAddAction i, "User Guide", "http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=User_Guide"
            g_QuickAddAction i, "Release Notes", "http://sourceforge.net/apps/mediawiki/snarlwin/index.php?title=New_Features#" & Replace$(App.Comments, " ", "_")
        
        End If

    End If

    ' /* get extensions */

    g_Debug "Main(): Starting extensions roster..."
    Set g_ExtnRoster = New TExtensionRoster
    melonLibInit g_ExtnRoster
    melonLibOpen g_ExtnRoster


    Set mUpdateCheck = New TAutoUpdate
    If g_ConfigGet("auto_update") = "1" Then
        g_Debug "Main(): Doing auto-update check..."
        If mUpdateCheck.Check(False, "http://www.fullphat.net/updateinfo/snarl.updateinfo") Then _
            g_Debug "Main(): auto-update check initiated"

    Else
        g_Debug "Main(): Auto-update is disabled"

    End If

    ' /* R2.4.1: check for beta release as well */

    Set mBetaUpdateCheck = New TAutoUpdate

    If g_ConfigGet("auto_beta_update") = "1" Then
        g_Debug "Main(): Doing auto-update (beta) check..."
        If mBetaUpdateCheck.Check(False, "http://snarlwin.svn.sourceforge.net/viewvc/snarlwin/snarl-beta.updateinfo") Then _
            g_Debug "Main(): beta auto-update check initiated"

    Else
        g_Debug "Main(): Auto-update is disabled"

    End If


    ' /* done */

    frmAbout.bReadyToRun

    ' /* */

    g_Debug "Main(): garbage collection"
    If g_IsWinXPOrBetter() Then _
        CoFreeUnusedLibrariesEx 0, 0

    g_Debug "Main(): startup complete"

    With New BMsgLooper
        .Run

    End With

    g_Debug "main(): stopping..."

    If mReqs Then
        For l = mReqs To 1 Step -1
            mReq(l).Quit

        Next l

    End If

    Set mUpdateCheck = Nothing         ' // this will abort the request if it's still running...

Dim uSnarlGlobal As Long
Dim t As Long

    t = GetTickCount()
    g_IsQuitting = True

    ' /* broadcast SNARL_QUIT */

    g_Debug "main(): broadcasting SNARL_QUIT..."
    PostMessage HWND_BROADCAST, g_GlobalMessage(), SNARL_BROADCAST_QUIT, ByVal CLng(App.Major)

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
'        .font_smoothing = E_MELONTYPE
'        .suppress_delay = 2000
'        If g_GetSystemFolder(CSIDL_PERSONAL, sz) Then _
            .last_sound_folder = sz

'        .UserDnD = False
        .use_dropshadow = True

    End With

    ' /* defaults */

    Set mDefaults = New BPackedData
    With mDefaults
        .Add "default_position", CStr(E_START_BOTTOM_RIGHT)
        .Add "show_msg_on_start", "1"
        .Add "run_on_logon", "1"

        ' /* R2.0 (V38.13) */
        .Add "default_style", "corporate/standard"    ' // as "<style>[/<scheme>]

        ' /* R2.0 (V38.32) */
        .Add "sticky_snarls", "0"
        .Add "log_only", "0"
        .Add "default_duration", "10"

        ' /* R2.04 (V38.82) - no longer used */
'        .Add "font_smoothing", CStr(E_MELONTYPE)
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
'        .Add "idle_timeout", "300"          ' // i.e. 5 minutes
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

        ' /* R2.4 DR8 */

        .Add "away_when_locked", "1"
        .Add "away_when_fullscreen", "1"
        .Add "away_when_screensaver", "1"
        .Add "away_mode", "2"               ' // sticky
        .Add "busy_mode", "1"               ' // log missed

        ' /* R2.4 Beta 4 */

        .Add "idle_minutes", "4"            ' // i.e. 5 minutes
        .Add "include_host_name_when_forwarding", "0"

        ' /* R2.4.1 */

        .Add "allow_right_clicks", "0"
        .Add "auto_beta_update", "0"

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
            g_AppRoster.SendToAll SNARL_BROADCAST_LAUNCHED

        ' /* R2.4: broadcast a started message */

        PostMessage HWND_BROADCAST, g_GlobalMessage(), SNARL_BROADCAST_STARTED, ByVal 0&


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

        ' /* R2.4: broadcast a stopped message */
        PostMessage HWND_BROADCAST, g_GlobalMessage(), SNARL_BROADCAST_STOPPED, ByVal 0&

        ' /* close all notifications */
        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.CloseMultiple 0

        ' /* tell the extensions */
        If Not (g_ExtnRoster Is Nothing) Then _
            g_ExtnRoster.SendSnarlState False

        ' /* tell our applications we've stopped */
        If Not (g_AppRoster Is Nothing) Then _
            g_AppRoster.SendToAll SNARL_BROADCAST_QUIT

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

Public Function gfAddClass(ByVal Pid As Long, ByVal Class As String, ByVal Flags As Long, ByVal Description As String) As M_RESULT
Dim pa As TApp

    g_Debug "gfAddClass('" & CStr(Pid) & "' '" & Class & "' #" & g_HexStr(Flags) & ")", LEMON_LEVEL_PROC

    ' /* find the app */

    If Not g_AppRoster.FindByPid(Pid, pa) Then
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


'Public Function gSetUpFontSmoothing(ByRef aView As mfxView, ByVal TextColour As Long, ByVal SmoothingColour As Long) As MFX_DRAWSTRING_FLAGS
'Dim dw As Long
'
'    aView.SetHighColour TextColour
'
'    Select Case gPrefs.font_smoothing
'    Case E_MELONTYPE
'        If SmoothingColour = 0 Then
'            ' /* calculate it (only really works for dark colours at present) */
'            aView.SetLowColour rgba(get_red(TextColour), _
'                                    get_green(TextColour), _
'                                    get_blue(TextColour), _
'                                    (Val(g_ConfigGet("melontype_contrast")) / 100) * 255)
'
'        Else
'            aView.SetLowColour SmoothingColour
'
'        End If
'
'        gSetUpFontSmoothing = MFX_SIMPLE_OUTLINE
'
'    Case E_NONE
'        aView.TextMode = MFX_TEXT_PLAIN
'
'    Case E_ANTIALIAS
'        aView.TextMode = MFX_TEXT_ANTIALIAS
'
'    Case E_CLEARTYPE
'        aView.TextMode = MFX_TEXT_CLEARTYPE
'
'    Case E_WINDOWS_DEFAULT
'        SystemParametersInfo SPI_GETFONTSMOOTHING, 0, dw, 0
'        If dw = 0 Then
'            ' /* none */
'            aView.TextMode = MFX_TEXT_PLAIN
'
'        Else
'            ' /* enabled - but which type? */
'            aView.TextMode = MFX_TEXT_ANTIALIAS     ' // assume antialias...
'            If g_IsWinXPOrBetter() Then
'                dw = 0
'                SystemParametersInfo SPI_GETFONTSMOOTHINGTYPE, 0, dw, 0
'
'                If dw = FE_FONTSMOOTHINGCLEARTYPE Then _
'                    aView.TextMode = MFX_TEXT_CLEARTYPE
'
''                FE_FONTSMOOTHINGSTANDARD and
'
'            End If
'
'        End If
'
'    End Select
'
'End Function

Public Function gfSetAlertDefault(ByVal Pid As Long, ByVal Class As String, ByVal Element As Long, ByVal Value As String) As M_RESULT
Dim pa As TApp
Dim pc As TAlert

    g_Debug "gfSetAlertDefault('" & Pid & "' '" & Class & "' #" & CStr(Element) & " '" & Value & "')", LEMON_LEVEL_PROC

    If (g_AppRoster Is Nothing) Then
        g_Debug "gfSetAlertDefault(): App not registered with Snarl", LEMON_LEVEL_CRITICAL
        gfSetAlertDefault = M_ABORTED
        Exit Function

    End If

    ' /* find the app */

    If Not g_AppRoster.FindByPid(Pid, pa) Then
        g_Debug "gfSetAlertDefault(): App '" & Pid & "' not registered with Snarl", LEMON_LEVEL_CRITICAL
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

'Public Sub g_AddIgnoreLock()
'
'    g_IgnoreLock = g_IgnoreLock + 1
'
'End Sub
'
'Public Sub g_RemIgnoreLock()
'
'    g_IgnoreLock = g_IgnoreLock - 1
'
'End Sub

Public Sub g_WriteToLog(ByVal Title As String, ByVal Text As String)
Dim sz As String
Dim n As Integer

    On Error Resume Next

    If Not g_GetUserFolderPath(sz) Then _
        Exit Sub

    n = FreeFile()
    Open g_MakePath(sz) & "snarl_log.txt" For Append As #n
    If err.Number <> 0 Then _
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

Public Function g_DoSchemePreview2(ByVal Name As String, ByVal Scheme As String, ByVal IsPriority As Boolean, ByVal Percent As Integer) As M_RESULT

    ' /* this handles external requests to Snarl to display a notification in a particular
    '    style and scheme - only the SNARL_PREVIEW_SCHEME message handler calls this */

    If (g_NotificationRoster Is Nothing) Or (g_StyleRoster Is Nothing) Then _
        Exit Function

Dim pStyle As TStyle

    ' /* find the style */

    If Not g_StyleRoster.Find(Name, pStyle) Then _
        Exit Function

    If Scheme = "" Then
        ' /* if no scheme, use "<Default>" */
        Scheme = "<Default>"

    Else
        ' /* otherwise, supplied scheme must exist */
        If pStyle.SchemeIndex(Scheme) = 0 Then _
            Exit Function

    End If

Dim szText As String

    If (Percent > 0) And (Percent <= 100) Then
        ' /* text is actually just a number (i.e. meter-friendly) */
        szText = CStr(Percent)

    Else
        
        szText = "otification using the " & pStyle.Name
        
        If Scheme <> "<Default>" Then
            szText = szText & "/" & Scheme & " style and scheme"

        Else
            szText = szText & " style"

        End If

        If IsPriority Then
            szText = "Priority n" & szText

        Else
            szText = "N" & szText

        End If

    End If


Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        .Title = "Scheme Preview"
        .Text = szText
        .Timeout = -1
        .IconPath = IIf(pStyle.IconPath = "", g_MakePath(App.Path) & "etc\icons\style_preview.png", pStyle.IconPath)
        .StyleName = pStyle.Name
        .SchemeName = LCase$(Scheme)
        .Position = E_START_DEFAULT_POS
        .Priority = IIf(IsPriority, 1, 0)
        Set .ClassObj = New TAlert

    End With

    g_DoSchemePreview2 = (g_NotificationRoster.Add(pInfo, Nothing) <> 0)

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

'Public Function g_StickyNotifications() As Boolean
'
'    g_StickyNotifications = (g_ConfigGet("sticky_snarls") = "1") 'Or (gIsAway)
'
'End Function

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
Dim pti As BTagList
Dim i As Long

    g_Debug "g_ProcessAck(): ACK is '" & Ack & "'", LEMON_LEVEL_INFO

    If g_SafeLeftStr(Ack, 1) = "!" Then
        ' /* bang command */

        arg = Split(g_SafeRightStr(Ack, Len(Ack) - 1), " ")
        argC = UBound(arg)

        Set pti = new_BTagList
        
        ' /* if there are any args, make them into a taglist */

        If argC > 0 Then
            For i = 1 To argC
                pti.Add new_BTagItem(arg(i), "")

            Next i
        End If

        Select Case LCase$(arg(0))

        Case "snarl"
            uProcessSnarl pti

        Case "system"
            uProcessSystem pti

        Case Else
            g_Debug "g_ProcessAck(): unknown command '" & arg(0) & "'"

        End Select

    Else
        ' /* treat as launchable */
        ShellExecute frmAbout.hWnd, vbNullString, Ack, vbNullString, vbNullString, SW_SHOW

    End If

End Sub

Private Sub uProcessSnarl(ByRef Args As BTagList)
Dim pti As BTagItem

    Set pti = Args.TagAt(1)

    If (pti Is Nothing) Then _
        Exit Sub

    Debug.Print pti.Name

Dim i As Long

    Select Case LCase$(pti.Name)

    Case "show_missed_panel"
        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.ShowMissedPanel


    Case "cfg"
        ' /* show our prefs panel targetted on the app in arg(2) */
        Set pti = Args.TagAt(2)
        If Not (pti Is Nothing) Then _
            frmAbout.DoAppConfigBySignature pti.Name


    Case "manage"
        ' /* ask the app in arg(2) to show it's GUI */
        Set pti = Args.TagAt(2)
        If Not (pti Is Nothing) Then
            i = g_AppRoster.IndexOfSig(pti.Name)
            If i Then _
                g_AppRoster.AppAt(i).DoSettings

        End If

    End Select

End Sub

Private Sub uProcessSystem(ByRef Args As BTagList)
Dim pti As BTagItem

    Set pti = Args.TagAt(1)

    If (pti Is Nothing) Then _
        Exit Sub

    Select Case LCase$(pti.Name)

    Case "shutdown_dialog", "shutdown"
        SHShutdownDialog 0

    Case "run_dialog", "run"
        SHRunDialog 0, 0, vbNullString, vbNullString, vbNullString, SHRD_DEFAULT

    Case "lock"
        LockWorkStation

    Case "about"
        ShellAbout 0, vbNullString, vbNullString, 0

    Case "access"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL access.cpl,,0", vbNullString, SW_SHOW
    
    Case "datetime"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL timedate.cpl,,0", vbNullString, SW_SHOW
    
    Case "display"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,3", vbNullString, SW_SHOW
    
    Case "fonts"
        ShellExecute 0, "open", "control.exe", "fonts", vbNullString, SW_SHOW
    
    Case "game"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL joy.cpl,,0", vbNullString, SW_SHOW
    
    Case "software"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL appwiz.cpl,,0", vbNullString, SW_SHOW

    Case "keyboard"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL main.cpl,@1,0", vbNullString, SW_SHOW

    Case "locale"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL intl.cpl,,0", vbNullString, SW_SHOW
    
    Case "mouse"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL main.cpl,,0", vbNullString, SW_SHOW
    
    Case "network"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL ncpa.cpl,,0", vbNullString, SW_SHOW

    Case "power"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL powercfg.cpl,,0", vbNullString, SW_SHOW

    Case "printers"
        ShellExecute 0, "open", "control.exe", "printers", vbNullString, SW_SHOW

    Case "screensaver"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,1", vbNullString, SW_SHOW

    Case "sounds"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL mmsys.cpl,,1", vbNullString, SW_SHOW

    Case "admin"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL sysdm.cpl,,0", vbNullString, SW_SHOW

    Case "telephony"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL telephon.cpl,,0", vbNullString, SW_SHOW

    Case "theme"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,2", vbNullString, SW_SHOW

    Case "users"
        ShellExecute 0, "open", "control.exe", "userpasswords", vbNullString, SW_SHOW

    Case "wallpaper"
        ShellExecute 0, "open", "RUNDLL32.EXE", "SHELL32.DLL,Control_RunDLL desk.cpl,,0", vbNullString, SW_SHOW

    Case "controlpanel"
        ShellExecute 0, "open", "control.exe", vbNullString, vbNullString, SW_SHOW

    Case "trash"
        ShellExecute 0, "open", "::{645FF040-5081-101B-9F08-00AA002F954E}", vbNullString, vbNullString, SW_SHOW

    Case "mycomputer"
        ShellExecute 0, "open", "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", vbNullString, vbNullString, SW_SHOW

    Case "nethood"
        ShellExecute 0, "open", "::{208D2C60-3AEA-1069-A2D7-08002B30309D}", vbNullString, vbNullString, SW_SHOW

    End Select

End Sub

Public Sub gSetLastError(ByVal Error As SNARL_STATUS_CODE)

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

Public Sub g_DoManualUpdateCheck()

    If mUpdateCheck.Check(True, "http://www.fullphat.net/updateinfo/snarl.updateinfo") Then
        g_Debug "g_DoManualUpdateCheck(): check initiated"

    Else
        g_Debug "g_DoManualUpdateCheck(): failed to start", LEMON_LEVEL_CRITICAL

    End If

End Sub

Public Function g_TrimLines(ByVal Text As String) As String

    Text = Replace$(Text, "\n", vbCrLf)

    ' /* pre-set default */

    g_TrimLines = Text

Dim n As Long

    n = g_SafeLong(g_ConfigGet("style.overflow_limit"))
    If (n < 4) Or (n > 12) Then _
        n = 12                  ' // must be at least 5 lines to make it meaningful to
                                ' // take up a line with the ellipsis
                                ' // i.e. line 1
                                ' //      ...
                                ' //      line 3
                                ' //      line 4
                                ' // is pointless as only line 2 is replaced

Dim sz() As String
Dim c As Long
Dim i As Long

    sz() = Split(Text, vbCrLf)
    c = UBound(sz)

    If (c + 1) > n Then
        ' /* <first line> \n <...> \n <remaining lines> */
        g_TrimLines = sz(0) & vbCrLf & "…"
        For i = c - (n - 3) To c
            g_TrimLines = g_TrimLines & vbCrLf & sz(i)

        Next i

    End If

End Function


Public Function g_CreateBadge(ByVal Content As String) As mfxBitmap
Const rx = 6
Dim pr As BRect

    With New mfxView
        .SetFont "Tahoma", 7, True
        .TextMode = MFX_TEXT_ANTIALIAS

        Set pr = new_BRect(0, 0, MAX(.StringWidth(Content), .StringHeight("A")), .StringHeight("A"))
        pr.ExpandBy 8, 8

        .SizeTo pr.Width, pr.Height
        .EnableSmoothing True

        .SetHighColour rgba(0, 0, 0, 190)
        .SetLowColour rgba(0, 0, 0, 140)
        .FillRoundRect pr, rx, rx, MFX_VERT_GRADIENT

        .SetHighColour rgba(255, 255, 255)
        .DrawString Content, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

        .SetHighColour rgba(255, 255, 255)
        .StrokeRoundRect pr.InsetByCopy(1, 1), rx, rx, 2

        .SetHighColour rgba(0, 0, 0, 150)
        .StrokeRoundRect pr, rx, rx, 1
        .StrokeRoundRect pr.InsetByCopy(3, 3), rx, rx, 1

        Set g_CreateBadge = .ConvertToBitmap()

    End With

End Function

Public Sub g_KludgeNotificationInfo(ByRef nInfo As T_NOTIFICATION_INFO)

    ' /* translates current T_NOTIFICATION_INFO content into a BPackedString
    '    and then stores that back into T_NOTIFICATION_INFO->OriginalContent
    '    this is so V42 styles can access the entire request */

    If (nInfo.ClassObj Is Nothing) Then
        g_Debug "g_KludgeNotificationInfo(): missing ClassObj", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim ppd As BPackedData

    Set ppd = New BPackedData

    With nInfo
        .Title = Replace$(.Title, "\n", vbCrLf)
        .Text = Replace$(.Text, "\n", vbCrLf)

        ppd.Add "id", .ClassObj.Name
        ppd.Add "title", .Title
        ppd.Add "text", .Text
        ppd.Add "timeout", CStr(.Timeout)
        ppd.Add "icon", .IconPath
        ppd.Add "priority", CStr(.Priority)
        ppd.Add "callback", .DefaultAck
        ppd.Add "value", .Value

'        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
            .Add "merge", "1"

    End With

Dim ppx As BPackedData
Dim szn As String
Dim szv As String

    ' /* add in all other custom content */

    Set ppx = New BPackedData
    With ppx
        If .SetTo(nInfo.OriginalContent) Then
            .Rewind
            Do While .GetNextItem(szn, szv)
                If Not ppd.Exists(LCase$(szn)) Then _
                    ppd.Add szn, szv

            Loop
        End If
    End With

    nInfo.OriginalContent = ppd.AsString()












'
'
'    ' /* generates a packed string from the provided T_NOTIFICATION_INFO
'    '    struct and assigns the string to the OriginalContent element,
'    '    which is required by V42 styles */
'
'    With nInfo
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'        .OriginalContent = "id::" & .ClassObj.Name & _
'                           "#?title::" & .Title & _
'                           "#?text::" & .Text & _
'                           "#?timeout::" & CStr(.Timeout) & _
'                           "#?icon::" & .IconPath & _
'                           "#?priority::" & CStr(.Priority) & _
'                           "#?ack::" & .DefaultAck & _
'                           "#?value::" & .Value
'
''        If (.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
''            .OriginalContent = .OriginalContent & "#?merge::1"
'
'    End With

End Sub

Public Function g_TranslateIconPath(ByVal Icon As String, ByVal StylePath As String) As String
Dim pbm As mfxBitmap
Dim pIcon As BIcon
Dim sz As String
Dim dw As Long
Dim i As Long


    If g_SafeLeftStr(Icon, 1) = "!" Then
        ' /* convert the icon into it's corresponding file */
        Icon = g_SafeRightStr(Icon, Len(Icon) - 1) & ".png"

        ' /* default location */

        g_TranslateIconPath = g_MakePath(App.Path) & "etc\default_theme\icons\" & Icon

        ' /* here, 'StylePath' sould be either empty or contain the full path to the
        '    current style.  If it's the former we use the built-in icon set; if
        '    it's the latter we try to get the icon from the style */

        If (g_ConfigGet("use_style_icons") = "1") And (StylePath <> "") Then
            ' /* try to get it from the style... */
            sz = g_MakePath(StylePath) & "icons\" & Icon
            If g_Exists(sz) Then _
                g_TranslateIconPath = sz

        End If

    ElseIf g_SafeLeftStr(Icon, 1) = "%" Then
        ' /* whatever is after the % should be a valid HICON */
        dw = Val(g_SafeRightStr(Icon, Len(Icon) - 1))
        If dw Then
            Set pbm = uGetIcon(dw)
            If Not (pbm Is Nothing) Then
                g_TranslateIconPath = g_GetSafeTempIconPath()
                pbm.Save g_TranslateIconPath, "image/png"

            End If

        Else
            g_Debug "TNotificationRoster.g_TranslateIconPath(): bad HICON '" & Icon & "'", LEMON_LEVEL_WARNING

        End If

    ElseIf LCase$(g_GetExtension(Icon)) = "ico" Then
        g_TranslateIconPath = uLoadICO(Icon)

    Else

        g_TranslateIconPath = Icon

        ' /* windows icon? */

        i = InStr(Icon, ",")
        If i <> 0 Then
            dw = Val(g_SafeRightStr(Icon, Len(Icon) - i))
            Icon = g_SafeLeftStr(Icon, i - 1)
            g_TranslateIconPath = uGetBestIcon(Icon, dw)

        End If
    End If

End Function

Private Function uLoadICO(ByVal IconPath As String) As String
Dim pbm As mfxBitmap
Dim pIcon As BIcon

    On Error Resume Next

    With New BIconContent
        If Not .LoadFromICO(IconPath) Then _
            Exit Function

        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pIcon) Then
            Set pbm = create_bitmap_from_image(pIcon.Render)
            uLoadICO = g_GetSafeTempIconPath()
            pbm.Save uLoadICO, "image/png"

        End If

    End With

End Function

Private Function uGetBestIcon(ByVal IconPath As String, ByVal Index As Long) As String
Dim pbm As mfxBitmap
Dim pIcon As BIcon

    On Error Resume Next

    With New BIconContent
        If Not .LoadFromResource(IconPath, Index) Then _
            Exit Function

        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pIcon) Then
            Set pbm = create_bitmap_from_image(pIcon.Render)
            uGetBestIcon = g_GetSafeTempIconPath()
            pbm.Save uGetBestIcon, "image/png"

        End If

    End With

End Function

Private Function uGetIcon(ByVal hIcon As Long) As mfxBitmap

    On Error Resume Next

Dim pi As BIcon

    err.Clear
    Set pi = New BIcon
    If err.Number <> 0 Then
        g_Debug "TNofiticationRoster.uGetIcon(): no icon.resource", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If Not pi.SetFromHICON(hIcon) Then
        g_Debug "TNofiticationRoster.uGetIcon(): bad icon handle", LEMON_LEVEL_CRITICAL
        Exit Function

    End If
        
    Set uGetIcon = create_bitmap_from_image(pi.Render())

End Function

Public Function g_ShowRequest(ByVal lPid As Long, ByRef Data As BPackedData) As Long
Dim pReq As TRequester

    Set pReq = New TRequester
    g_ShowRequest = pReq.Go(lPid, Data)

    If g_ShowRequest <> 0 Then

        ' /* disable existing */
        If mReqs > 0 Then _
            mReq(mReqs).SetEnabled False

        mReqs = mReqs + 1
        ReDim Preserve mReq(mReqs)
        Set mReq(mReqs) = pReq

    End If

End Function

Public Sub g_PopRequest2()

    If mReqs < 1 Then _
        Exit Sub

    mReqs = mReqs - 1
    ReDim Preserve mReq(mReqs)
    If mReqs > 0 Then _
        mReq(mReqs).SetEnabled True

End Sub

Private Function uCreatePacked(ByVal ClassId As String, ByVal Title As String, ByVal Text As String, Optional ByVal Timeout As Long = -1, Optional ByVal Icon As String, Optional ByVal Priority As Long = 0, Optional ByVal Ack As String, Optional ByVal Flags As SNARL41_NOTIFICATION_FLAGS, Optional ByVal Password As String) As BPackedData

    ' /* translate notification arguments into packed data
    '    currently this is only used by g_PrivateNotify()
    '    but it's flexible enough to be used elsewhere */

    Set uCreatePacked = New BPackedData
    With uCreatePacked
        If ClassId <> "" Then _
            .Add "class", ClassId

        If Title <> "" Then _
            .Add "title", Title

        If Text <> "" Then _
            .Add "text", Text

        .Add "timeout", Timeout

        If Icon <> "" Then _
            .Add "icon", Icon

        .Add "priority", CStr(Priority)
        
        If Ack <> "" Then _
            .Add "ack", Ack

        .Add "flags", Hex$(Flags)               ' // flags are sent as a hex value

        If Password <> "" Then _
            .Add "password", Password

    End With

End Function

Public Function g_QuickAddAction(ByVal Token As Long, ByVal Label As String, ByVal Command As String) As Long

    g_QuickAddAction = g_NotificationRoster.AddAction(Token, g_newBPackedData("label::" & Label & "#?cmd::" & Command))

End Function

Public Function g_QuickLastError() As Long

    g_QuickLastError = GetProp(ghWndMain, "last_error")

End Function

'Public Function g_QuickAddClass(ByVal AppToken As Long, ByVal Id As String, ByVal Name As String, Optional ByVal Enabled As Boolean, Optional ByVal Password As String) As Long
'Dim pp As BPackedData
'
'    Set pp = New BPackedData
'    With pp
'        .Add "id", Id
'        .Add "name", Name
'        .Add "enabled", IIf(Enabled, "1", "0")
'        If Password <> "" Then _
'            .Add "password", Password
'
'    End With
'
'    g_QuickAddClass = g_DoAction("addclass", AppToken, pp)
'
'End Function

Public Function g_PrivateNotify(ByVal ClassId As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Timeout As Long = -1, Optional ByVal Icon As String, Optional ByVal Priority As Long = 0, Optional ByVal Ack As String, Optional ByVal Flags As SNARL41_NOTIFICATION_FLAGS, Optional ByVal IntFlags As E_NOTIFICATION_FLAGS) As Long

    ' /* internal notification generator
    '
    '    uses g_DoNotify() to display a Snarl-generated notification
    '    without going via the Win32 messaging system */

    g_PrivateNotify = g_DoNotify(gSnarlToken, _
                                 uCreatePacked(ClassId, Title, Text, Timeout, Icon, Priority, Ack, Flags, gSnarlPassword), _
                                 , _
                                 IntFlags Or App.Major)

End Function

Public Function g_DoNotify(ByVal AppToken As Long, ByRef pData As BPackedData, Optional ByRef ReplySocket As CSocket, Optional ByVal IntFlags As E_NOTIFICATION_FLAGS, Optional ByVal RemoteHost As String) As Long

    ' /* master notification generator
    '
    '    all roads should lead here - there should be no use of sn41EZNotify() or any other
    '    Win32 API function.  Similarly, there should be no by-passing of this function,
    '    except in very specific circumstances (style previews, for example) */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
        g_Debug "g_DoNotify(): app and/or notification roster missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    If (pData Is Nothing) Then
        g_Debug "g_DoNotify(): arg missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* look for the new "replace-uid" and "update-uid" and "merge-uid" args:
    '    "replace" will remove the notification with the specified uid if it's
    '    still on-screen; "update-uid" will cause the specified notification
    '    to be updated with this content and "merge-uid" will cause the
    '    provided content to be merged with the existing notification */

Dim pn As TNotification

    If pData.Exists("replace-uid") Then
        ' /* if the specified uid (NOT token) exists, remove it */
        g_NotificationRoster.Hide 0, pData.ValueOf("replace-uid"), pData.ValueOf("app-sig"), pData.ValueOf("password")

    ElseIf pData.Exists("update-uid") Then
        ' /* if the specified uid (NOT token) exists, update with this content otherwise create a new notification */
        g_Debug "g_DoNotify(): looking for (update-)uid: " & pData.ValueOf("update-uid") & "..."

        If g_NotificationRoster.Find(0, pData.ValueOf("update-uid"), pData.ValueOf("app-sig"), pData.ValueOf("password"), pn) Then
            pn.UpdateOrMerge pData, False
            g_DoNotify = pn.Info.Token
            Exit Function

        End If

    ElseIf pData.Exists("merge-uid") Then
        ' /* if the specified uid (NOT token) exists, merge this content with that one, otherwise create a new notificaton */
        g_Debug "g_DoNotify(): looking for (merge-)uid: " & pData.ValueOf("merge-uid") & "..."

        If g_NotificationRoster.Find(0, pData.ValueOf("merge-uid"), pData.ValueOf("app-sig"), pData.ValueOf("password"), pn) Then
            pn.UpdateOrMerge pData, True
            g_DoNotify = pn.Info.Token
            Exit Function

        End If

    End If

    ' /* this still takes effect even if other options above have been used */

    If pData.Exists("uid") Then
        ' /* if the specified uid (NOT token) exists, update this content with that one, otherwise create a new notificaton */
        g_Debug "g_DoNotify(): looking for uid: " & pData.ValueOf("uid") & "..."

        If g_NotificationRoster.Find(0, pData.ValueOf("uid"), pData.ValueOf("app-sig"), pData.ValueOf("password"), pn) Then
            pn.UpdateOrMerge pData, False
            g_DoNotify = pn.Info.Token
            Exit Function

        Else
            g_Debug "g_DoNotify(): uid " & pData.ValueOf("uid") & " not found"

        End If

    End If

Dim szClass As String
Dim pApp As TApp

    ' /* R2.4 DR7: if "app-sig" argument is specified then look for the app by signature */

    If pData.ValueOf("app-sig") <> "" Then
        If g_AppRoster.FindBySignature(pData.ValueOf("app-sig"), pApp, pData.ValueOf("password")) Then
            ' /* R2.4.1 - support for "class" keyword */
            If pData.Exists("class") Then
                szClass = pData.ValueOf("class")

            Else
                szClass = pData.ValueOf("id")

            End If

        Else
            ' /* not found / auth failure (lasterror will have been set) */
            Exit Function

        End If

    Else
        ' /* special case: if the app token is 0 we use ourself as the sending app  */
        If AppToken = 0 Then
            If g_AppRoster.FindByToken(gSnarlToken, pApp, gSnarlPassword) Then
                ' /* if we're using the Snarl app, we need the anonymous class */
                g_Debug "g_DoNotify(): using Snarl anonymous class"
                szClass = IIf((IntFlags And NF_REMOTE) = 0, SNARL_CLASS_ANON, SNARL_CLASS_ANON_NET)

            Else
               ' /* Snarl's registration not found */
                g_Debug "g_DoNotify(): Snarl internal app not in roster", LEMON_LEVEL_CRITICAL
                gSetLastError SNARL_ERROR_SYSTEM
                Exit Function

            End If

        ElseIf g_AppRoster.FindByToken(AppToken, pApp, pData.ValueOf("password")) Then
            ' /* R2.4.1 - support for "class" keyword */
            If pData.Exists("class") Then
                szClass = pData.ValueOf("class")

            Else
                szClass = pData.ValueOf("id")

            End If

        Else
            ' /* not found / auth failure (lasterror will have been set) */
            Exit Function

        End If

    End If

    ' /* include the remote sender (if there is one) - small kludge here for
    '    Growl/UDP which won't have a reply socket */

Dim szRemoteHost As String

    If Not (ReplySocket Is Nothing) Then
        If ReplySocket.RemoteHost <> "" Then
            szRemoteHost = ReplySocket.RemoteHost & " (" & ReplySocket.RemoteHostIP & ")"

        Else
            szRemoteHost = ReplySocket.RemoteHostIP

        End If

    Else
        szRemoteHost = RemoteHost

    End If

    ' /* R2.4 DR7 - merging is now controlled via an internal flag */

    If pData.ValueOf("merge") = "1" Then _
        IntFlags = IntFlags Or NF_MERGE

    ' /* now we have the app object and we know the class, we can pass it over */

    g_DoNotify = pApp.Show41(szClass, pData, ReplySocket, IntFlags, szRemoteHost)

End Function

Public Function g_DoAction(ByVal action As String, ByVal Token As Long, ByRef Args As BPackedData, Optional ByVal InternalFlags As E_NOTIFICATION_FLAGS, Optional ByRef ReplySocket As CSocket, Optional ByVal SenderPID As Long) As Long

    ' /* this is the central hub for all incoming requests, be they from SNP, Growl/UDP
    '    or Win32.  "Token" here can be either the app token or the notification token;
    '    the action determines which one */

    ' /* Return zero on error (and set lasterror), -1 or a +ve value on success */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Or (Args Is Nothing) Then
        ' /* pretty much all of these require either or both rosters to be
        '    available, so let's bail out now if something's really wrong */
        gSetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

Dim pApp As TApp

    ' /* assume all okay... */

    gSetLastError 0

    Select Case action

    Case "addaction"
        g_DoAction = g_NotificationRoster.AddAction(Token, Args)

    Case "addclass"
        g_DoAction = uAddClass(Token, Args)

    Case "clearclasses", "killclasses"
        g_DoAction = uRemClass(Token, Args, True)

    Case "clearactions"
        g_DoAction = g_NotificationRoster.ClearActions(Token, Args)

    Case "hello"
        ' /* reply our major version number */
        ' /* To-do: reply with an error message if Snarl isn't
        '    accepting requests, or DND mode enabled? */
        g_DoAction = App.Major

    Case "hide"
        g_DoAction = CLng(g_NotificationRoster.Hide(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password")))

    Case "isvisible"
        g_DoAction = CLng(g_NotificationRoster.IsVisible(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password")))

    Case "notify"
        g_DoAction = g_DoNotify(Token, Args, ReplySocket, InternalFlags)

    Case "reg", "register"
        g_DoAction = g_AppRoster.Add41(Args, (InternalFlags And NF_REMOTE), SenderPID)

    Case "remclass"
        g_DoAction = uRemClass(Token, Args)

    Case "test"
        ' /* only available when Snarl is running in debug mode */
        If gDebugMode Then
            g_PrivateNotify "", _
                            IIf(Args.ValueOf("alpha") = "", "Snarl", Args.ValueOf("alpha")), _
                            IIf(Args.ValueOf("beta") = "", "Test Message", Args.ValueOf("beta"))
            g_DoAction = -1

        Else
            gSetLastError SNARL_ERROR_UNKNOWN_COMMAND
            g_DoAction = 0

        End If

    Case "unreg", "unregister"
        ' /* R2.4 DR7: can unregister using signature/password combo */
        If Args.Exists("app-sig") Then
            g_DoAction = g_AppRoster.UnregisterBySig(Args.ValueOf("app-sig"), Args.ValueOf("password"))

        Else
            g_DoAction = g_AppRoster.Unregister(Token, Args.ValueOf("password"))

        End If

    Case "update"
        g_DoAction = g_NotificationRoster.Update(Token, Args)

    Case "updateapp", "update_app"
        g_DoAction = g_AppRoster.Update(Token, Args)

    Case "version"
        g_DoAction = GetProp(ghWndMain, "_version")


    ' /* V42 only (no corresponding V41 command ID) */


    Case "request"
        ' /* PRIVATE: for internal use only under V42 */
        g_DoAction = g_ShowRequest(Token, Args)

    Case "wasmissed"
        g_DoAction = g_NotificationRoster.WasMissed(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"))

'    Case "merge"
'        ' /* specify an existing token or uid/app-sig pair that identifies the notification
'        '    to merge with.  Creates a new notification (same uid, different token) if
'        '    specified notification doesn't exist */
'        g_DoAction = g_NotificationRoster.Merge(Token, Args)

    Case "setmode"
        If Args.Exists("busy") Then _
            g_DoAction = uSetBusy(Token, Args)

    Case Else
        gSetLastError SNARL_ERROR_UNKNOWN_COMMAND
        g_DoAction = 0

    End Select

End Function

Public Function g_newBPackedData(ByVal Content As String) As BPackedData

    Set g_newBPackedData = New BPackedData
    g_newBPackedData.SetTo Content

End Function

'Public Function g_DoMerge(ByRef Args As BPackedData) As Boolean
'
'    ' /* token may be null, in which case we must have an app-sig/uid pair */
'
'Dim i As Long
'
'    If (Args.Exists("app-sig")) And (Args.Exists("uid")) Then
'        i = g_NotificationRoster.UIDToToken(Args.ValueOf("app-sig"), Args.ValueOf("uid"), Args.ValueOf("password"))
'
'    Else
'        i = g_SafeLong(Args.ValueOf("token"))
'
'    End If
'
'
'    If i Then
'        ' /* merge with this one */
''        mItem(i).Window.MergeWith
'
'    Else
'        ' /* not found? create new... */
'
'
'
'    End If
'
'    gSetLastError SNARL_ERROR_UNKNOWN_COMMAND
'    Exit Function
'
'
''    If i Then
''        g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' found"
''        Update = mItem(i).Window.Update(Args)
''
''    Else
''        i = uFindInMissedList(Token)
''        If i Then
''            g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' is in missed list"
''
''        Else
''            Update = False
''
''        End If
''
''    End If
'
'End Function





'Public Function g_DoNotify(ByVal Token As Long, ByRef pData As BPackedData, Optional ByRef ReplySocket As CSocket) As Long
'
'    ' /* sanity checking */
'
'    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
'        g_Debug "g_DoNotify(): app and/or notification roster missing", LEMON_LEVEL_CRITICAL
'        gSetLastError SNARL_ERROR_SYSTEM
'        Exit Function
'
'    End If
'
'    If (pData Is Nothing) Then
'        g_Debug "g_DoNotify(): arg missing", LEMON_LEVEL_CRITICAL
'        gSetLastError SNARL_ERROR_ARG_MISSING
'        Exit Function
'
'    End If
'
'Dim szClass As String
'Dim pApp As TApp
'
'    ' /* special case: if the app token is 0 we use ourself as the sending app  */
'
'    If Token = 0 Then
'        If g_AppRoster.FindByToken(gSnarlToken, pApp, "") Then                  ' // <--- Snarl should be password protected?
'            ' /* if we're using the Snarl app, we need the anonymous class */
'            g_Debug "g_DoNotify(): using Snarl anonymous class"
'            szClass = SNARL_CLASS_ANON
'
'        Else
'            ' /* Snarl's registration not found */
'            g_Debug "g_DoNotify(): Snarl internal app not in roster", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_SYSTEM
'            Exit Function
'
'        End If
'
'    ElseIf g_AppRoster.FindByToken(Token, pApp, pData.ValueOf("password")) Then
'        szClass = pData.ValueOf("id")
'
'    Else
'        ' /* not found / auth failure (lasterror will have been set) */
'        Exit Function
'
'    End If
'
'
''Dim pInfo As T_NOTIFICATION_INFO
''Dim i As Long
''
''    With pInfo
''        .hWndReply = Val(pData.ValueOf("hwnd"))
''        .uReplyMsg = Val(pData.ValueOf("umsg"))
''        .IconPath = pData.ValueOf("icon")
''        .Text = pData.ValueOf("text")
''
''        If pData.Exists("timeout") Then
''            .Timeout = Val(pData.ValueOf("timeout"))
''
''        Else
''            .Timeout = -1
''
''        End If
''
''        .Title = pData.ValueOf("title")
''        .Priority = Val(pData.ValueOf("priority"))
''        .DefaultAck = pData.ValueOf("ack")
''        .Value = pData.ValueOf("value")
''
''        If pData.Exists("flags") Then
''            i = Val("&H" & pData.ValueOf("flags"))
''            .Flags = (i And &HFFFF&)                        ' // only keep user flags
''
''        End If
''
''        ' /* these can't be set by external applications - it's a bit klunky at
''        '    present but the notify command handling code in TMainWindow will
''        '    bounce any requests with these tags in them */
''
''        If pData.Exists("remote") Then _
''            .Flags = .Flags Or SNARL42_NOTIFICATION_REMOTE
''
''        If pData.Exists("secure") Then _
''            .Flags = .Flags Or SNARL42_NOTIFICATION_SECURE
''
''
''        .OriginalContent = pData.AsString()
''
''        Set .Socket = ReplySocket
''
''    End With
'
'    g_DoNotify = pApp.Show41(szClass, pData, ReplySocket)
'
'End Function

'Public Function g_DoUpdate41(ByVal Token As Long, ByRef pData As BPackedData) As Long
'
'    ' /* return -1 on success, 0 on failure */
'
'    gSetLastError SNARL_ERROR_SYSTEM
'    If (g_NotificationRoster Is Nothing) Then _
'        Exit Function
'
'Dim pInfo As notification_info
'
'    If pData.Exists("title") Then
'        pInfo.Title = pData.ValueOf("title")
'
'    Else
'        pInfo.Title = Chr$(255)
'
'    End If
'
'    If pData.Exists("text") Then
'        pInfo.Text = pData.ValueOf("text")
'
'    Else
'        pInfo.Text = Chr$(255)
'
'    End If
'
'    If pData.Exists("icon") Then
'        pInfo.Icon = pData.ValueOf("icon")
'
'    Else
'        pInfo.Icon = Chr$(255)
'
'    End If
'
'Dim hr As M_RESULT
'
'    ' /* call the pre-V41 stuff here - LastError will be set */
'
'    hr = g_NotificationRoster.Update(Token, pInfo.Title, pInfo.Text, pInfo.Icon, pData.AsString)
'    If hr = M_OK Then
'        ' /* success, was timeout specified? */
'
'        If pData.Exists("timeout") Then _
'            g_NotificationRoster.SetAttribute Token, SNARL_ATTRIBUTE_TIMEOUT, pData.ValueOf("timeout")
'
'        g_DoUpdate41 = -1
'
'    Else
'        g_DoUpdate41 = 0
'
'    End If
'
'End Function

Public Function g_GlobalMessage() As Long

    g_GlobalMessage = RegisterWindowMessage(SNARL_XXX_GLOBAL_MSG)

End Function

Private Function uAddClass(ByVal Token As Long, ByRef Args As BPackedData) As Long
Dim pApp As TApp

    If Token Then
        ' /* FindByToken() will set lasterror for us */
        If g_AppRoster.FindByToken(Token, pApp, Args.ValueOf("password")) Then _
            uAddClass = pApp.AddClass(Args)

    Else
        ' /* FindBySignature() will set lasterror for us */
        If g_AppRoster.FindBySignature(Args.ValueOf("app-sig"), pApp, Args.ValueOf("password")) Then _
            uAddClass = pApp.AddClass(Args)

    End If

End Function

Private Function uRemClass(ByVal Token As Long, ByRef Args As BPackedData, Optional ByVal RemoveAll As Boolean = False) As Long
Dim pApp As TApp

    If Token Then
        ' /* FindByToken() will set lasterror for us */
        If g_AppRoster.FindByToken(Token, pApp, Args.ValueOf("password")) Then _
            uRemClass = pApp.RemClass(Args, RemoveAll)

    Else
        ' /* FindBySignature() will set lasterror for us */
        If g_AppRoster.FindBySignature(Args.ValueOf("app-sig"), pApp, Args.ValueOf("password")) Then _
            uRemClass = pApp.RemClass(Args, RemoveAll)

    End If

End Function

Public Function taglist_as_string(ByRef aList As BTagList) As String

    If (aList Is Nothing) Then _
        Exit Function

Dim pt As BTagItem
Dim sz As String

    With aList
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            sz = sz & pt.Name & "::" & pt.Value & "#?"

        Loop

    End With

    taglist_as_string = g_SafeLeftStr(sz, Len(sz) - 2)

End Function

Public Sub g_SetPresence(ByVal Flags As SP_PRESENCE_FLAGS)
Dim fWasAway As Boolean

    fWasAway = g_IsAway()
    mPresFlags = mPresFlags Or Flags

    ' /* if we've transitioned to Away, notify registered apps */

    If (Not fWasAway) And ((mPresFlags And SP_AWAY_MASK) <> 0) Then _
        g_AppRoster.SendToAll SNARL_BROADCAST_USER_AWAY
        ' /* TO-DO: change the tray icon? */

End Sub

Public Sub g_ClearPresence(ByVal Flags As SP_PRESENCE_FLAGS)
Dim f As Boolean

    f = (mPresFlags <> 0)
    mPresFlags = mPresFlags And (Not Flags)

    ' /* if we've transitioned from Away, notify registered apps */

    If (f) And (mPresFlags = 0) Then
        g_AppRoster.SendToAll SNARL_BROADCAST_USER_BACK
        g_NotificationRoster.CheckMissed
        ' /* TO-DO: change the tray icon? */

    End If

End Sub

Public Function g_IsAway() As Boolean

    g_IsAway = ((mPresFlags And SP_AWAY_MASK) <> 0)

End Function

Public Function g_IsDND() As Boolean

    g_IsDND = ((mPresFlags And SP_DND_MASK) <> 0)

End Function

Public Function g_IsPresence(ByVal Flags As SP_PRESENCE_FLAGS) As Boolean

    g_IsPresence = ((mPresFlags And Flags) <> 0)

End Function

Public Function g_GetPresence() As Long

    g_GetPresence = mPresFlags

End Function

Public Function g_GetBase64Icon(ByVal Data As String) As String
Dim sz As String
Dim bErr As Boolean

    On Error Resume Next

    sz = Decode64(Replace$(Data, "%", "="), bErr)
    If (sz = "") Or (bErr) Then
        g_Debug "TNotificationRoster.g_GetBase64Icon(): failed to decode Base64", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* get a suitably unique path */

    g_GetBase64Icon = g_GetSafeTempIconPath()

Dim i As Integer

    ' /* write the data out */

    i = FreeFile()

    err.Clear
    Open g_GetBase64Icon For Binary Access Write As #i
    If err.Number = 0 Then
        Put #i, , sz
        Close #i

    End If

    g_Debug "TNotificationRoster.g_GetBase64Icon(): writing icon to '" & g_GetBase64Icon & "'"

End Function

Private Function uSetBusy(ByVal Token As Long, ByRef Args As BPackedData) As Long
Dim pApp As TApp

    If Token Then
        ' /* FindByToken() will set lasterror for us */
        If Not g_AppRoster.FindByToken(Token, pApp, Args.ValueOf("password")) Then _
            Exit Function

    Else
        ' /* FindBySignature() will set lasterror for us */
        If Not g_AppRoster.FindBySignature(Args.ValueOf("app-sig"), pApp, Args.ValueOf("password")) Then _
            Exit Function

    End If

    ' /* no app? gah... */

    If (pApp Is Nothing) Then
        g_Debug "uSetBusy(): no returned app object", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    ' /* TO-DO: allow for user to prevent the app from changing busy mode */

    Select Case g_SafeLong(Args.ValueOf("busy"))
    Case 0
        ' /* reduce count */
        uSetBusyCount False

    Case 1
        ' /* increase count */
        uSetBusyCount True

    Case Else
        ' /* error */
        gSetLastError SNARL_ERROR_INVALID_ARG

    End Select

End Function

Private Sub uSetBusyCount(ByVal Increment As Boolean)
Static nBusy As Long

    If Increment Then
        g_Debug "uSetBusyCount(): increasing..."
        nBusy = nBusy + 1
        If nBusy = 1 Then _
            g_SetPresence SP_DND_EXTERNAL

    Else
        g_Debug "uSetBusyCount(): decreasing..."
        nBusy = nBusy - 1
        If nBusy = 0 Then _
            g_ClearPresence SP_DND_EXTERNAL

    End If

End Sub

Public Function g_When(ByVal Timestamp As Date) As String
Dim i As Long

    On Error GoTo fail

    ' /* default response is "<date> at <time>" */

    g_When = Format$(Timestamp, "d mmm yyyy") & " at " & Format$(Timestamp, "ttttt")

    ' /* if more than a day ago, default is enough */

    If DateDiff("d", Timestamp, Now) > 0 Then _
        Exit Function

    ' /* if an hour or more ago, use hours */

    i = DateDiff("h", Timestamp, Now)
    If i > 0 Then
        g_When = CStr(i) & " hour" & IIf(i = 1, "", "s") & " ago"
        Exit Function

    End If

    ' /* if a minute or more ago. use minutes */

    i = DateDiff("n", Timestamp, Now)
    If i > 0 Then
        g_When = CStr(i) & " min" & IIf(i = 1, "", "s") & " ago"

    Else
        g_When = "Just now"

    End If

fail:

End Function

