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
    '/  Created:        2-Feb-2010
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

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

Private Const LIB_DATE = "10-Jun-2011"
Private Const LIB_RELEASE = "3.05"
Private Const EXT_NAME = "SysInfo"
Private Const EXT_COMMENT = "Tracks events specific to this computer"

Private Const CLASS_LOW_MEM = "lmem"
Private Const CLASS_NO_MEM = "nmem"
Private Const CLASS_HIGH_CPU = "hcpu"
Private Const CLASS_LOW_DISK = "ldsk"
Private Const CLASS_NIC_CONN = "nicc"
Private Const CLASS_NIC_DISC = "nicd"
Private Const CLASS_DUN_CONN = "dunc"
Private Const CLASS_DUN_DISC = "dund"
Private Const CLASS_VOL_MNT = "volm"
Private Const CLASS_VOL_UMNT = "volu"
Private Const CLASS_POWER_CONNECT = "pwrc"
Private Const CLASS_POWER_DISCONNECT = "pwrd"
Private Const CLASS_POWER_LOW = "pwrl"
Private Const CLASS_LOCK_KEY = "keyb"
Private Const CLASS_IP_CHANGE = "ipch"
Private Const CLASS_IE_DOWNLOAD = "iedl"
Private Const CLASS_CLIPBOARD = "clip"

Dim mLastErr As String
Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)

Dim mToken As Long
Dim mConfig As CConfFile

Dim mAdapterList As BTagList

Dim WithEvents theNICTimer As BTimer
Attribute theNICTimer.VB_VarHelpID = -1
Dim WithEvents theVolTimer As BTimer
Attribute theVolTimer.VB_VarHelpID = -1

Dim WithEvents theNetRoster As TNetRoster2
Attribute theNetRoster.VB_VarHelpID = -1
Dim WithEvents theDUNRoster As TDUNRoster
Attribute theDUNRoster.VB_VarHelpID = -1
Dim WithEvents theVolRoster As TSimpleVolRoster
Attribute theVolRoster.VB_VarHelpID = -1

Dim WithEvents theHandler As TWindow
Attribute theHandler.VB_VarHelpID = -1

Dim mPanel As BPrefsPanel

    ' /* notification tokens */

Dim mPowerId As Long
Dim mLowPowerId As Long
Dim mLowRamId As Long
Dim mCapsId As Long
Dim mNumId As Long
Dim mScrlId As Long

Implements KPrefsPanel
Implements KPrefsPage

'Implements MVersionInfo             ' // optional: allows the SDOS version command to query the object
Implements ISnarlExtension

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage
Dim pCtl As BControl
Dim pp As MPoint
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetWindow 1
            .SetTitle "SysInfo Preferences"
            .SetWidth 128

            ' /* CPU */

            Set pPage = new_BPrefsPage("CPU", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "cpu.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("fancytoggle2", "cpu.enabled", "Monitor system load?", , mConfig.ValueOf("cpu.enabled"))
                .Add new_BPrefsControl("label", "", "Warn when system load is at least:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 100&
                pm.Add "freq", 10&
                pm.Add "label_size", 64&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "cpu.high_mark", "", "", mConfig.ValueOf("cpu.high_mark"), pm, (mConfig.ValueOf("cpu.enabled") = "1"))

                .Add new_BPrefsControl("label", "", "For at least:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 60&
                pm.Add "freq", 5&
                pm.Add "label_size", 64&
                pm.Add "suffix", " seconds"
                .Add new_BPrefsControl("fancyslider", "cpu.duration", "", "", mConfig.ValueOf("cpu.duration"), pm, (mConfig.ValueOf("cpu.enabled") = "1"))
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "cpu.test", "Test Settings", , , , (mConfig.ValueOf("cpu.enabled") = "1"))

            End With

            .AddPage pPage


            ' /* RAM */

            Set pPage = new_BPrefsPage("Memory", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "memory.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("label", "", "Warn when remaining physical RAM reaches:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "ram.low_mark", "", "", mConfig.ValueOf("ram.low_mark"), pm)

'                .Add new_BPrefsControl("fancytoggle2", "ram.show_title", "Display title?", , mConfig.ValueOf("ram.show_title"))
                .Add new_BPrefsControl("fancytoggle2", "ram.show_text", "Display hint text?", , mConfig.ValueOf("ram.show_text"))
                .Add new_BPrefsControl("label", "", "Only applies to the Physical RAM Exhausted notification.")

                .Add new_BPrefsControl("fancytoggle2", "ram.use_meter", "Display meter?", , mConfig.ValueOf("ram.use_meter"))
                .Add new_BPrefsControl("label", "", "Only applies to the Low Physical RAM notification.")
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "ram.test", "Test Settings")

            End With

            .AddPage pPage


            ' /* Disk Space */

            Set pPage = new_BPrefsPage("Disk", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "diskfree.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("fancytoggle2", "disk.enabled", "Monitor disk space?", , mConfig.ValueOf("disk.enabled"))
                .Add new_BPrefsControl("label", "", "Warn when disk free space drops to:")

                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "disk.low_mark", "", "", mConfig.ValueOf("disk.low_mark"), pm, (mConfig.ValueOf("disk.enabled") = "1"))

                .Add new_BPrefsControl("fancytoggle2", "disk.watch_optical", "Watch optical media?", , mConfig.ValueOf("disk.watch_optical"), , (mConfig.ValueOf("disk.enabled") = "1"))
                .Add new_BPrefsControl("fancytoggle2", "disk.watch_removable", "Watch removable media (excludes floppy drives)?", , mConfig.ValueOf("disk.watch_removable"), , (mConfig.ValueOf("disk.enabled") = "1"))
'                .Add new_BPrefsControl("label", "", "Note that floppy drives are never scanned, even if the above setting is enabled.")
                .Add new_BPrefsControl("label", "", "Information to include in notification:")
                .Add new_BPrefsControl("fancycycle", "disk.verbosity", "Simple|Meter|Detailed", "", mConfig.ValueOf("disk.verbosity"), , (mConfig.ValueOf("disk.enabled") = "1"))
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "disk.test", "Test Settings", , , , (mConfig.ValueOf("disk.enabled") = "1"))

            End With

            .AddPage pPage


            ' /* power */

            Set pPage = new_BPrefsPage("Battery", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "battery.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("label", "", "Warn when remaining battery life reaches:")

                Set pm = New CTempMsg
                pm.Add "min", 10&
                pm.Add "max", 90&
                pm.Add "freq", 10&
                pm.Add "label_size", 50&
                pm.Add "suffix", "%"
                pm.Add "colour", rgba(255, 172, 172)
                .Add new_BPrefsControl("fancyslider", "power.low_mark", "", "", mConfig.ValueOf("power.low_mark"), pm)
                .Add new_BPrefsControl("fancytoggle2", "power.warn_when_charging", "Show notification even if charging?", , mConfig.ValueOf("power.warn_when_charging"))

                .Add new_BPrefsControl("fancytoggle2", "power.show_text", "Display text?", , mConfig.ValueOf("power.show_text"))
                .Add new_BPrefsControl("fancytoggle2", "power.use_meter", "Display meter?", , mConfig.ValueOf("power.use_meter"))
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "power.test", "Test Settings")

            End With

            .AddPage pPage


            ' /* NIC and DUN */

            Set pPage = new_BPrefsPage("Network", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "network.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("fancytoggle2", "nic.enabled", "Monitor network and dial-up connections?", , mConfig.ValueOf("nic.enabled"))
                .Add new_BPrefsControl("label", "", "Information to include in notification:")
                .Add new_BPrefsControl("fancycycle", "nic.verbosity", "None (Icon Only)|Simple|Detailed", "", mConfig.ValueOf("nic.verbosity"), , (mConfig.ValueOf("nic.enabled") = "1"))
                .Add new_BPrefsSeparator
                .Add new_BPrefsControl("fancybutton2", "nic.test", "Test Settings (Network)", , , , (mConfig.ValueOf("nic.enabled") = "1"))
                .Add new_BPrefsControl("fancybutton2", "dun.test", "Test Settings (Dial-up / VPN)", , , , (mConfig.ValueOf("nic.enabled") = "1"))

            End With

            .AddPage pPage


            ' /* Other */

            Set pPage = new_BPrefsPage("Other", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "keyboard.png"), Me)
            With pPage
                .SetMargin 32

                .Add new_BPrefsControl("banner", "", "Keyboard")
                .Add new_BPrefsControl("fancytoggle2", "keyb.watch_caps", "Watch Caps Lock?", , mConfig.ValueOf("keyb.watch_caps"))
                .Add new_BPrefsControl("fancytoggle2", "keyb.watch_numl", "Watch Num Lock?", , mConfig.ValueOf("keyb.watch_numl"))
                .Add new_BPrefsControl("fancytoggle2", "keyb.watch_scrl", "Watch Scroll Lock?", , mConfig.ValueOf("keyb.watch_scrl"))

                .Add new_BPrefsControl("banner", "", "Miscellaneous")
                .Add new_BPrefsControl("fancytoggle2", "other.ie", "Notify when an Internet Explorer download completes?", , mConfig.ValueOf("other.ie"))
                .Add new_BPrefsControl("fancytoggle2", "other.watch_clipboard", "Notify when clipboard content changes?", , mConfig.ValueOf("other.watch_clipboard"))

'                .Add new_BPrefsSeparator
'                .Add new_BPrefsControl("fancybutton2", "nic.test", "Test Settings (Network)")
'                .Add new_BPrefsControl("fancybutton2", "dun.test", "Test Settings (Dial-up / VPN)")

            End With

            .AddPage pPage


            ' /* about page */

            Set pPage = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "about.png"), Me)
            With pPage
                .SetMargin 32
                .Add new_BPrefsControl("label", "", "SysInfo Extension R" & LIB_RELEASE)
                .Add new_BPrefsControl("label", "", "Version " & CStr(App.Major) & "." & CStr(App.Revision) & " (" & LIB_DATE & ")")
                .Add new_BPrefsControl("label", "", App.LegalCopyright)
                .Add new_BPrefsControl("label", "", "")
                .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net")

            End With

            .AddPage pPage

            .Go
'            g_SetWindowIconToAppResourceIcon .hWnd

        End With

    End If

    ISnarlExtension_GetConfigWindow = mPanel.hwnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "fullphat products"
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = EXT_COMMENT
        .Flags = SNARL_EXTN_WANTS_PULSE Or SNARL_EXTN_IS_CONFIGURABLE
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = EXT_NAME
        .Path = App.Path
        .Release = LIB_RELEASE
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long

    On Error Resume Next

    l3OpenLog "%APPDATA%\full phat\snarl\" & EXT_NAME & ".log", True
    g_Debug "pwd: " & g_MakePath(App.Path)

    If sn41GetVersion = 0 Then
        g_Debug "Needs Snarl V41 or better", LEMON_LEVEL_CRITICAL
        mLastErr = "Requires Snarl V41 or greater"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    Err.Clear
    get_host_name
    If Err.Number <> 0 Then
        g_Debug "misc.resource missing", LEMON_LEVEL_CRITICAL
        mLastErr = "Requires misc.resource"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

    ' /* create window */

    g_Debug "creating window..."

    Set theHandler = New TWindow
    If Not theHandler.IsValid Then
        ' /* no window */
        Set theHandler = Nothing
        mLastErr = "Couldn't create message looper"
        ISnarlExtension_Initialize = -1
        Exit Function

    End If

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()
Dim szPath As String
Dim i As Integer

    On Error Resume Next

    mToken = 0

    ' /* load config */

    Set mConfig = New CConfFile
    mConfig.SetTo sn41GetConfigPath() & "sysinfo.config"

    uCfgCheck "cpu.enabled", "1"        ' // R3.04
    uCfgCheck "cpu.high_mark", "90"
    uCfgCheck "cpu.duration", "10"

    uCfgCheck "power.low_mark", "20"
    uCfgCheck "power.warn_when_charging", "0"
    uCfgCheck "power.show_text", "0"
    uCfgCheck "power.use_meter", "1"

    uCfgCheck "ram.low_mark", "20"
'        uCfgCheck "ram.show_title", "1"     ' // include title
    uCfgCheck "ram.show_text", "1"      ' // include help text
    uCfgCheck "ram.use_meter", "1"      ' // meter instead of text (low RAM only)

    uCfgCheck "nic.enabled", "1"        ' // R3.04
    uCfgCheck "nic.verbosity", "3"

    uCfgCheck "disk.enabled", "1"       ' // R3.04
    uCfgCheck "disk.low_mark", "10"
    uCfgCheck "disk.watch_removable", "1"
    uCfgCheck "disk.watch_optical", "0"
    uCfgCheck "disk.verbosity", "3"

    uCfgCheck "keyb.watch_caps", "1"
    uCfgCheck "keyb.watch_numl", "0"
    uCfgCheck "keyb.watch_scrl", "0"

    uCfgCheck "other.ie", "0"
    uCfgCheck "other.watch_clipboard", "0"

    mConfig.Save

    ' /* register */

    mToken = sn41RegisterApp("application/vnd.k23-" & LCase$(EXT_NAME), EXT_NAME, g_MakePath(App.Path) & "icon.png")
    If mToken = 0 Then _
        Exit Sub

    sn41AddClass mToken, CLASS_LOW_MEM, "Physical memory low", False
    sn41AddClass mToken, CLASS_NO_MEM, "Physical memory exhausted", False
    sn41AddClass mToken, CLASS_HIGH_CPU, "High processor utilisation"
    sn41AddClass mToken, CLASS_LOW_DISK, "Low disk space", False
    sn41AddClass mToken, CLASS_NIC_CONN, "Network connected", False
    sn41AddClass mToken, CLASS_NIC_DISC, "Network disconnected", False
    sn41AddClass mToken, CLASS_VOL_MNT, "Volume mounted"
    sn41AddClass mToken, CLASS_VOL_UMNT, "Volume unmounted"
    sn41AddClass mToken, CLASS_POWER_CONNECT, "Power connected"
    sn41AddClass mToken, CLASS_POWER_DISCONNECT, "Power disconnected"
    sn41AddClass mToken, CLASS_POWER_LOW, "Low power warning"
    sn41AddClass mToken, CLASS_LOCK_KEY, "Caps, Num or Scroll lock activated"
    sn41AddClass mToken, CLASS_IP_CHANGE, "IP address change"
    ' /* R3.05 */
    sn41AddClass mToken, CLASS_IE_DOWNLOAD, "Internet Explorer download completed"
    sn41AddClass mToken, CLASS_CLIPBOARD, "Clipboard content changed"

    Set mAdapterList = new_BTagList()
    Set theVolRoster = New TSimpleVolRoster


    ' /* call kernel info once to get baseline */

    krnl_GetSystemInfo
    g_Debug "[Start]: krnl_GetProcessorCount()=" & CStr(krnl_GetProcessorCount())
    krnl_GetCurrentTotalCPUPercent

    ' /* configure the handler */

    With theHandler
        .SetLowPowerThreshold Val(mConfig.ValueOf("power.low_mark"))
        .SetLowRAMThreshold Val(mConfig.ValueOf("ram.low_mark"))
'        .SetHighCPUThreshold Val(mConfig.ValueOf("cpu.high_mark"))
'        .SetHighCPUDuration Val(mConfig.ValueOf("cpu.duration"))
'        .SetLowDiskThreshold Val(mConfig.ValueOf("disk.low_mark"))
'        .SetWatchOptical (mConfig.ValueOf("disk.watch_optical") = "1")
'        .SetWatchRemovable (mConfig.ValueOf("disk.watch_removable") = "1")

        .ConfigChanged "cpu.high_mark", mConfig.ValueOf("cpu.high_mark")
        .ConfigChanged "cpu.duration", mConfig.ValueOf("cpu.duration")
        .ConfigChanged "cpu.enabled", mConfig.ValueOf("cpu.enabled")            ' // set last

        .ConfigChanged "disk.low_mark", mConfig.ValueOf("disk.low_mark")
        .ConfigChanged "disk.watch_optical", mConfig.ValueOf("disk.watch_optical")
        .ConfigChanged "disk.watch_removable", mConfig.ValueOf("disk.watch_removable")
        .ConfigChanged "disk.enabled", mConfig.ValueOf("disk.enabled")            ' // set last

        .ConfigChanged "other.watch_clipboard", mConfig.ValueOf("other.watch_clipboard")

        .SetEnabled True

    End With

    ' /* NIC/DUN and volume scanning is handled by this object */

    uEnableNICScanning (mConfig.ValueOf("nic.enabled") = "1")
    uEnableVolScanning True

End Sub

Private Sub ISnarlExtension_Stop()

    uEnableNICScanning False
    uEnableVolScanning False

    sn41UnregisterApp mToken
    mToken = 0

End Sub

Private Sub ISnarlExtension_TidyUp()

    Set theHandler = Nothing

End Sub

Private Sub theDUNRoster_Connected(ByVal Name As String, Details As melon4.BTagList)

    If mToken = 0 Then _
        Exit Sub

Dim szTitle As String
Dim szText As String

    Select Case Val(mConfig.ValueOf("nic.verbosity"))
    Case 1
        ' /* icon only */

    Case 2
        ' /* simple */
        szTitle = Name
        szText = ""

    Case Else
        ' /* detailed */
        szTitle = g_Capitalize(Details.ValueOf("device_type", "Dial-up"), False) & " connected"
        szText = Name & " is now connected at " & uSpeed(Val(Details.ValueOf("speed", "0"))) & " using " & Details.ValueOf("device_name")

    End Select

Dim i As Long

    i = sn41EZNotify(mToken, _
                     CLASS_DUN_CONN, _
                     szTitle, _
                     szText, , "!dun-connected")

End Sub

Private Sub theDUNRoster_Disconnected(ByVal Name As String, Details As melon4.BTagList)

    If mToken = 0 Then _
        Exit Sub

Dim szTitle As String
Dim szText As String

    Select Case Val(mConfig.ValueOf("nic.verbosity"))
    Case 1
        ' /* icon only */

    Case 2
        ' /* simple */
        szTitle = Name
        szText = ""

    Case Else
        ' /* detailed */
        szTitle = g_Capitalize(Details.ValueOf("device_type", "Dial-up"), False) & " disconnected"
        szText = Name & " is disconnected"

    End Select

Dim i As Long

    i = sn41EZNotify(mToken, CLASS_DUN_DISC, _
                     szTitle, _
                     szText, , "!dun-disconnected")

End Sub

Private Sub theHandler_ClipboardChanged()
Dim pContent As BTagList
Dim bShow As Boolean
Dim szcb As String
Dim sz As String
Dim szic As String
Dim hr As Long
Dim szText As String

    On Error Resume Next

    g_Debug "Extension/theHandler_ClipboardChanged()", LEMON_LEVEL_PROC_ENTER

    Set pContent = new_BTagList()

    If OpenClipboard(0) Then
        Err.Clear
        hr = EnumClipboardFormats(0)

        Do While hr
            ' /* add the content */
            g_Debug g_HexStr(hr, 4) & ": " & clipboard_GetFormatName(hr)

            Select Case hr
            Case CF_TEXT
                pContent.Add new_BTagItem(CStr(hr), clipboard_GetAsString())

            Case Else
                pContent.Add new_BTagItem(CStr(hr), "")

            End Select

            sz = sz & clipboard_GetFormatName(hr) & " (" & g_HexStr(hr) & ")" & vbCrLf
            hr = EnumClipboardFormats(hr)

        Loop

        CloseClipboard

        If Err.LastDllError = 0 Then
            g_Debug "count=" & pContent.CountItems
            If pContent.CountItems > 0 Then
                ' /* apply some intelligence to what's been copied */

                If pContent.IndexOf(CStr(CF_TEXT)) Then
                    ' /* standard text - set notification text to it */
                    sz = pContent.ValueOf(CStr(CF_TEXT))

                    If InStr(sz, "://") <> 0 Then
                        ' /* url: set it as the default callback */
                        szcb = sz

                    ElseIf g_IsWebColour(sz, hr) Then
                        ' /* icon is a block of the colour */
                        szic = g_MakePath(App.Path) & "colourblock.png"
                        sz = "Colour " & sz & " (" & g_ToRGBNotation(hr) & ")"

                        With New mfxView
                            .SetHighColour (hr Or &HFF000000)
                            .EnableSmoothing False
                            .SizeTo 96, 96
                            .FillRect .Bounds
                            .WriteToFile szic, "image/png"

                        End With

                    End If

                ElseIf pContent.IndexOf(CStr(CF_BITMAP)) Then
                    ' /* bitmap */
                    If OpenClipboard(0) <> 0 Then
                        hr = GetClipboardData(CF_BITMAP)
                        uSaveBitmap hr, szic

                    End If

                End If

                bShow = True

            Else
                g_Debug "nothing in the clipboard", LEMON_LEVEL_CRITICAL

            End If

        Else
            g_Debug "error enumerating clipboard content: " & Err.LastDllError, LEMON_LEVEL_CRITICAL

        End If

    Else
        g_Debug "error opening clipboard", LEMON_LEVEL_CRITICAL

    End If

    If bShow Then _
        sn41EZNotify mToken, CLASS_CLIPBOARD, "Clipboard changed", sz, , IIf(szic <> "", szic, g_MakePath(App.Path) & "clipboard.png"), , szcb

    If szic <> "" Then _
        DeleteFile szic

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theHandler_LocalIPAddressChanged(ByVal Details As String)
Dim l As Long

    l = sn41EZNotify(mToken, CLASS_IP_CHANGE, "IP address change", Details, , g_MakePath(App.Path) & "ipaddrchange.png")

'    If mNotify Then _
        snShowMessageEx Me.CLASS_1, "Local IP Change", IPAddr & " added", 10, g_MakePath(App.Path) & "icon.png", Me.hWnd, &H401

End Sub

Private Sub theHandler_OnCPUMonitoringStarted()

    Debug.Print "theHandler_OnCPUMonitoringStarted"

'                l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Started monitoring" & vbCrLf & _
'                                 "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")

End Sub

Private Sub theHandler_OnCPUMonitoringStopped()

    Debug.Print "theHandler_OnCPUMonitoringStopped"

'            l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Stopped monitoring" & vbCrLf & _
'                             "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")

End Sub

Private Sub theHandler_OnCPUThresholdReached(ByVal LoadPercent As Integer, ByVal Duration As Integer)
Dim l As Long

    l = sn41EZNotify(mToken, CLASS_HIGH_CPU, "High processor utilisation", _
                     "The system has been at least " & CStr(LoadPercent) & "% busy for " & CStr(Duration) & " seconds", , _
                     "!system-warning", , , , "action::Task Manager...,taskmgr.exe#?action::Run...,!system run_dialog#?action::Shutdown...,!system shutdown_dialog")

End Sub

Private Sub theHandler_OnDiskThresholdReached(ByVal FreePercent As Integer, ByVal Path As String)

    If mToken = 0 Then _
        Exit Sub

Dim szTitle As String
Dim szText As String
Dim szExtra As String

    szTitle = Path

'    With New TSimpleVolume
'        If .SetTo(Path) Then
'            If .Label <> "" Then _
'                szTitle = Chr$(34) & .Label & Chr$(34)
'
'        End If
'
'    End With

    szTitle = szTitle & " is nearly full"

    Select Case Val(mConfig.ValueOf("disk.verbosity"))
    Case 1
        ' /* simple */
        szText = FreePercent & "% remaining"

    Case 2
        ' /* meter */
        szText = CStr(100 - FreePercent)

    Case 3
        ' /* detailed */
        With New TSimpleVolume
            If .SetTo(Path) Then
                szText = g_FileSizeToStringEx2(.Available) & " of " & g_FileSizeToStringEx2(.Capacity) & " remaining"
                szExtra = "value-percent::" & CStr(100 - FreePercent)

            Else
                szText = "??" & Path

            End If

        End With

    End Select

Dim l As Long

    ' /* verbosity      title           text
    '       1           "disk low"      path
    '       2                           % as meter
    '       3                           "x gb of y gb (%) remaining
    ' */

    l = sn41EZNotify(mToken, _
                     CLASS_LOW_DISK, _
                     szTitle, _
                     szText, , "!disk-low_space", , , , "action::Open," & Path & IIf(szExtra <> "", "#?" & szExtra, ""))

'    If l Then _
        sn42AddAction l, "Open", Path

End Sub

Private Sub theHandler_OnLockKeyPressed(ByVal Key As melon4.B_MODIFIER_KEYS, ByVal Enabled As Boolean)

    If mToken = 0 Then _
        Exit Sub

    Select Case Key
    Case B_CAPS_LOCK
        If (Enabled) And (mConfig.ValueOf("keyb.watch_caps") = "1") Then
            mCapsId = sn41EZNotify(mToken, CLASS_LOCK_KEY, "Caps Lock on", "", , "!system-warning")

        ElseIf Not Enabled Then
            sn41Hide mCapsId

        End If

    Case B_NUM_LOCK
        If (Enabled) And (mConfig.ValueOf("keyb.watch_numl") = "1") Then
            mNumId = sn41EZNotify(mToken, CLASS_LOCK_KEY, "Num Lock on", "", , "!system-warning")

        ElseIf Not Enabled Then
            sn41Hide mNumId

        End If

    Case B_SCROLL_LOCK
        If (Enabled) And (mConfig.ValueOf("keyb.watch_scrl") = "1") Then
            mScrlId = sn41EZNotify(mToken, CLASS_LOCK_KEY, "Scroll Lock on", "", , "!system-warning")

        ElseIf Not Enabled Then
            sn41Hide mScrlId

        End If

    End Select

End Sub

Private Sub theHandler_OnLowPowerWarning(ByVal CurrentChargePercent As Integer)

    g_Debug "theHandler.OnLowPowerWarning()", LEMON_LEVEL_PROC_ENTER

    ' /* show even if charging? */

    If (get_power_status() = B_POWER_STATE_CONNECTED) And (mConfig.ValueOf("power.warn_when_charging") = "0") Then _
        Exit Sub

Dim sz As String

    ' /* matrix */
    '
    '   show_text   use_meter
    '       0           0           icon only
    '       0           1           icon with meter
    '       1           0           title and text
    '       1           1           title and meter

    sz = CStr(CurrentChargePercent)

    If mConfig.ValueOf("power.show_text") = "1" Then
        If mConfig.ValueOf("power.use_meter") = "0" Then _
            sz = "Current battery charge is " & sz & "%"

    ElseIf mConfig.ValueOf("power.use_meter") = "0" Then
        sz = ""

    End If

    mLowPowerId = sn41EZNotify(mToken, _
                               CLASS_POWER_LOW, _
                               IIf(mConfig.ValueOf("power.show_text") = "1", "Battery getting low", ""), _
                               sz, _
                               , _
                               "!battery-low", _
                               1)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theHandler_OnPhysicalRAMGone()

    ' /* if the low RAM warning is still visible, remove it */

    If sn41IsVisible(mLowRamId) Then _
        sn41Hide mLowRamId

Dim sz As String

    If mConfig.ValueOf("ram.show_text") = "1" Then _
        sz = "There is no physical memory available - Windows will begin to move memory to and from disk (paging) which will " & _
             "make the computer respond more slowly." & vbCrLf & vbCrLf & _
             "Closing unwanted applications, browser tabs or system tray applications will help restore physical memory."

'    ElseIf mConfig.ValueOf("ram.show_title") = "0" Then
'        sz = "Physical memory exhausted"
'
'    End If

    sn41EZNotify mToken, _
                 CLASS_NO_MEM, _
                 "Physical memory exhausted", _
                 sz, , "!system-warning"

End Sub

Private Sub theHandler_OnPhysicalRAMLow(ByVal PercentRemaining As Integer)
Dim sz As String

    If mConfig.ValueOf("ram.use_meter") = "1" Then
        sz = CStr(PercentRemaining)

    Else
        sz = g_FileSizeToStringEx2(g_GetPhysMem(False), , , "0.0") & " of " & g_FileSizeToStringEx2(g_GetPhysMem(True), , , "0.0") & _
             " (" & CStr(PercentRemaining) & "%) remaining"

    End If

    mLowRamId = sn41EZNotify(mToken, _
                             CLASS_LOW_MEM, _
                             "Physical memory low", _
                             sz, , "!system-info")

End Sub

Private Sub theHandler_OnPowerStatusChanged(ByVal Connected As Boolean, ByVal CurrentChargePercent As Integer)

    g_Debug "theHandler.OnPowerStatusChanged()", LEMON_LEVEL_PROC_ENTER

Dim sz As String

    ' /* matrix */
    '
    '   show_text   use_meter
    '       0           0           icon only
    '       0           1           N/A
    '       1           0           icon and title
    '       1           1           N/A

    sz = CStr(CurrentChargePercent)

    If mConfig.ValueOf("power.show_text") = "1" Then
        If mConfig.ValueOf("power.use_meter") = "0" Then _
            sz = "Current battery charge is " & sz & "%"

    ElseIf mConfig.ValueOf("power.use_meter") = "0" Then
        sz = ""

    End If

    ' /* remove existing notification */

    If sn41IsVisible(mPowerId) Then _
        sn41Hide mPowerId

    If Connected Then

        ' /* also remove low power warning if it's still on screen */

        If sn41IsVisible(mLowPowerId) Then _
            sn41Hide mLowPowerId

        mPowerId = sn41EZNotify(mToken, _
                                       CLASS_POWER_CONNECT, _
                                       IIf(mConfig.ValueOf("power.show_text") = "1", "Power connected", ""), _
                                       sz, _
                                       , _
                                       "!power-connected" _
                                       )

    Else
        mPowerId = sn41EZNotify(mToken, _
                                       CLASS_POWER_DISCONNECT, _
                                       IIf(mConfig.ValueOf("power.show_text") = "1", "Power disconnected", ""), _
                                       sz, _
                                       , _
                                       "!power-disconnected" _
                                       )

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub theNetRoster_Connected(ByVal Name As String, Device As melon4.BTagList)
Dim szTitle As String
Dim szText As String
Dim i As Long

'    sz = Device.ValueOf("name")

    If mToken Then
    
        Select Case Val(mConfig.ValueOf("nic.verbosity"))
        Case 1
            ' /* icon only */

        Case 2
            ' /* simple */
            szTitle = Name
            szText = ""

        Case Else
            ' /* detailed */
            szTitle = "Network connected"
            szText = Name & " is now connected" & vbCrLf & "Speed: " & Device.ValueOf("formatted_speed", "0") '& _
                     IIf(sz <> "", vbCrLf & "Device: " & sz, "")

        End Select

        i = sn41EZNotify(mToken, _
                         CLASS_NIC_CONN, _
                         szTitle, _
                         szText, , _
                         "!network-connected")

'        sn41EZNotify mToken, "", "Connected", _
'                     Name & vbCrLf & _
'                     "Speed: " & Details.ValueOf("formatted_speed") & vbCrLf & _
'                     "Type: " & Details.ValueOf("type") & vbCrLf & _
'                     "Media: " & Details.ValueOf("media_type"), , _
'                     "!network-connected"

    End If

End Sub

Private Sub theNetRoster_Disconnected(ByVal Name As String, ByVal DeviceType As Long)
Dim szTitle As String
Dim szText As String
Dim i As Long

    If mToken Then

        Select Case Val(mConfig.ValueOf("nic.verbosity"))
        Case 1
            ' /* icon only */

        Case 2
            ' /* simple */
            szTitle = Name
            szText = ""

        Case Else
            ' /* detailed */
            szTitle = "Network disconnected"
            szText = Name & " is disconnected"

        End Select

        i = sn41EZNotify(mToken, _
                         CLASS_NIC_DISC, _
                         szTitle, _
                         szText, , _
                         "!network-disconnected")

    End If

End Sub

Private Sub theNICTimer_Pulse()

    theNetRoster.Refresh
    theDUNRoster.Refresh

End Sub

Private Function uSpeed(ByVal Speed As Long) As String

    If Speed > 999999999 Then
        uSpeed = CStr(Speed / 1000000000) & "Gbps"

    ElseIf Speed > 999999 Then
        uSpeed = CStr(Speed / 1000000) & "Mbps"

    ElseIf Speed > 999 Then
        uSpeed = CStr(Speed / 1000) & "Kbps"

    Else
        uSpeed = CStr(Speed) & "bps"

    End If

End Function

Private Sub theVolRoster_VolumeMounted(ByVal Path As String, ByVal Label As String, ByVal VolType As Long)
Dim sz As String
Dim n As Integer
Dim i As Long

    sz = uCapacityInfo(Path, n)

    i = sn41EZNotify(mToken, CLASS_VOL_MNT, _
                     "Voume mounted", _
                     uLabel(Label, Path) & sz, , uIcon(Path), , Path, n)

End Sub

Private Sub theVolRoster_VolumeUnmounted(ByVal Path As String, ByVal Label As String, ByVal VolType As Long)
Dim i As Long

    i = sn41EZNotify(mToken, CLASS_VOL_UMNT, _
                     "Volume unmounted", _
                     uLabel(Label, Path), , "!vol-unmounted", , , , "merge::1")

End Sub

Private Function uLabel(ByVal Label As String, Optional ByVal Path As String, Optional ByVal AddPath As Boolean = True) As String

    uLabel = IIf(Label = "", "<untitled>", Label)

    If AddPath Then _
        uLabel = uLabel & " (" & Path & ")"

End Function

Private Function uCapacityInfo(ByVal Path As String, ByRef PercentUsed As Integer) As String
Dim pv As TSimpleVolume

    PercentUsed = -1

    Set pv = New TSimpleVolume
    If Not pv.SetTo(Path) Then _
        Exit Function

    If (pv.IsNetworkVolume) Or (Not pv.HasCapacityInfo) Then _
        Exit Function

    PercentUsed = pv.PercentUsed
    uCapacityInfo = vbCrLf & g_FileSizeToStringEx2(pv.Capacity) & ", " & g_FileSizeToStringEx2(pv.Used) & " (" & CStr(pv.PercentUsed) & "%) used"

End Function

Private Function uIcon(ByVal Path As String) As String

    Select Case GetDriveType(Path)
    Case DRIVE_FIXED
        uIcon = "!dev-hd"

    Case DRIVE_CDROM
        uIcon = "!dev-cd"

    Case DRIVE_REMOVABLE
        uIcon = "!dev-fd"

    Case Else
        uIcon = "!vol-mounted"

    End Select

End Function

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)

    If (mConfig Is Nothing) Then _
        Exit Sub

    mConfig.Update Control.GetName, Value
    mConfig.Save

Dim pc As BControl

    Select Case Control.GetName

    Case "cpu.enabled"
        If Control.Page.Find("cpu.high_mark", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("cpu.duration", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("cpu.test", pc) Then _
            pc.SetEnabled (Value = "1")


    Case "power.low_mark"
        theHandler.SetLowPowerThreshold Val(Value)

    Case "ram.low_mark"
        theHandler.SetLowRAMThreshold Val(Value)


    Case "disk.enabled"
        If Control.Page.Find("disk.low_mark", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.watch_optical", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.watch_removable", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.verbosity", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("disk.test", pc) Then _
            pc.SetEnabled (Value = "1")


    Case "nic.enabled"
        
        uEnableNICScanning (Value = "1")

        If Control.Page.Find("nic.verbosity", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("nic.test", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("dun.test", pc) Then _
            pc.SetEnabled (Value = "1")


    End Select

    theHandler.ConfigChanged Control.GetName, Value

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Select Case Control.GetName
    Case "power.test"
        theHandler_OnPowerStatusChanged (get_power_status = B_POWER_STATE_CONNECTED), "50"
        theHandler_OnLowPowerWarning "50"

    Case "ram.test"
        theHandler_OnPhysicalRAMGone
        theHandler_OnPhysicalRAMLow 100 - g_PercentMemInUse(True)

    Case "cpu.test"
        theHandler_OnCPUThresholdReached mConfig.ValueOf("cpu.high_mark"), mConfig.ValueOf("cpu.duration")

    Case "nic.test"
        theNetRoster_Connected "<test_network>", new_BTagList()
        theNetRoster_Disconnected "<test_network>", 0

    Case "dun.test"
        theDUNRoster_Connected "<test_dialup>", new_BTagList()
        theDUNRoster_Disconnected "<test_dialup>", new_BTagList()

    Case "disk.test"
        theHandler_OnDiskThresholdReached mConfig.ValueOf("disk.low_mark"), g_SafeLeftStr(App.Path, 3)

    Case "go_website"
        ShellExecute Control.Page.Panel.hwnd, "open", "http://www.fullphat.net", vbNullString, vbNullString, 5

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
'    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"
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

Private Sub uCfgCheck(ByVal Name As String, ByVal Default As String)

    If (mConfig Is Nothing) Then _
        Exit Sub

    If Not mConfig.Exists(Name) Then _
        mConfig.Add Name, Default

End Sub




'Private Sub theNetRoster_Connected(Details As melon4.BTagList)
'Static i As Long
'
'    If Details.ValueOf("active") = "1" Then _
'        i = sn41EZNotify(mtoken, CLASS_NIC_CONN, "Network connected", _
'                         Details.ValueOf("connection") & " is now connected at " & _
'                         uSpeed(Val(Details.ValueOf("speed", "0"))) & _
'                         " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")
'
'End Sub
'
'Private Sub theNetRoster_Disconnected(Details As melon4.BTagList)
'Static i As Long
'
'    ' /* was the adapter active at the time it disappeared? */
'
'    If Details.ValueOf("active") = "1" Then _
'        i = sn41EZNotify(mtoken, CLASS_NIC_DISC, "Network disconnected", _
'                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
'
'End Sub
'
'Private Sub theNetRoster_StatusChanged(Details As melon4.BTagList, ByVal PreviousState As Long, ByVal NewState As Long)
'Static i As Long
'
'    If (PreviousState = IF_OPER_STATUS_OPERATIONAL) And (NewState = IF_OPER_STATUS_NON_OPERATIONAL) Then
'        ' /* gone from connected to disconnected */
'        i = sn41EZNotify(mtoken, CLASS_NIC_DISC, "Network disconnected", _
'                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
'
'    ElseIf (PreviousState = IF_OPER_STATUS_NON_OPERATIONAL) And (NewState = IF_OPER_STATUS_OPERATIONAL) Then
'        ' /* gone from disconnected to connected */
'        i = sn41EZNotify(mtoken, CLASS_NIC_CONN, "Network connected", _
'                         Details.ValueOf("connection") & " is now connected at " & _
'                         uSpeed(Val(Details.ValueOf("speed", "0"))) & _
'                         " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")
'
'    Else
'        i = sn41EZNotify(mtoken, CLASS_NIC_CONN, "Network status changed", _
'                         Details.ValueOf("connection") & " status changed from " & CStr(PreviousState) & " to " & CStr(NewState), , "")
'
'    End If
'
'End Sub

'Private Sub theCPUTimer_Pulse()
'Dim dCurrentCPU As Double
'Dim t As Long
'Dim l As Long
'
'    dCurrentCPU = krnl_GetCurrentTotalCPUPercent()
'
'    If dCurrentCPU > mSettings.CPUThreshold Then
'        ' /* over the threshold */
'
'        If mSettings.CPUHitMonitoring = 0 Then
'            ' /* first time (not tracking) */
'            If mSettings.TellMeWhenMonitoringStarts Then _
'                l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Started monitoring" & vbCrLf & _
'                                 "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")
'
'            mSettings.CPUHitMonitoring = GetTickCount()
'
'        ElseIf Not mSettings.NotifiedHighCPU Then
'            ' /* already tracking */
'
'            t = GetTickCount() - mSettings.CPUHitMonitoring
'            If t > (mSettings.CPUHitDuration * 1000) Then
'                ' /* and past the cut-off duration */
'                l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", _
'                                 "The system has been at least " & mSettings.CPUThreshold & "% busy for " & CStr(t \ 1000) & " seconds", , "!system-info")
'
'                mSettings.NotifiedHighCPU = True
'
'            End If
'
'        End If
'
'    ElseIf mSettings.CPUHitMonitoring > 0 Then
'        ' /* we were monitoring so stop */
'        If mSettings.TellMeWhenMonitoringStops Then _
'            l = sn41EZNotify(mtoken, CLASS_HIGH_CPU, "High processor utilisation", "Stopped monitoring" & vbCrLf & _
'                             "Current utilisation: " & CStr(dCurrentCPU) & "%", , "!system-info")
'
'        mSettings.CPUHitMonitoring = 0
'        mSettings.NotifiedHighCPU = False
'
'    End If
'
'    ' /* if we have a valid notification token and we're running V42+, add actions */
'
'    If (l) And (sn41GetVersion() > 41) Then
'        sn42AddAction l, "Task Manager...", "taskmgr.exe"
'        sn42AddAction l, "Run...", "!system run_dialog"
'        sn42AddAction l, "Shutdown...", "!system shutdown_dialog"
'
'    End If
'
'End Sub
'Private Sub uCheckDiskSpace()
''Const DISK_LIMIT = 90
'Dim szPath As String
'Dim szLabel As String
'Dim lType As Long
'Dim l As Long
'
'    With theVolRoster
'        .Rewind
'
'        Do While .GetNextVolume(szPath, szLabel, lType)
'            If (lType = DRIVE_FIXED) Or (lType = DRIVE_REMOVABLE) Then
'                With New TSimpleVolume
'                    If .SetTo(szPath) Then
'                        If (mVolsDone.IndexOf(szPath) = 0) And (.PercentUsed >= mSettings.DiskThreshold) Then
'
'                            l = sn41EZNotify(mtoken, CLASS_LOW_DISK, uLabel(szLabel, szPath) & " is nearly full", _
'                                             g_FileSizeToStringEx2(.Available) & " of " & g_FileSizeToStringEx2(.Capacity) & " remaining", , "!system-info")
'
'                            mVolsDone.Add new_BTagItem(szPath, "")
'
'                        ElseIf mVolsDone.IndexOf(szPath) Then
'                            ' /* have we dropped below the limit? */
'                            If .PercentUsed < mSettings.DiskThreshold Then _
'                                mVolsDone.Remove mVolsDone.IndexOf(szPath)
'
'                        End If
'
'                    Else
'                        g_Debug "Extension.uCheckDiskSpace(): can't create CVolume for '" & szPath & "'", LEMON_LEVEL_CRITICAL
'
'                    End If
'                End With
'            End If
'        Loop
'    End With
'
'End Sub

'Private Function uIsConnected(ByVal OperStatus As Long) As String
'
'    Select Case OperStatus
'    Case IF_OPER_STATUS_OPERATIONAL
'        uIsConnected = "1"
'
'    Case Else
'        uIsConnected = "0"
'
'    End Select
'
'End Function

'Private Function uFind(ByVal dwIndex As Long) As Long
'
'    If mIFCount = 0 Then _
'        Exit Function
'
'Static pif As MIB_IFROW
'Static cb As Long
'Static i As Long
'
'    cb = Len(pif)
'    For i = 1 To mIFCount
'        ' /* retrieve a single MIB_IFROW */
'        CopyMemory pif, mIFBuffer(4 + (i - 1) * cb), cb
'        If pif.dwIndex = dwIndex Then
'            uFind = i
'            Exit Function
'
'        End If
'
'    Next i
'
'End Function

Private Sub theVolTimer_Pulse()

    theVolRoster.Refresh

End Sub

Private Sub uEnableNICScanning(ByVal Enabled As Boolean)

    If Enabled Then

        ' /* scan the current NIC and DUN adapters */

        Set theNetRoster = New TNetRoster2
        With theNetRoster
            .LockRoster
            .Refresh
            .UnlockRoster

        End With

        Set theDUNRoster = New TDUNRoster
        With theDUNRoster
            .LockRoster
            .Refresh
            .UnlockRoster

        End With

        ' /* start watching */

        Set theNICTimer = new_BTimer(4000)
        g_Debug "uEnableNICScanning(): enabled"

    Else
        ' /* kill the timer */
        Set theNICTimer = Nothing

        ' /* delete the rosters */
        Set theDUNRoster = Nothing
        Set theNetRoster = Nothing
        g_Debug "uEnableNICScanning(): disabled"

    End If

End Sub

Private Sub uEnableVolScanning(ByVal Enabled As Boolean)

    If Enabled Then

        ' /* scan the current mounted volumes */

        With theVolRoster
            .LockRoster
            .SetWatchedVolumes False, True
            .Refresh
            .UnlockRoster

        End With

        ' /* start watching */

        g_Debug "uEnableVolScanning(): enabled"
        Set theVolTimer = new_BTimer(2000)

    Else
        ' /* kill the timer */
        Set theVolTimer = Nothing
        g_Debug "uEnableVolScanning(): disabled"

    End If

End Sub

Private Sub uSaveBitmap(ByVal hBitmap As Long, ByRef Path As String)

    If hBitmap = 0 Then _
        Exit Sub

Dim pbm As Long
Dim pbx As mfxBitmap
Dim cx As Long
Dim cy As Long
Dim sz As String

    If GdipCreateBitmapFromHBITMAP(hBitmap, 0, pbm) = Ok Then
        GdipGetImageWidth pbm, cx
        GdipGetImageHeight pbm, cy

        With New mfxView
            .SizeTo cx, cy
            GdipDrawImagePointRectI .Handle, pbm, 0, 0, 0, 0, cx, cy, UnitPixel
            Path = "c:\" & g_CreateGUID(True) & ".bmp"
            .WriteToFile Path, "image/bmp"

        End With

    End If

End Sub
