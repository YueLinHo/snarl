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
    '/  Created:        21-Jan-2010
    '/
    '/  © 2010 full phat products
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

Private Const LIB_DATE = "17-Dec-2010"
Private Const LIB_REL = "2.0"

'Private Const WINDOW_CLASS = "w>audiomon_extension"

'Private Const CLASS_AUDIO_LEVEL_CHANGED = "audc"
'Private Const CLASS_AUDIO_MUTED = "audm"
'Private Const CLASS_AUDIO_UNMUTED = "audu"
Private Const CLASS_AUDIO_CHANGE = "audx"

Dim mLastErr As String
Dim mAppToken As Long

Dim mConfigFile As ConfigFile
Dim mSettings As ConfigSection

Dim mPanel As BPrefsPanel           ' // our prefs panel

Dim WithEvents InputMixer As BAudioMixer
Attribute InputMixer.VB_VarHelpID = -1
'Dim WithEvents VistaMixer As TVistaMixer

'Dim mVolId As Long
'Dim mMuteId As Long



Dim mVol As Long
Dim mCurrentId As Long

Private Const CLASS_NAME = "w>audiomon"
Dim mhWnd As Long

Implements KPrefsPanel              ' // we're a panel handler
Implements KPrefsPage               ' // we're also a page handler
Implements ISnarlExtension
Implements MWndProcSink

Private Sub InputMixer_ControlChanged(Control As audio_resource.BMixerLineControl, ByVal Value As Long)

    If mAppToken = 0 Then
        g_Debug "InputMixer.ControlChanged(): not registered with Snarl", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    Select Case LCase$(Control.Name)
    Case "master volume"
        If mSettings.GetValueWithDefault("show_mute", "1") = "1" Then _
            uVolChanged Control, Value, g_SafeLong(mSettings.GetValueWithDefault("vol_detail", "2"))

    Case "master mute"
        If mSettings.GetValueWithDefault("show_vol", "1") = "1" Then _
            uMuteChanged Value, g_SafeLong(mSettings.GetValueWithDefault("mute_detail", "2"))

    Case Else
        g_Debug "InputMixer.ControlChanged(): unrecognised control '" & Control.Name & "'", LEMON_LEVEL_WARNING

    End Select

End Sub

Private Function ISnarlExtension_GetConfigWindow() As Long
Dim pPage As BPrefsPage

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetWindow 1                            ' // so Go() doesn't actually display the window...
            .SetWidth 380
            .SetTitle App.Title & " Preferences"

            Set pPage = new_BPrefsPage("", Nothing, Me)

            With pPage
                .SetMargin 68

                .Add new_BPrefsControl("banner", "", "General")
'                .Add new_BPrefsControl("fancytoggle2", "MultiNotificationMode", "Use separate mute and volume change classes?", , mSettings.GetValueWithDefault("MultiNotificationMode", "1"))
                .Add new_BPrefsControl("fancytoggle2", "MeterStyleMode", "Show volume as a meter?", , mSettings.GetValueWithDefault("MeterStyleMode", "0"))


                .Add new_BPrefsControl("banner", "", "Volume Mute/Unmute")
                .Add new_BPrefsControl("fancytoggle2", "show_mute", "Notify when muted/unmuted?", , mSettings.GetValueWithDefault("show_mute", "1"))
                .Add new_BPrefsControl("fancycycle", "mute_detail", "None (Icon Only)|Simple|Detailed", "Content:", mSettings.GetValueWithDefault("mute_detail", "2"))

                .Add new_BPrefsControl("banner", "", "Volume Level Change")
                .Add new_BPrefsControl("fancytoggle2", "show_vol", "Notify when volume changes?", , mSettings.GetValueWithDefault("show_vol", "1"))
                .Add new_BPrefsControl("fancycycle", "vol_detail", "None (Icon Only)|Simple|Detailed", "Content:", mSettings.GetValueWithDefault("vol_detail", "2"))
'                .Add new_BPrefsControl("fancytoggle2", "xxxxx", "Show title?", , mSettings.GetValueWithDefault("xxxx", "0"), , False)
'                .Add new_BPrefsControl("fancytoggle2", "xxxxx", "Show current volume?", , mSettings.GetValueWithDefault("xxxx", "0"), , False)
'                .Add new_BPrefsControl("fancytoggle2", "xxxxx", "Show volume as meter?", , mSettings.GetValueWithDefault("xxxx", "0"), , False)

                .Add new_BPrefsControl("separator", "")
                .Add new_BPrefsControl("label", "", App.Title & " " & LIB_REL)
                .Add new_BPrefsControl("label", "", "V" & CStr(App.Major) & "." & CStr(App.Revision) & " (" & LIB_DATE & ")", "")
                .Add new_BPrefsControl("label", "", App.LegalCopyright)
                .Add new_BPrefsControl("fancybutton2", "go_website", "www.fullphat.net", "")
        
        '        .Add new_BPrefsControl("label", "", App.FileDescription & " V" & CStr(App.Major) & "." & CStr(App.Revision), , , , False)

            End With

            .AddPage pPage

            .Go

        End With

    End If

    ISnarlExtension_GetConfigWindow = mPanel.hWnd

End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "full phat products"
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = "Notifies when the master volume changes state"

        If g_IsWinVistaOrBetter() Then
            .Flags = 0

        Else
            .Flags = SNARL_EXTN_IS_CONFIGURABLE

        End If

        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Name = App.Title
        .Path = g_MakePath(App.Path)
        .Release = LIB_REL
        .Revision = App.Revision
        .SupportEmail = "info@fullphat.net"
        .URL = "http://www.fullphat.net"
        .Version = App.Major

    End With

End Sub

Private Function ISnarlExtension_Initialize() As Long

    g_Debug "ISnarlExtension.Initialize()", LEMON_LEVEL_PROC_ENTER

    l3OpenLog "%APPDATA%\" & App.Title & ".log"

    g_Debug "pwd is '" & g_MakePath(App.Path) & "'"

    mLastErr = ""
    g_Debug "checking o/s..."

    If g_IsWinVistaOrBetter() Then
        g_Debug "Vista/Win7 detected", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        mLastErr = "Not supported on Vista or Windows 7"
        ISnarlExtension_Initialize = 201
        Exit Function

    End If

    g_Debug "ok"

    ' /* have we got the right version of audio.resource? */

    g_Debug "checking audio.resource..."
    If Not uAudioResourceOk() Then
        g_Debug "need audio.resource V48 or later", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        mLastErr = "Needs audio.resource V48 or later"
        ISnarlExtension_Initialize = 202
        Exit Function

    End If

    g_Debug "ok"

    ' /* have we got a mixer? */

    g_Debug "checking mixers..."
    If audio_resource.count_devices < 1 Then
        g_Debug "no mixers", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        mLastErr = "No audio devices detected"
        ISnarlExtension_Initialize = 203
        Exit Function

    End If

    g_Debug "ok: " & CStr(audio_resource.count_devices) & " mixer(s) found"

Dim i As Long

    For i = 1 To audio_resource.count_devices
        With audio_resource.device_at(i)
            g_Debug .Name & ": line_count=" & .CountLines & " mfr=" & .ManufacturerId & " prd=" & .ProductId

        End With
    Next i

    ' /* all tests passed */

    Set mConfigFile = New ConfigFile
    mConfigFile.File = sn41GetConfigPath() & ".audiomon"



'    Set mHandler = New TWindow
'    mhWnd = EZAddWindow(WINDOW_CLASS, mHandler) ', , , , HWND_MESSAGE)
'    If IsWindow(mhWnd) = 0 Then
'        ' /* no window */
'        mLastErr = "Failed to create window"
'        EZUnregisterClass WINDOW_CLASS
'        Set mHandler = Nothing
'        Exit Function
'
'    End If


'    If g_IsWinVistaOrBetter Then _
        ShellExecute 0, "open", g_MakePath(App.Path) & "audiomond.exe", CStr(mhWnd), vbNullString, SW_SHOW

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()

End Sub

Private Sub ISnarlExtension_Start()

    g_Debug "ISnarlExtension.Start()", LEMON_LEVEL_PROC_ENTER

    Set InputMixer = device_at(1)
    mVol = current_master_volume_percent()

    ' /* set up the config file */

Dim i As Long

    With mConfigFile
        .Load

        i = .FindSection("general")
        If i = 0 Then
            Set mSettings = .AddSectionObj("general")

        Else
            Set mSettings = .SectionAt(i)

        End If

        ' /* defaults */

        With mSettings
            If .IndexOf("MeterStyleMode") = 0 Then _
                .Add "MeterStyleMode", "0"

'            If .IndexOf("SingleNotificationMode") = 0 Then _
                .Add "SingleNotificationMode", "1"              ' // new for 1.7

'            If .IndexOf("IncludeTitle") = 0 Then _
'                .Add "IncludeTitle", "1"
'
'            If .IndexOf("IncludeText") = 0 Then _
'                .Add "IncludeText", "1"

            ' /* alpha one had 'SingleNotificationMode' - we now default
            '    to this and use 'MultiNotificationMode' instead */

'            If .IndexOf("MultiNotificationMode") = 0 Then _
                .Add "MultiNotificationMode", "0"              ' // 1.7 Alpha 2

            If .IndexOf("mute_detail") = 0 Then _
                .Add "mute_detail", "2"

            If .IndexOf("vol_detail") = 0 Then _
                .Add "vol_detail", "2"


        End With

        .Save

    End With

    ' /* create handler window (for forthcoming Vista/Win7 notifier) */

    EZRegisterClass CLASS_NAME
    mhWnd = EZAddWindow(CLASS_NAME, Me, CLASS_NAME)

    ' /* register */

    mAppToken = sn41RegisterApp(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png")

    If mAppToken Then
        sn41AddClass mAppToken, CLASS_AUDIO_CHANGE, "Master volume"

    Else
        g_Debug "failed to register with Snarl (" & CStr(sn41GetLastError()) & ")", LEMON_LEVEL_CRITICAL

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub ISnarlExtension_Stop()

    ' /* unregister */

    sn41UnregisterApp mAppToken

    ' /* zap window */

    EZRemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Sub ISnarlExtension_TidyUp()

    If Not (mPanel Is Nothing) Then _
        mPanel.Quit

    Set InputMixer = Nothing

End Sub

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)

    g_Debug "Extension.ControlChanged()", LEMON_LEVEL_PROC_ENTER

    mSettings.Update Control.GetName, Value
    mConfigFile.Save

'    If Control.GetName = "MultiNotificationMode" Then _
        uSetClasses

    g_Debug "'" & Control.GetName & "' was set to '" & Value & "'", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

'    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()

    Case "go_website"
        ShellExecute mPanel.hWnd, "open", "http://www.fullphat.net/", vbNullString, vbNullString, 5

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

Private Function uGetIcon(ByVal Value As Long) As String
Dim sz As String

    Select Case Value
    Case 100
        sz = "!audio-volume-full"

    Case Is >= 50
        sz = "!audio-volume-mid"

    Case Is > 0
        sz = "!audio-volume-low"

    Case Else
        sz = "!audio-volume-off"

    End Select

    uGetIcon = sz

End Function

Private Function uAudioResourceOk() As Boolean

    On Error Resume Next

Dim c As Long

    Err.Clear
    c = audio_resource.count_devices()
    uAudioResourceOk = (Err.Number = 0)

End Function

'Private Sub VistaMixer_MuteChanged(ByVal Muted As Boolean)
'
'    If Muted Then
'        ' /* muted */
'        If snIsMessageVisible(mMuteId) Then _
'            snHideMessage mMuteId
'
'        mMuteId = snShowMessageEx(CLASS_AUDIO_MUTED, "Volume Muted", CStr(mVolPrevious), mSettings.NotificationTimeoutUNUSED, "!audio-volume-mute")
'
'    Else
'        ' /* unmuted */
'        If snIsMessageVisible(mMuteId) Then _
'            snHideMessage mMuteId
'
'        mMuteId = snShowMessageEx(CLASS_AUDIO_UNMUTED, "Volume Unmuted", CStr(mVolPrevious), mSettings.NotificationTimeoutUNUSED, uGetIcon(mVolPrevious))
'
'    End If
'
'End Sub
'
'Private Sub VistaMixer_VolChanged(ByVal Decibels As Single)
'
'    ' /* remove any muted/unmuted notification first */
'
'    If snIsMessageVisible(mMuteId) Then _
'        snHideMessage mMuteId
'
'    ' /* if there's an existing volume changed notification, update that */
'
'    If snIsMessageVisible(mNotificationId) Then
'        snUpdateMessage mNotificationId, "", CStr(Decibels), uGetIcon(Decibels)
'
'    Else
'        mNotificationId = snShowMessageEx(CLASS_AUDIO_LEVEL_CHANGED, "Volume Changed", CStr(Decibels), mSettings.NotificationTimeoutUNUSED, uGetIcon(Decibels))
'
'    End If
'
'End Sub

'Private Sub uSetClasses()
''Dim bMultiNotify As Boolean
'
''    sn41RemoveAllClasses mAppToken
'
''    bMultiNotify = (mSettings.GetValueWithDefault("MultiNotificationMode", "1") = "1")
''
''    If bMultiNotify Then
''        sn41AddClass mAppToken, CLASS_AUDIO_LEVEL_CHANGED, "Master volume changed"
''        sn41AddClass mAppToken, CLASS_AUDIO_MUTED, "Audio muted"
''        sn41AddClass mAppToken, CLASS_AUDIO_UNMUTED, "Audio unmuted"
''
''    Else
'        sn41AddClass mAppToken, CLASS_AUDIO_CHANGE, "Master volume "
''
''    End If
'
'End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

End Function

Private Sub uVolChanged(Control As audio_resource.BMixerLineControl, ByVal Value As Long, ByVal Detail As Long)

    ' /* bounds check */

    If Control.Maximum < 1 Then
        g_Debug "Extension.uVolChanged(): bad maximum setting '" & Control.Maximum & "'", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim i As Long

    ' /* has it changed enough to make a difference? */

    i = Fix((Value / Control.Maximum) * 100)
    If i = mVol Then _
        Exit Sub

    ' /* yes */

    mVol = i

    ' /* if volume is muted don't bother notifying */

    If current_master_mute() Then _
        Exit Sub

    ' /* notify */

Dim szTitle As String
Dim szText As String

    Select Case Detail
    Case 1
        ' /* icon only */

    Case 2
        If mSettings.GetValueWithDefault("MeterStyleMode", "0") = "1" Then
            szText = CStr(mVol)

        Else
            szTitle = CStr(mVol) & "%"

        End If

    Case 3
        szTitle = "Volume changed"

        If mSettings.GetValueWithDefault("MeterStyleMode", "0") = "1" Then
            szText = CStr(mVol)

        Else
            szText = "Current volume is " & CStr(mVol) & "%"

        End If

    End Select

    If sn41IsVisible(mCurrentId) Then
        sn41EZUpdate mCurrentId, szTitle, szText, , uGetIcon(mVol)

    Else
        mCurrentId = sn41EZNotify(mAppToken, CLASS_AUDIO_CHANGE, szTitle, szText, -1, uGetIcon(mVol))

    End If

End Sub

Private Sub uMuteChanged(ByVal Value As Long, ByVal Detail As Long)
Dim szTitle As String
Dim szText As String

    Select Case Detail
    Case 1
        ' /* icon only */

    Case 2
        If mSettings.GetValueWithDefault("MeterStyleMode", "0") = "1" Then
            szText = CStr(mVol)

        Else
            szTitle = IIf(Value = 1, "Muted", "Unmuted")

        End If

    Case 3
        szTitle = IIf(Value = 1, "Muted", "Unmuted")

        If mSettings.GetValueWithDefault("MeterStyleMode", "0") = "1" Then
            szText = CStr(mVol)

        Else
            szText = "Current volume is " & CStr(mVol) & "%"

        End If

    End Select

    Debug.Print "uMuteChanged(): title=" & szTitle & " text=" & szText

    If sn41IsVisible(mCurrentId) Then
        sn41EZUpdate mCurrentId, szTitle, szText, , IIf(Value = 1, "!audio-volume-mute", uGetIcon(mVol))

    Else
        mCurrentId = sn41EZNotify(mAppToken, CLASS_AUDIO_CHANGE, szTitle, szText, , IIf(Value = 1, "!audio-volume-mute", uGetIcon(mVol)))

    End If

End Sub
