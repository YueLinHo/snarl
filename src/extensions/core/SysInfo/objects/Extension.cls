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

Private Const LIB_DATE = "6-Oct-2010"
Private Const LIB_RELEASE = "1.1"
Private Const EXT_NAME = "SysInfo"
Private Const EXT_COMMENT = "Notifies about various system-wide events"

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

Dim mLastErr As String
Dim mhWndOwner As Long              ' // handle of the window which launched our config (typically Snarl's Prefs panel)

Private Type T_SETTINGS
    Token As Long
    CurrentRAM As Double
    NotifiedLowRAM As Boolean
    NotifiedNoRAM As Boolean

    DiskThreshold As Integer        ' // % free
    CPUThreshold As Integer         ' // % utilisation
    CPUHitDuration As Integer       ' // number of seconds
    CPUHitMonitoring As Long        ' // tick count when CPU first hit threshold
    TellMeWhenMonitoringStarts As Boolean
    TellMeWhenMonitoringStops As Boolean
    NotifiedHighCPU As Boolean

End Type

Dim mSettings As T_SETTINGS

Dim mVolsDone As BTagList
Dim mAdapterList As BTagList

Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1
Dim WithEvents theNICTimer As BTimer
Attribute theNICTimer.VB_VarHelpID = -1
Dim WithEvents theCPUTimer As BTimer
Attribute theCPUTimer.VB_VarHelpID = -1

Dim WithEvents theNetRoster As TNetRoster2
Attribute theNetRoster.VB_VarHelpID = -1
Dim WithEvents theDUNRoster As TDUNRoster
Attribute theDUNRoster.VB_VarHelpID = -1
Dim WithEvents theVolRoster As TSimpleVolRoster
Attribute theVolRoster.VB_VarHelpID = -1

'Implements MVersionInfo             ' // optional: allows the SDOS version command to query the object
Implements ISnarlExtension

Private Function ISnarlExtension_GetConfigWindow() As Long
End Function

Private Sub ISnarlExtension_GetInfo(Info As libSnarlExtn.extension_info)

    With Info
        .Author = "fullphat products"
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = EXT_COMMENT
        .Flags = SNARL_EXTN_WANTS_PULSE
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

    l3OpenLog "%APPDATA%\" & EXT_NAME & ".log", True
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

End Function

Private Sub ISnarlExtension_LastError(Description As String)

    Description = mLastErr

End Sub

Private Sub ISnarlExtension_Pulse()
End Sub

Private Sub ISnarlExtension_Start()
Dim pConfig As CConfFile
Dim szPath As String
Dim i As Integer

    On Error Resume Next

    With mSettings
        .CurrentRAM = 0
        .NotifiedLowRAM = False
        .NotifiedNoRAM = False
        .Token = 0

        .CPUThreshold = 95          ' // 95%
        .CPUHitDuration = 30        ' // for 30 seconds...
        .DiskThreshold = 90         ' // 90% full

        ' /* does snarl's /etc/ path exist? */

        If snext_getuserpath(szPath) Then

            Set pConfig = New CConfFile
            If pConfig.SetTo(szPath & "etc\sysinfo.config", True) Then
                .CPUThreshold = Val(pConfig.ValueOf("CPUThreshold"))
                .CPUHitDuration = Val(pConfig.ValueOf("CPUHitDuration"))
                .DiskThreshold = Val(pConfig.ValueOf("DiskThreshold"))
                .TellMeWhenMonitoringStarts = (pConfig.ValueOf("TellMeWhenMonitoringStarts") = "1")
                .TellMeWhenMonitoringStops = (pConfig.ValueOf("TellMeWhenMonitoringStops") = "1")

            End If

        End If

'        MsgBox "cpu: " & .CPUThreshold & " for " & .CPUHitDuration

        ' /* register */

        .Token = sn41RegisterApp("application/vnd.k23-" & LCase$(EXT_NAME), EXT_NAME, g_MakePath(App.Path) & "icon.png")
        If .Token = 0 Then _
            Exit Sub

        sn41AddClass .Token, CLASS_LOW_MEM, "Physical memory low", False
        sn41AddClass .Token, CLASS_NO_MEM, "Physical memory exhausted", False
        sn41AddClass .Token, CLASS_HIGH_CPU, "High processor utilisation"
        sn41AddClass .Token, CLASS_LOW_DISK, "Low disk space", False
        sn41AddClass .Token, CLASS_NIC_CONN, "Network connected", False
        sn41AddClass .Token, CLASS_NIC_DISC, "Network disconnected", False
        sn41AddClass .Token, CLASS_DUN_CONN, "Dial-up connected", False
        sn41AddClass .Token, CLASS_DUN_DISC, "Dial-up disconnected", False
        sn41AddClass .Token, CLASS_VOL_MNT, "Volume mounted"
        sn41AddClass .Token, CLASS_VOL_UMNT, "Volume unmounted"

    End With

    Set mAdapterList = new_BTagList()

    Set theTimer = new_BTimer(5000)
    Set theNICTimer = new_BTimer(2000)
    Set theCPUTimer = new_BTimer(1000)

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

    Set theVolRoster = New TSimpleVolRoster
    With theVolRoster
        .LockRoster
        .Refresh
        .UnlockRoster

    End With

    ' /* call once to get baseline */

    krnl_GetSystemInfo
    krnl_GetCurrentTotalCPUPercent

End Sub

Private Sub ISnarlExtension_Stop()

    Set theTimer = Nothing
    Set theNICTimer = Nothing
    Set theCPUTimer = Nothing

    Set theDUNRoster = Nothing
    Set theNetRoster = Nothing

    sn41UnregisterApp mSettings.Token
    mSettings.Token = 0

End Sub

Private Sub ISnarlExtension_TidyUp()
End Sub

Private Sub theCPUTimer_Pulse()
Dim dCurrentCPU As Double
Dim t As Long
Dim l As Long

    dCurrentCPU = krnl_GetCurrentTotalCPUPercent()

    If dCurrentCPU > mSettings.CPUThreshold Then
        ' /* over the threshold */

        If mSettings.CPUHitMonitoring = 0 Then
            ' /* first time (not tracking) */
            If mSettings.TellMeWhenMonitoringStarts Then _
                l = sn41EZNotify(mSettings.Token, CLASS_HIGH_CPU, "High processor utilisation", "Started monitoring" & vbCrLf & _
                                 "Current utilization: " & CStr(dCurrentCPU) & "%", , "!system-info")

            mSettings.CPUHitMonitoring = GetTickCount()

        ElseIf Not mSettings.NotifiedHighCPU Then
            ' /* already tracking */

            t = GetTickCount() - mSettings.CPUHitMonitoring
            If t > (mSettings.CPUHitDuration * 1000) Then
                ' /* and past the cut-off duration */
                l = sn41EZNotify(mSettings.Token, CLASS_HIGH_CPU, "High processor utilisation", _
                                 "Total processor utilization has been over " & mSettings.CPUThreshold & "% for " & CStr(t \ 1000) & " seconds", , "!system-info")

                mSettings.NotifiedHighCPU = True

            End If

        End If

    ElseIf mSettings.CPUHitMonitoring > 0 Then
        ' /* we were monitoring so stop */
        If mSettings.TellMeWhenMonitoringStops Then _
            l = sn41EZNotify(mSettings.Token, CLASS_HIGH_CPU, "High processor utilisation", "Stopped monitoring" & vbCrLf & _
                             "Current utilization: " & CStr(dCurrentCPU) & "%", , "!system-info")

        mSettings.CPUHitMonitoring = 0
        mSettings.NotifiedHighCPU = False

    End If

End Sub

Private Sub theDUNRoster_Connected(ByVal Name As String, Details As melon4.BTagList)
Dim i As Long

    i = sn41EZNotify(mSettings.Token, CLASS_DUN_CONN, _
                     g_Capitalize(Details.ValueOf("device_type", "Dial-up"), False) & " connected", _
                     Name & " is now connected at " & _
                     uSpeed(Val(Details.ValueOf("speed", "0"))) & _
                     " using " & Details.ValueOf("device_name"), , "!dun-connected")

End Sub

Private Sub theDUNRoster_Disconnected(ByVal Name As String, Details As melon4.BTagList)
Dim i As Long

    i = sn41EZNotify(mSettings.Token, CLASS_DUN_DISC, _
                     g_Capitalize(Details.ValueOf("device_type", "Dial-up"), False) & " disconnected", _
                     Name & " is disconnected", , "!dun-disconnected")

End Sub

'Private Sub theNetRoster_Connected(Details As melon4.BTagList)
'Static i As Long
'
'    If Details.ValueOf("active") = "1" Then _
'        i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network connected", _
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
'        i = sn41EZNotify(mSettings.Token, CLASS_NIC_DISC, "Network disconnected", _
'                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
'
'End Sub
'
'Private Sub theNetRoster_StatusChanged(Details As melon4.BTagList, ByVal PreviousState As Long, ByVal NewState As Long)
'Static i As Long
'
'    If (PreviousState = IF_OPER_STATUS_OPERATIONAL) And (NewState = IF_OPER_STATUS_NON_OPERATIONAL) Then
'        ' /* gone from connected to disconnected */
'        i = sn41EZNotify(mSettings.Token, CLASS_NIC_DISC, "Network disconnected", _
'                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
'
'    ElseIf (PreviousState = IF_OPER_STATUS_NON_OPERATIONAL) And (NewState = IF_OPER_STATUS_OPERATIONAL) Then
'        ' /* gone from disconnected to connected */
'        i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network connected", _
'                         Details.ValueOf("connection") & " is now connected at " & _
'                         uSpeed(Val(Details.ValueOf("speed", "0"))) & _
'                         " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")
'
'    Else
'        i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network status changed", _
'                         Details.ValueOf("connection") & " status changed from " & CStr(PreviousState) & " to " & CStr(NewState), , "")
'
'    End If
'
'End Sub

Private Sub theNetRoster_Connected(ByVal Name As String, Device As melon4.BTagList)
Dim sz As String
Dim i As Long

'    sz = Device.ValueOf("name")

    If mSettings.Token Then _
        i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network connected", _
                         Name & " is now connected" & vbCrLf & _
                         "Speed: " & Device.ValueOf("formatted_speed", "0") & _
                         IIf(sz <> "", vbCrLf & "Device: " & sz, ""), , "!network-connected")

'        sn41EZNotify mToken, "", "Connected", _
'                     Name & vbCrLf & _
'                     "Speed: " & Details.ValueOf("formatted_speed") & vbCrLf & _
'                     "Type: " & Details.ValueOf("type") & vbCrLf & _
'                     "Media: " & Details.ValueOf("media_type"), , _
'                     "!network-connected"

End Sub

Private Sub theNetRoster_Disconnected(ByVal Name As String, ByVal DeviceType As Long)
Dim i As Long

    If mSettings.Token Then _
        i = sn41EZNotify(mSettings.Token, CLASS_NIC_DISC, "Network disconnected", _
                         Name & " is disconnected", , "!network-disconnected")

End Sub

Private Sub theNICTimer_Pulse()

    theNetRoster.Refresh
    theDUNRoster.Refresh
    theVolRoster.Refresh

    uCheckDiskSpace

End Sub

Private Sub theTimer_Pulse()
Const WARNING_LIMIT = 20
Const ZERO_LIMIT = 0

Dim l As Long

    ' /* check RAM */

    If mSettings.NotifiedLowRAM Then
        ' /* see if it's back above the warning limit */
        If g_PercentMemInUse(True) > WARNING_LIMIT Then _
            mSettings.NotifiedLowRAM = False

    Else
        ' /* is it at or below the warning limit */
        If g_PercentMemInUse(True) <= WARNING_LIMIT Then
            l = sn41EZNotify(mSettings.Token, CLASS_LOW_MEM, "Physical memory low", _
                             "There is " & g_FileSizeToStringEx2(g_GetPhysMem(False)) & " of " & g_FileSizeToStringEx2(g_GetPhysMem(True)) & " physical memory left.", , "!system-info")

            mSettings.NotifiedLowRAM = True

        End If
    End If

    If mSettings.NotifiedNoRAM Then
        ' /* see if it's back in the black... */
        If g_PercentMemInUse(True) > ZERO_LIMIT Then _
            mSettings.NotifiedNoRAM = False

    Else
        ' /* is it at or below zero */
        If g_PercentMemInUse() <= ZERO_LIMIT Then
            l = sn41EZNotify(mSettings.Token, CLASS_NO_MEM, "Physical memory exhausted", _
                             "There is no physical memory available - the operating system will begin to move memory to and from disk (paging) which will make the computer respond more slowly." & vbCrLf & vbCrLf & _
                             "Closing unwanted applications, browser tabs or system tray applications will restore physical memory.", , "!system-warning")

            mSettings.NotifiedNoRAM = True

        End If
    End If

End Sub

Private Sub uCheckDiskSpace()
'Const DISK_LIMIT = 90
Dim szPath As String
Dim szLabel As String
Dim lType As Long
Dim l As Long

    With theVolRoster
        .Rewind

        Do While .GetNextVolume(szPath, szLabel, lType)
            If (lType = DRIVE_FIXED) Or (lType = DRIVE_REMOVABLE) Then
                With New TSimpleVolume
                    If .SetTo(szPath) Then
                        If (mVolsDone.IndexOf(szPath) = 0) And (.PercentUsed >= mSettings.DiskThreshold) Then

                            l = sn41EZNotify(mSettings.Token, CLASS_LOW_DISK, uLabel(szLabel, szPath) & " is nearly full", _
                                             g_FileSizeToStringEx2(.Available) & " of " & g_FileSizeToStringEx2(.Capacity) & " remaining", , "!system-info")

                            mVolsDone.Add new_BTagItem(szPath, "")

                        ElseIf mVolsDone.IndexOf(szPath) Then
                            ' /* have we dropped below the limit? */
                            If .PercentUsed < mSettings.DiskThreshold Then _
                                mVolsDone.Remove mVolsDone.IndexOf(szPath)

                        End If

                    Else
                        g_Debug "Extension.uCheckDiskSpace(): can't create CVolume for '" & szPath & "'", LEMON_LEVEL_CRITICAL

                    End If
                End With
            End If
        Loop
    End With

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

Private Sub theVolRoster_VolumeMounted(ByVal Path As String, ByVal Label As String, ByVal VolType As Long)
Dim i As Long

    i = sn41EZNotify(mSettings.Token, CLASS_VOL_MNT, _
                     "Voume mounted", _
                     uLabel(Label, Path) & uCapacityInfo(Path), , uIcon(Path), , Path)

End Sub

Private Sub theVolRoster_VolumeUnmounted(ByVal Path As String, ByVal Label As String, ByVal VolType As Long)
Dim i As Long

    i = sn41EZNotify(mSettings.Token, CLASS_VOL_UMNT, _
                     "Volume unmounted", _
                     uLabel(Label, Path), , "!vol-unmounted")

End Sub

Private Function uLabel(ByVal Label As String, Optional ByVal Path As String, Optional ByVal AddPath As Boolean = True) As String

    uLabel = IIf(Label = "", "<untitled>", Label)

    If AddPath Then _
        uLabel = uLabel & " (" & Path & ")"

End Function

Private Function uCapacityInfo(ByVal Path As String) As String
Dim pv As TSimpleVolume

    Set pv = New TSimpleVolume
    If pv.SetTo(Path) Then _
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
