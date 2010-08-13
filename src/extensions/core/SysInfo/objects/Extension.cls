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

Private Const LIB_DATE = "12-Aug-2010"
Private Const LIB_RELEASE = "1.02"
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

    LastCheckCPU As Long
    NotifiedHighCPU As Boolean
    'CPU() As BCPU

End Type

Dim mVolRoster As TVolumeRoster
Dim mVolsDone As BTagList

Dim mSettings As T_SETTINGS

Private Type T_ADAPTER
    pif As MIB_IFROW
    pai As IP_ADAPTER_INFO
    bHasAdapterInfo As Boolean
'    AdapterName As String
'    ConnectionName As String

End Type

Dim mIFBufferLen As Long        ' // length of MIB_IFROW buffer
Dim mIFBuffer() As Byte         ' // MIB_IFROW buffer
Dim mIFCount As Long            ' // number of MIB_IFROWs in buffer

Dim mAdapterList As BTagList

Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1
Dim WithEvents theNICTimer As BTimer
Attribute theNICTimer.VB_VarHelpID = -1

Dim WithEvents theNetRoster As TNetRoster
Attribute theNetRoster.VB_VarHelpID = -1
Dim WithEvents theDUNRoster As TDUNRoster
Attribute theDUNRoster.VB_VarHelpID = -1
Dim WithEvents theVolRoster As TVolumeRoster
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

    l3OpenLog "%APPDATA%\" & EXT_NAME & ".log"
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
Dim i As Integer

    With mSettings
        .CurrentRAM = 0
        .NotifiedLowRAM = False
        .NotifiedNoRAM = False
        .Token = 0

        ' /* register */

        .Token = sn41RegisterApp("application/vnd.k23-" & LCase$(EXT_NAME), EXT_NAME, g_MakePath(App.Path) & "icon.png")
        If .Token = 0 Then _
            Exit Sub

        sn41AddClass .Token, CLASS_LOW_MEM, "Physical memory low", False
        sn41AddClass .Token, CLASS_NO_MEM, "Physical memory exhausted", False
'        sn41AddClass .Token, CLASS_HIGH_CPU, "High processor utilisation"
        sn41AddClass .Token, CLASS_LOW_DISK, "Low disk space", False
'        sn41AddClass .Token, CLASS_NIC_CONN, "Network connected", False
'        sn41AddClass .Token, CLASS_NIC_DISC, "Network disconnected", False
        sn41AddClass .Token, CLASS_DUN_CONN, "Dial-up connected", False
        sn41AddClass .Token, CLASS_DUN_DISC, "Dial-up disconnected", False
        sn41AddClass .Token, CLASS_VOL_MNT, "Volume mounted"
        sn41AddClass .Token, CLASS_VOL_UMNT, "Volume unmounted"

'        If processor_count Then
'            ReDim .CPU(processor_count - 1)
'            For i = 0 To processor_count - 1
'                Set .CPU(i) = New BCPU
'                .CPU(i).SetTo i
'                .CPU(i).OpenQuery i
'
'            Next i
'        End If

    End With

    Set mVolRoster = New TVolumeRoster
'    mVolRoster.ExcludeFloppies = True
    Set mVolsDone = new_BTagList()

    Set mAdapterList = new_BTagList()

    Set theTimer = new_BTimer(5000)
    Set theNICTimer = new_BTimer(2000)

    Set theNetRoster = New TNetRoster
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

    Set theVolRoster = New TVolumeRoster
    With theVolRoster
        .LockRoster
        .Refresh
        .UnlockRoster

    End With

End Sub

Private Sub ISnarlExtension_Stop()

    Set theTimer = Nothing
    Set theNICTimer = Nothing

    Set theDUNRoster = Nothing
    Set theNetRoster = Nothing

    sn41UnregisterApp mSettings.Token
    mSettings.Token = 0

End Sub

Private Sub ISnarlExtension_TidyUp()
End Sub

Private Function uGetAvgCPU() As Double

'    If processor_count < 1 Then _
'        Exit Function
'
'Dim i As Long
'Dim c As Double
'
'    For i = 0 To processor_count - 1
'        c = c + mSettings.CPU(i).CurrentLoad
'
'    Next i

'    uGetAvgCPU = c '/ processor_count

End Function

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

Private Sub theNetRoster_Connected(Details As melon4.BTagList)
'Static i As Long
'
'    i = Val(Details.ValueOf("status"))
'    If i <> IF_OPER_STATUS_OPERATIONAL Then _
'        Exit Sub
'
'    i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network connected", _
'                     Details.ValueOf("connection") & " is now connected at " & _
'                     uSpeed(Val(Details.ValueOf("speed", "0"))) & _
'                     " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")

End Sub

Private Sub theNetRoster_Disconnected(Details As melon4.BTagList)

'    ' /* was the adapter active at the time it disappeared? */
'
'    If Val(Details.ValueOf("status", "0")) <> IF_OPER_STATUS_OPERATIONAL Then _
'        Exit Sub
'
'Static i As Long
'
'    i = sn41EZNotify(mSettings.Token, CLASS_NIC_DISC, "Network disconnected", _
'                     Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")

End Sub

Private Sub theNetRoster_StatusChanged(Details As melon4.BTagList, ByVal PreviousState As Long, ByVal NewState As Long)
Static i As Long

    If (PreviousState = IF_OPER_STATUS_OPERATIONAL) And (NewState = IF_OPER_STATUS_NON_OPERATIONAL) Then
        ' /* gone from connected to disconnected */
        i = sn41EZNotify(mSettings.Token, CLASS_NIC_DISC, "Network disconnected", _
                         Details.ValueOf("connection") & " is disconnected", , "!network-disconnected")
    
    ElseIf (PreviousState = IF_OPER_STATUS_NON_OPERATIONAL) And (NewState = IF_OPER_STATUS_OPERATIONAL) Then
        ' /* gone from disconnected to connected */
        i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network connected", _
                         Details.ValueOf("connection") & " is now connected at " & _
                         uSpeed(Val(Details.ValueOf("speed", "0"))) & _
                         " using " & Details.ValueOf("name", "<unknown adapter>"), , "!network-connected")

    Else
        i = sn41EZNotify(mSettings.Token, CLASS_NIC_CONN, "Network status changed", _
                         Details.ValueOf("connection") & " status changed from " & CStr(PreviousState) & " to " & CStr(NewState), , "")

    End If

End Sub

Private Sub theNICTimer_Pulse()

    theNetRoster.Refresh
    theDUNRoster.Refresh
    theVolRoster.Refresh

End Sub

Private Sub theTimer_Pulse()
Const WARNING_LIMIT = 20
Const ZERO_LIMIT = 0
Const CPU_LIMIT = 1

Dim n As Double
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



    ' /* CPU */

'    n = uGetAvgCPU()
'    l = sn41EZNotify(mSettings.Token, CLASS_HIGH_CPU, "High processor utilisation", CStr(n) & " / " & CStr(processor_count), , "!system-info")

    ' /* Disk */

Const DISK_LIMIT = 90
Dim pDisk As CVolume

'    With mVolRoster
'        .Refresh False, False
'        .Rewind
'
'        Do While .GetNextVolume(pDisk)
'            With pDisk
'                If (.VolumeType And &HF) = DRIVE_FIXED Then
'                    If (mVolsDone.IndexOf(.Path) = 0) And (.PercentUsed >= DISK_LIMIT) Then
'                        l = sn41EZNotify(mSettings.Token, CLASS_LOW_DISK, "The " & g_SafeLeftStr(.Path, 2) & " drive is nearly full", _
'                                         g_FileSizeToStringEx2(.Available) & " of " & g_FileSizeToStringEx2(.Capacity) & " remaining", , "!system-info")
'
'                        mVolsDone.Add new_BTagItem(.Path, "")
'
'                    ElseIf mVolsDone.IndexOf(.Path) Then
'                        ' /* have we dropped below the limit? */
'                        If .PercentUsed < DISK_LIMIT Then _
'                            mVolsDone.Remove mVolsDone.IndexOf(.Path)
'
'                    End If
'                End If
'            End With
'        Loop
'
'    End With

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

Private Function uIsConnected(ByVal OperStatus As Long) As String

    Select Case OperStatus
    Case IF_OPER_STATUS_OPERATIONAL
        uIsConnected = "1"

    Case Else
        uIsConnected = "0"

    End Select

End Function

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

Private Sub theVolRoster_VolumeMounted(Volume As CVolume)
Dim i As Long

    i = sn41EZNotify(mSettings.Token, CLASS_VOL_MNT, _
                     "Voume mounted", _
                     IIf(Volume.Label = "", "<untitled>", Volume.Label) & " (" & Volume.Path & ")", , "!vol-mounted", , Volume.Path)

End Sub

Private Sub theVolRoster_VoulmeUnmounted(Volume As CVolume)
Dim i As Long

    i = sn41EZNotify(mSettings.Token, CLASS_VOL_UMNT, _
                     "Volume unmounted", _
                     IIf(Volume.Label = "", "<untitled>", Volume.Label) & " (" & Volume.Path & ")", , "!vol-unmounted")

End Sub
