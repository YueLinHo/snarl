VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snarl"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UTF8 support by Tomas, Icons by Mattahan.  Additional development and testing:   "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   1980
      TabIndex        =   6
      Top             =   1500
      Width           =   3555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1042
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Index           =   4
      Left            =   1980
      TabIndex        =   5
      Top             =   1980
      Width           =   4395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.fullphat.net"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008A504A&
      Height          =   255
      Left            =   2370
      TabIndex        =   4
      Top             =   3180
      Width           =   1710
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   5580
      Picture         =   "frmAbout.frx":10E9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   430
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   60
      Picture         =   "frmAbout.frx":1D4F
      Top             =   180
      Width           =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Snarl"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1980
      TabIndex        =   3
      Top             =   180
      Width           =   2955
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A Notification System for Windows"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005-2010 full phat products"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1980
      TabIndex        =   1
      Top             =   1200
      Width           =   4155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Inspired by Growl for Macintosh"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1980
      TabIndex        =   0
      Top             =   840
      Width           =   3435
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   3000
      Left            =   0
      Top             =   0
      Width           =   6450
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           frmAbout.frm
'/
'/  Description:    Displays the product info and handles various other tasks
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As Any, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long

Private Declare Function WTSRegisterSessionNotification Lib "Wtsapi32" (ByVal hWnd As Long, ByVal THISSESS As Long) As Long
Private Declare Function WTSUnRegisterSessionNotification Lib "Wtsapi32" (ByVal hWnd As Long) As Long

Private Const NOTIFY_FOR_ALL_SESSIONS As Long = 1

Private Const WM_WTSSESSION_CHANGE As Long = &H2B1
Private Const WTS_SESSION_LOCK As Long = 7
Private Const WTS_SESSION_UNLOCK As Long = 8

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETSCREENSAVERRUNNING = &H72


Dim mSysKeyPrefs As Long
Dim mSysKeyTest As Long

Dim mTrayIcon As BNotifyIcon

Dim m_About As String
'Dim mTaskbarCreated As Long
Dim m_SelectedApp As String         ' // current selected application in listbox

Dim mPrefs As T_CONFIG
Dim mCurAlert As TAlert

Dim mPanel As BPrefsPanel
Dim mAppsPage As TAppsPage

    ' /* listening sockets */
Dim WithEvents JSONSocket As CSocket            ' // 9889 (TCP)
Attribute JSONSocket.VB_VarHelpID = -1
Dim WithEvents GrowlUDPSocket As CSocket        ' // 9887 (UDP)
Attribute GrowlUDPSocket.VB_VarHelpID = -1
Dim mListener() As CSnarlListener               ' // 9887 (TCP) - all local ip addresses
Dim mListenerCount As Long

    ' /* active JSON connections */
Dim mJSONSocket() As CJSONSocket
Dim mSockets As Long

Dim mClickThruOver As CSnarlWindow
Dim mMenuOpen As Boolean
Dim mDownloadId As Long
Dim WithEvents theReadyTimer As BTimer
Attribute theReadyTimer.VB_VarHelpID = -1

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long

End Type

Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO) As Boolean

Dim WithEvents theIdleTimer As BTimer
Attribute theIdleTimer.VB_VarHelpID = -1

Implements MMessageSink
Implements KPrefsPanel
Implements KPrefsPage
Implements MWndProcSink

Private Sub Form_Load()
Dim sz As String
Dim pm As OMMenu
Dim n As Integer

    On Error Resume Next

    g_HideFromView Me.hWnd

    ' /* R2.4 DR7: check for Calibri and default to Tahoma */

    For n = Label3.LBound To Label3.UBound
        Label3(n).Font.Name = "Calibri"
        If Label3(n).Font.Name <> "Calibri" Then
            Label3(n).Font.Name = "Tahoma"
            Label3(n).Font.Size = Label3(n).Font.Size - 1

        End If

    Next n

    ' /* add user messages to UIPI allowed filter */

'    g_Debug "frmAbout.Load(): relaxing UIPI message filter..."
'    g_Debug "frmAbout.Load(): o/s is 0x" & g_HexStr(g_GetNTVersion())
'
'    If g_GetNTVersion() >= NTWIN7 Then
'        g_Debug "Windows 7 / Windows 2008 R2 (or better)..."
'        ChangeWindowMessageFilterEx Me.hWnd, WM_SNARL_TRAY_ICON, MSGFLT_ALLOW, 0&
'        ChangeWindowMessageFilterEx Me.hWnd, WM_REMOTENOTIFY, MSGFLT_ALLOW, 0&
'        ChangeWindowMessageFilterEx Me.hWnd, WM_INSTALL_SNARL, MSGFLT_ALLOW, 0&
'        ChangeWindowMessageFilterEx Me.hWnd, MSG_SHOW_PREFS, MSGFLT_ALLOW, 0&
''        ChangeWindowMessageFilterEx Me.hWnd, MSG_QUIT, MSGFLT_ALLOW, 0&
'        ChangeWindowMessageFilterEx Me.hWnd, WM_SNARLTEST, MSGFLT_ALLOW, 0&
'        ChangeWindowMessageFilterEx Me.hWnd, WM_MANAGE_SNARL, MSGFLT_ALLOW, 0&
'
''WM_SNARL_TRAY_ICON     ' // WM_USER + 3
''WM_REMOTENOTIFY        ' // WM_USER + 9
''WM_INSTALL_SNARL       ' // WM_USER + 12
''MSG_SHOW_PREFS         ' // WM_USER + 80
''MSG_QUIT               ' // WM_USER + 81
''WM_SNARLTEST           ' // WM_USER + 237
''WM_MANAGE_SNARL        ' // WM_USER + 238
'
'    ElseIf g_GetNTVersion() = NTVISTA Then
'        ' /* we will do this here but *NOT* in TMainWindow construction */
'        g_Debug "Windows Vista / Windows 2008"
'        ChangeWindowMessageFilter WM_SNARL_TRAY_ICON, MSGFLT_ADD
'        ChangeWindowMessageFilter WM_REMOTENOTIFY, MSGFLT_ADD
'        ChangeWindowMessageFilter WM_INSTALL_SNARL, MSGFLT_ADD
'        ChangeWindowMessageFilter MSG_SHOW_PREFS, MSGFLT_ADD
''        ChangeWindowMessageFilter MSG_QUIT, MSGFLT_ADD
'        ChangeWindowMessageFilter WM_SNARLTEST, MSGFLT_ADD
'        ChangeWindowMessageFilter WM_MANAGE_SNARL, MSGFLT_ADD
'
'    End If





    ' /* R2.4 DR8: register for TS session events */

    WTSRegisterSessionNotification Me.hWnd, NOTIFY_FOR_ALL_SESSIONS




    ' /* register the hotkeys */

    Me.bSetHotkeys

    ' /* pre-load our 'About' text */

    g_Debug "_load: pre-loading readme..."
    n = FreeFile()
    err.Clear
    Open g_MakePath(App.Path) & "read-me.rtf" For Input As #n
    If err.Number = 0 Then
        Do While Not EOF(n)
            Line Input #n, sz
            m_About = m_About & sz & vbCrLf

        Loop
        Close #n
    End If



    AddSubClass Me.hWnd, Me

    ' /* create the tray menus */

'    Set myIconMenu = New OMMenu

'    With myIconMenu
'        .AddItem .CreateItem("sticky", "Sticky Notifications")
'        .AddSeparator
'
'        .AddItem .CreateItem("dnd", "Do Not Disturb", , , gPrefs.do_not_disturb)
'        .AddItem .CreateItem("missed", "Missed Notifications...")
'
'        .AddSeparator
'        .AddItem .CreateItem("restart", "Restart Snarl")
'        .AddItem .CreateItem("stop", "Stop Snarl")
'        .AddItem .CreateItem("quit", "Quit Snarl")
'        .AddSeparator
'        .AddItem .CreateItem("prefs", "Settings...")
'        .AddItem .CreateItem("app_mgr", "App Manager...")
'        .AddSeparator
'        .AddItem .CreateItem("about", "About Snarl")
'
'    End With

    ' /* create the tray icon */

'    mTaskbarCreated = RegisterWindowMessage("TaskbarCreated")
'    g_Debug "_load: 'TaskbarCreated' = " & g_HexStr(mTaskbarCreated, 4)

    Set mTrayIcon = New BNotifyIcon
    mTrayIcon.SetTo Me.hWnd, WM_SNARL_TRAY_ICON
    uAddTrayIcon

    ' /* create our JSON listener */

    If g_ConfigGet("listen_for_json") = "1" Then _
        EnableJSON True

    ' /* create our Snarl listener */

    If g_ConfigGet("listen_for_snarl") = "1" Then _
        EnableSNP True


    Label3(0).Caption = "Snarl " & g_Version()
    g_Debug "_load: Version = " & g_Version(), LEMON_LEVEL_INFO

    ' /* create the idle input timer */

    Set theIdleTimer = New BTimer
    theIdleTimer.SetTo 250

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then
        Me.Hide
        Cancel = -1             ' // close gadget

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

    g_Debug "_Unload()", LEMON_LEVEL_PROC
    Me.Hide

    ' /* R2.4 DR8: unregister session events */

    WTSUnRegisterSessionNotification Me.hWnd

    ' /* stop the idle timer */

    Set theIdleTimer = Nothing

    ' /* close our Snarl listeners */

    If g_ConfigGet("listen_for_snarl") = "1" Then _
        EnableSNP False

    ' /* close our JSON listener */

    If g_ConfigGet("listen_for_json") = "1" Then _
        EnableJSON False



    If Not (mPanel Is Nothing) Then
        g_Debug "_Unload(): closing prefs window..."
        mPanel.Quit
        Set mPanel = Nothing
        g_Debug "_Unload(): prefs window closed"

    End If

    Set mTrayIcon = Nothing

'    g_Quit

    g_Debug "_Unload(): unsubclassing window..."
    RemoveSubClass Me.hWnd


    uUnregisterHotkeys


'    lemonUnregister

End Sub

Private Sub GrowlUDPSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim b() As Byte

    Debug.Print bytesTotal
    GrowlUDPSocket.GetData b(), vbArray + vbByte
    g_ProcessGrowlUDP b(), bytesTotal, GrowlUDPSocket.RemoteHost

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub Label2_Click()

End Sub

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim sz() As String

    Select Case Control.GetName

    Case "use_hotkey"
        ' /* R2.2: we have a separate config entry now */
        g_ConfigSet Control.GetName, Value
        prefskit_SafeEnable Control.Page.Panel, "hotkey_prefs", (Value = "1")
        frmAbout.bSetHotkeys

    Case "hotkey_prefs"
        ' /* the key picker control should return a pair of values separated by a comma.  The first
        '    value is the set of modifiers; the second value is the keycode of the key pressed */

        sz() = Split(Value, ",")
        If UBound(sz()) <> 1 Then _
            Exit Sub

        ' /* we're only interested in the keycode here */

        If sz(1) <> g_ConfigGet("hotkey_prefs") Then
            If frmAbout.bSetHotkeys(Val(sz(1))) Then
                ' /* registered okay so store the new keycode */
                g_ConfigSet "hotkey_prefs", sz(1)
                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: hotkey changed to #" & sz(1)

            Else
                g_Debug "TGeneralPage.ControlChanged(): [hotkey_prefs]: couldn't set hotkey to #" & sz(1), LEMON_LEVEL_WARNING
                sz(1) = g_ConfigGet("hotkey_prefs")

            End If

        End If

        Control.SetValue CStr(MOD_CONTROL) & "," & sz(1)


    Case "idle_minutes"

        Select Case Val(Value)
        Case 0
            Control.SetText "Never"

        Case 1
            Control.SetText "1 min"

        Case Else
            Control.SetText Value & " mins"

        End Select

        g_ConfigSet Control.GetName, Value


    Case "away_style", "busy_style"
        ' /* away and busy forwarding style */
        g_ConfigSet Control.GetName, LCase$(Replace$(prefskit_GetItem(Control), ": ", "/"))


    Case Else
        ' /* other controls */
        g_ConfigSet Control.GetName, Value

        If Control.GetName = "run_on_logon" Then
            g_SetAutoRun2

        ElseIf Control.GetName = "away_mode" Then
            prefskit_SafeEnable Control.Page.Panel, "away_style", (Value = "6")

        ElseIf Control.GetName = "busy_mode" Then
            prefskit_SafeEnable Control.Page.Panel, "busy_style", (Value = "6")
                
        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim hWnd As Long

    Select Case Control.GetName()

    Case "update_now"
        g_DoManualUpdateCheck

    Case "go_lemon"
        ShellExecute 0, "open", "notepad.exe", l3LogPath(), vbNullString, SW_SHOW

    Case "open_config"
        ShellExecute 0, "open", g_GetPath(g_SettingsPath()), vbNullString, vbNullString, SW_SHOW

    Case "cycle_config"
        g_ConfigInit

    Case "go_garbage"
        If g_IsWinXPOrBetter() Then _
            CoFreeUnusedLibrariesEx 0, 0

    Case "test"
        ' /* mimic how we would do it from an external app... */
        hWnd = FindWindow("w>Snarl", "Snarl")
        If IsWindow(hWnd) <> 0 Then _
            SendMessage hWnd, WM_SNARLTEST, 0, ByVal 0&

    Case "go_tray_icon"
        uAddTrayIcon


'    Case "restart_style_roster"
'        If Not (g_StyleRoster Is Nothing) Then
'            melonLibClose g_StyleRoster
'            MsgBox "Click OK when you're ready for the Style Roster to start up", vbInformation Or vbOKOnly, App.Title
'            melonLibOpen g_StyleRoster
'
'        End If

    End Select

'    Case "go_app_manager"
'        ShellExecute 0, "open", g_MakePath(App.Path) & "SNARLAPP_Manager.exe", vbNullString, vbNullString, SW_SHOW
'
'    Case "restart_style_roster"
'        If Not (g_StyleRoster Is Nothing) Then
'            melonLibClose g_StyleRoster
'            Sleep 500
''            MsgBox "Click OK when you're ready for the Style Roster to start up", vbInformation Or vbOKOnly, App.Title
'            melonLibOpen g_StyleRoster
'
'            If mPage.Panel.Find("installed_styles", pc) Then _
'                pc.Notify "update_list", Nothing
'
'
'        End If
'
''    Case "dnd_settings"
''        With New TStyleEnginePanel
''            .Go mPage.Panel.hWnd
''
''        End With
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static fIgnoreNext As Boolean
Dim dw As Long

    Select Case uMsg

    Case WM_COPYDATA, WM_MANAGE_SNARL, WM_SNARLTEST
        ' /* backwards compatability - send these directly to our handling window */
        ReturnValue = SendMessage(ghWndMain, uMsg, wParam, ByVal lParam)
        MWndProcSink_WndProc = True


    Case WM_HOTKEY
        If LoWord(wParam) = mSysKeyPrefs Then
'            sosOutput "ISubClassed.WndProc(): Old config hotkey pressed!", LEMON_LEVEL_WARNING
            Me.NewDoPrefs

        ElseIf LoWord(wParam) = mSysKeyTest Then
            uDoSysInfoNotification

        Else
            g_Debug "ISubClassed.WndProc(): Spurious WM_HOTKEY received: " & _
                    g_HexStr(HiWord(wParam), 4) & " " & g_HexStr(LoWord(wParam), 4), LEMON_LEVEL_WARNING

        End If

        MWndProcSink_WndProc = True


    Case MSG_SHOW_PREFS
        ' /* this message shouldn't arrive here anymore, being directed to TMainWindow instead */
        Me.NewDoPrefs


    Case MSG_QUIT, WM_CLOSE
        PostQuitMessage 0


    Case WM_SNARL_TRAY_ICON
        Select Case lParam
        Case WM_RBUTTONUP
            If Not fIgnoreNext Then
                uDoMainMenu

            Else
                fIgnoreNext = False

            End If

        Case WM_LBUTTONDBLCLK
            If g_NotificationRoster.ActualMissedCount > 0 Then
                g_NotificationRoster.ShowMissedPanel

            Else
                Me.NewDoPrefs

            End If

        End Select

    Case WM_ENTERMENULOOP
        mMenuOpen = True

    Case WM_EXITMENULOOP
        mMenuOpen = False

'    Case mTaskbarCreated
'        g_Debug "frmAbout.WndProc(): 'TaskbarCreated' received - adding icon..."
'        uAddTrayIcon
'
'    Case Is > WM_USER
'        Debug.Print ">> " & g_HexStr(uMsg, 4)

    Case WM_CLOSE
        Unload Me
        MWndProcSink_WndProc = True

    Case WM_INSTALL_SNARL
        If LoWord(wParam) = SNARL_CALLBACK_INVOKED Then _
            ShellExecute hWnd, "open", g_MakePath(App.Path) & gUpdateFilename, vbNullString, vbNullString, SW_SHOW

    Case RegisterWindowMessage("TaskbarCreated")
        ' /* R2.4 DR8 */
        uAddTrayIcon

    Case WM_WTSSESSION_CHANGE
        Select Case wParam
        Case WTS_SESSION_LOCK
            Debug.Print "WM_WTSSESSION_CHANGE: =locked= " & Now()
            If g_ConfigGet("away_when_locked") = "1" Then _
                g_SetPresence SP_AWAY_COMPUTER_LOCKED

        Case WTS_SESSION_UNLOCK
            Debug.Print "WM_WTSSESSION_CHANGE: =unlocked= " & Now()
            If g_ConfigGet("away_when_locked") = "1" Then _
                g_ClearPresence SP_AWAY_COMPUTER_LOCKED

        End Select



'    Case WM_REMOTENOTIFY
'
'        Debug.Print "WM_REMOTENOTIFY: token=" & lParam & " notification=" & LoWord(wParam)
'
'        dw = uFindRemoteNotification(lParam)
'        If dw = 0 Then
'            g_Debug "WM_REMOTENOTIFY: token " & g_HexStr(lParam) & " not in remote list", LEMON_LEVEL_CRITICAL
'
'        Else
'            mRemoteNotification(dw).Socket.Notify LoWord(wParam), CStr(lParam)
'            Select Case LoWord(wParam)
'            Case SNARL_NOTIFICATION_ACK, SNARL_NOTIFICATION_CLOSED
'                uRemoveRemoteNotification dw
'
'            End Select
'
'        End If
'
'        MWndProcSink_WndProc = True

    End Select

End Function

Private Sub uDoMainMenu()

    ' /* R2.31: only if admin says so */

    If gSysAdmin.InhibitMenu Then
        g_Debug "frmAbout.uDoMainMenu(): blocked by admin", LEMON_LEVEL_WARNING
        Exit Sub

    End If

Dim pi      As OMMenuItem
Dim i       As Long
Dim sz      As String
Dim szData  As String
Dim rc      As RECT
Dim hIcon   As Long
Dim update_config   As Boolean

    ' /* track the menu */

    SetForegroundWindow Me.hWnd

    With New OMMenu
        .AddItem .CreateItem("sticky", "Sticky Notifications", , , (g_ConfigGet("sticky_snarls") = "1"))
        .AddSeparator

        .AddItem .CreateItem("dnd", "Do Not Disturb", , , g_IsPresence(SP_DND_USER))
        .AddItem .CreateItem("missed", IIf(g_NotificationRoster.ActualMissedCount > 0, CStr(g_NotificationRoster.ActualMissedCount) & " ", "") & "Missed Notification" & IIf(g_NotificationRoster.ActualMissedCount = 1, "", "s") & "...")
        .AddSeparator
        .AddItem .CreateItem("restart", "Restart Snarl", , g_IsRunning)

        If g_IsRunning Then
            .AddItem .CreateItem("stop", "Stop Snarl")

        Else
            .AddItem .CreateItem("start", "Start Snarl")

        End If

        .AddItem .CreateItem("quit", "Quit Snarl", , Not gSysAdmin.InhibitQuit)
        .AddSeparator
        .AddItem .CreateItem("prefs", "Settings...", , Not gSysAdmin.InhibitPrefs)
'        .AddItem .CreateItem("app_mgr", "App Manager...")
        .AddItem .CreateItem("", "Snarl Apps", , , , , , g_AppRoster.SnarlAppsMenu())
        .AddSeparator
        .AddItem .CreateItem("about", "About Snarl...")

        Set pi = .Track(Me.hWnd)

    End With

    PostMessage hWnd, WM_NULL, 0, ByVal 0&

    If Not (pi Is Nothing) Then
        Select Case pi.Name
        Case "quit"
            PostQuitMessage 0
            Exit Sub

        Case "about"
            frmAbout.Show

        Case "restart"
            g_SetRunning False
            DoEvents
            Sleep 1500
            DoEvents
            g_SetRunning True

        Case "start"
            g_SetRunning True

        Case "stop"
            g_SetRunning False

        Case "prefs"
            Me.NewDoPrefs

'        Case "app_mgr"
'            ShellExecute 0, "open", g_MakePath(App.Path) & "SNARLAPP_Manager.exe", vbNullString, vbNullString, SW_SHOW

        Case "sticky"
            g_ConfigSet "sticky_snarls", IIf(g_ConfigGet("sticky_snarls") = "1", "0", "1")

        Case "dnd"
            If g_IsPresence(SP_DND_USER) Then
                ' /* clear it */
                g_ClearPresence SP_DND_USER

            Else
                ' /* set it */
                g_SetPresence SP_DND_USER
                g_NotificationRoster.ResetMissedCount

            End If

        Case "missed"
            If Not (g_NotificationRoster Is Nothing) Then _
                g_NotificationRoster.ShowMissedPanel

        Case Else
            If g_SafeLeftStr(pi.Name, 1) = "!" Then _
                g_AppRoster.SnarlAppDo Val(g_SafeRightStr(pi.Name, Len(pi.Name) - 1)), SNARLAPP_DO_PREFS

'            sz = g_SafeLeftStr(pi.Name, 3)
'            szData = g_SafeRightStr(pi.Name, Len(pi.Name) - 3)
'
'            Select Case sz
'            Case "cfg"
'                ' /* Snarl App -> Settings... szData is App Roster index */
'                g_AppRoster.SnarlAppDo Val(szData), SNARLAPP_SHOW_PREFS
'
'            Case "abt"
'                ' /* Snarl App -> About... szData is App Roster index */
'                g_AppRoster.SnarlAppDo Val(szData), SNARLAPP_SHOW_ABOUT
'
'            End Select

        End Select
    End If

    If update_config Then _
        g_WriteConfig

End Sub

Public Sub NewDoPrefs(Optional ByVal PageToSelect As Integer)

    ' /* R2.31: only if admin says we can... */

    If gSysAdmin.InhibitPrefs Then
        g_Debug "frmAbout.NewDoPrefs(): access blocked by admin", LEMON_LEVEL_WARNING
        MsgBox "Access to Snarl's preferences has been blocked by your system administrator.", vbInformation Or vbOKOnly, App.Title
        Exit Sub

    End If

Dim pp As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me

            mPanel.SetTitle "Snarl Preferences"
            mPanel.SetWidth 540

            ' /* general page */

            Set pp = new_BPrefsPage("General", load_image_obj(g_MakePath(App.Path) & "etc\icons\general.png"), Me)

            With pp
                .SetMargin 96

                ' /* launch options */

                .Add new_BPrefsControl("banner", "", "Launch Options")
                .Add new_BPrefsControl("fancytoggle2", "run_on_logon", "Start at login?", "", g_ConfigGet("run_on_logon"))
                .Add new_BPrefsControl("fancytoggle2", "show_msg_on_start", "Show Welcome Message on startup?", "", g_ConfigGet("show_msg_on_start"), pm)
                .Add new_BPrefsControl("fancytoggle2", "auto_update", "Check for updates on launch?", "", g_ConfigGet("auto_update"))
                .Add new_BPrefsControl("fancybutton2", "update_now", "Check now...")

                ' /* applications */

                .Add new_BPrefsControl("banner", "", "Applications")
                .Add new_BPrefsControl("fancytoggle2", "notify_on_first_register", "Only notify the first time an application registers?", , g_ConfigGet("notify_on_first_register"))

                ' /* forwarding */

                .Add new_BPrefsControl("banner", "", "Forwarding")
                .Add new_BPrefsControl("fancytoggle2", "include_host_name_when_forwarding", "Include computer name when forwarding notifications?", , g_ConfigGet("include_host_name_when_forwarding"))

                ' /* misc */

                .Add new_BPrefsControl("banner", "", "Miscellaneous")
                .Add new_BPrefsControl("fancytoggle2", "log_only", "Log only (don't display)?", , g_ConfigGet("log_only"))

    '        .Add new_BPrefsControl("fancytoggle2", "sticky_snarls", "Sticky notifications?", , g_ConfigGet("sticky_snarls"))

            End With

            .AddPage pp


            ' /* apps */


            Set mAppsPage = New TAppsPage
            .AddPage new_BPrefsPage("Apps", load_image_obj(g_MakePath(App.Path) & "etc\icons\apps.png"), mAppsPage)
            .AddPage new_BPrefsPage("Display", load_image_obj(g_MakePath(App.Path) & "etc\icons\display.png"), New TDisplayPage)
            .AddPage new_BPrefsPage("Styles", load_image_obj(g_MakePath(App.Path) & "etc\icons\styles.png"), New TStylesPage)
            .AddPage new_BPrefsPage("Extensions", load_image_obj(g_MakePath(App.Path) & "etc\icons\extensions.png"), New TExtPage)
            .AddPage new_BPrefsPage("Network", load_image_obj(g_MakePath(App.Path) & "etc\icons\network.png"), New TNetworkPage)


            ' /* presence */


            Set pp = new_BPrefsPage("Presence", load_image_obj(g_MakePath(App.Path) & "etc\icons\presence.png"), Me)

            With pp
                .SetMargin 96

                ' /* away mode */

                .Add new_BPrefsControl("banner", "", "Away Mode")
                .Add new_BPrefsControl("label", "", "Set away after the following period of inactivity:")

                Set pm = New CTempMsg
                pm.Add "min", 0&
                pm.Add "max", 30&
                pm.Add "freq", 5&
                pm.Add "label_size", 56&
                .Add new_BPrefsControl("fancyslider", "idle_minutes", "", "", IIf(g_ConfigGet("idle_minutes") = "2", "1", "2"), pm)

                .Add new_BPrefsControl("fancytoggle2", "away_when_locked", "Set away when computer is locked?", , g_ConfigGet("away_when_locked"))
                .Add new_BPrefsControl("fancytoggle2", "away_when_screensaver", "Set away when the screensaver starts?", , g_ConfigGet("away_when_screensaver"))

                .Add new_BPrefsControl("fancycycle", "away_mode", "Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority|Forward...", "When Away:", g_ConfigGet("away_mode"))
                .Add new_BPrefsCombo("away_style", g_StyleRoster.GetNonWindowStyleList(), "", uFindForward(g_ConfigGet("away_style"), g_StyleRoster.GetNonWindowStyleList()), 24, , , , (g_ConfigGet("away_mode") = "6"))

                ' /* busy mode */

                .Add new_BPrefsControl("banner", "", "Busy Mode")
                .Add new_BPrefsControl("fancytoggle2", "away_when_fullscreen", "Set busy when the foreground application is fullscreen?", , g_ConfigGet("away_when_fullscreen"))
                .Add new_BPrefsControl("fancycycle", "busy_mode", "Log as Missed|Display Sticky|Discard|Display Normally|Display High Priority|Forward...", "When Busy:", g_ConfigGet("busy_mode"))
                .Add new_BPrefsCombo("busy_style", g_StyleRoster.GetNonWindowStyleList(), "", uFindForward(g_ConfigGet("busy_style"), g_StyleRoster.GetNonWindowStyleList()), 24, , , , (g_ConfigGet("busy_mode") = "6"))

                .Add new_BPrefsControl("label", "", "Note that the above settings only apply to normal priority notifications.  More granular settings are available in the class configuration preferences panel.")

            End With

            .AddPage pp




            ' /* advanced page */

            Set pp = new_BPrefsPage("Advanced", load_image_obj(g_MakePath(App.Path) & "etc\icons\advanced.png"), Me)

            With pp
                .SetMargin 96

                ' /* hotkeys */

                .Add new_BPrefsControl("banner", "", "Hotkeys")
                .Add new_BPrefsControl("fancytoggle2", "use_hotkey", "Use a hotkey to activate Snarl's Preferences?", "", g_ConfigGet("use_hotkey"))
                .Add new_BPrefsControl("key_picker", "hotkey_prefs", , , CStr(MOD_CONTROL) & "," & g_ConfigGet("hotkey_prefs"), , (g_ConfigGet("use_hotkey") = "1"))
                .Add new_BPrefsControl("fancytoggle2", "", "Use a hotkey to activate Snarl's menu?", "", "0", , False)
                .Add new_BPrefsControl("key_picker", "", , , CStr(MOD_WIN) & "," & g_ConfigGet("hotkey_prefs"), , False)
                .Add new_BPrefsControl("label", "", "Press the key you want to use in the boxes above.  Note that the modifiers (the combination of SHIFT and CTRL keys) used are automatically set.")

                ' /* other */

                .Add new_BPrefsControl("banner", "", "Other")
                .Add new_BPrefsControl("fancybutton2", "go_tray_icon", "Recreate Tray Icon")

'                .Add new_BPrefsControl("banner", "", "Presence Management")
'                .Add new_BPrefsControl("fancycycle", "away_mode", "Log as Missed|Make Sticky|Discard|Display", "When Away:", g_ConfigGet("away_mode"))
'                .Add new_BPrefsControl("fancycycle", "busy_mode", "Log as Missed|Make Sticky|Discard|Display", "When Busy:", g_ConfigGet("busy_mode"))

        ' /* other stuff */

'        .Add new_BPrefsControl("banner", "", "System Functions")
'        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Reload Styles")
'        .Add new_BPrefsControl("label", "", "Forces Snarl to reload all installed styles.  Under normal circumstances you shouldn't need to do this; it's provided for users who are developing their own styles and want to test them without restarting Snarl.")

            End With

            .AddPage pp


            .AddPage new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "etc\icons\about.png"), New TAboutPage)

            If gDebugMode Then

                Set pp = new_BPrefsPage("Debug", load_image_obj(g_MakePath(App.Path) & "etc\icons\debug.png"), Me)

                With pp
                    .SetMargin 96
                    .Add new_BPrefsControl("banner", "", "Debugging")
                    .Add new_BPrefsControl("fancybutton2", "go_lemon", "Open debug log")
            '        .Add new_BPrefsControl("label", "", "The log file can be useful for debugging purposes.")
            
                    .Add new_BPrefsControl("fancybutton2", "go_garbage", "Garbage collection", , , , g_IsWinXPOrBetter())
            
            '        .Add new_BPrefsControl("separator", "")
                    .Add new_BPrefsControl("banner", "", "Configuration")
                    .Add new_BPrefsControl("fancybutton2", "open_config", "Open config folder")
                    .Add new_BPrefsControl("label", "", "Opens the current config folder in Explorer so the various configuration files can be edited manually.")
            
            '        .Add new_BPrefsControl("fancybutton2", "cycle_config", "Reload Config File")
            '        .Add new_BPrefsControl("label", "", "Reloads the current configuration file.")
            
            '        .Add new_BPrefsControl("separator", "")
                    .Add new_BPrefsControl("banner", "", "Diagnostics")
                    .Add new_BPrefsControl("fancybutton2", "test", "Test notification")
                    .Add new_BPrefsControl("label", "", "Sends a special test message to the Snarl engine which should result in a notification appearing.  This message is sent using the same mechanism a 3rd party application would use and therefore should prove (or otherwise) that the Snarl notification engine is running correctly.")
            
            '        .Add new_BPrefsControl("separator", "")
            '        .Add new_BPrefsControl("fancybutton2", "restart_style_roster", "Restart Style Roster")
            
                End With

                .AddPage pp

            End If

            .Go
            g_SetWindowIconToAppResourceIcon .hWnd
            SetForegroundWindow .hWnd

        End With

    End If

    If (PageToSelect > 0) And (PageToSelect <= mPanel.CountPages) Then _
        mPanel.SetPage PageToSelect

End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    ' /* prefs panel has been closed */
    Set mPanel = Nothing
    Set mAppsPage = Nothing

    g_WriteConfig

End Sub

Private Sub KPrefsPanel_Ready()
Dim pc As BControl

    ' /* panel is now ready and visible, so select the first item in the registered apps combo - this then
    '    cascades a changed event down which configures all the other controls on that page */

    If mPanel.Find("cb>apps", pc) Then _
        pc.SetValue "1"

    ' /* R2.4 DR8: set here so we pick up the custom label change */

    If mPanel.Find("idle_minutes", pc) Then _
        pc.SetValue g_ConfigGet("idle_minutes")

    ' /* find our current style and select it in the 'Display' sub page */

Dim i As Long
Dim px As TStyle
Dim j As Long

'    Debug.Print gPrefs.default_style

    If Not (g_StyleRoster Is Nothing) Then
        i = g_StyleRoster.IndexOf(style_GetStyleName(g_ConfigGet("default_style")))
        If i Then
            Set px = g_StyleRoster.StyleAt(i)
            j = px.SchemeIndex(style_GetSchemeName(g_ConfigGet("default_style")))

            If j Then
                ' /* R2.4 RC1: select default style and scheme in [Styles] page*/
                prefskit_SetValue mPanel, "installed_styles", CStr(i)
                prefskit_SetValue mPanel, "installed_schemes", CStr(j)

'            prefskit_SetValue mPanel, "default_style", CStr(i)
'            prefskit_SetValue mPanel, "default_scheme", CStr(j)
'
'            If mPanel.Find("default_style", pc) Then _
'                pc.SetValue CStr(i)
'
'            If mPanel.Find("default_scheme", pc) Then _
'                pc.SetValue CStr(j)

            End If
        End If
    End If

    If mPanel.Find("busy_style", pc) Then _
        g_StyleRoster.SetNonWindowStyleIcons pc

    If mPanel.Find("away_style", pc) Then _
        g_StyleRoster.SetNonWindowStyleIcons pc

'    If mPanel.Find("melontype_contrast", pc) Then _
'        pc.SetEnabled (gPrefs.font_smoothing = E_MELONTYPE)

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)

End Sub

Private Sub Label1_Click()

    ShellExecute Me.hWnd, "open", "http://www.fullphat.net/", vbNullString, vbNullString, SW_SHOW

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = False Then _
        Label1.Font.Underline = True

End Sub

Private Property Get MMessageSink_Name() As String

    MMessageSink_Name = "p>snarl"

End Property

Private Function MMessageSink_Received(message As melon.MMessage) As Boolean

End Function

Friend Sub bUpdateExtList()

    If (mPanel Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If Not (mPanel.Find("lb>extensions", pc)) Then _
        Exit Sub

    ' /* how's this? we send a fake control notification which
    '    translates as a request to update the extensions list! */

    pc.Notify "update_list", Nothing

End Sub

Private Sub uAddTrayIcon()
Dim hIcon As Long

    If (mTrayIcon Is Nothing) Or (g_ConfigGet("show_tray_icon") = "0") Or (gSysAdmin.HideIcon) Then _
        Exit Sub

    hIcon = LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0)
    If hIcon = 0 Then _
        hIcon = Me.Icon.Handle

    mTrayIcon.Remove "tray_icon"

    mTrayIcon.Add "tray_icon", hIcon, "Snarl"

End Sub

Friend Sub bMissedNotificationsChanged()

    On Error Resume Next

    If (mTrayIcon Is Nothing) Then _
        Exit Sub

Dim hIcon As Long

    If g_NotificationRoster.ActualMissedCount > 0 Then

        hIcon = LoadResPicture(50, vbResIcon).Handle
        If hIcon = 0 Then _
            hIcon = Me.Icon.Handle

        mTrayIcon.Update "tray_icon", hIcon, "Snarl - " & CStr(g_NotificationRoster.ActualMissedCount) & " missed notification" & IIf(g_NotificationRoster.ActualMissedCount = 1, "", "s")

    Else
        hIcon = LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0)
        If hIcon = 0 Then _
            hIcon = Me.Icon.Handle
        
        mTrayIcon.Update "tray_icon", hIcon, "Snarl"

    End If

End Sub

Private Function uIsAlertEnabled(ByVal ConfigString As String) As Boolean
Dim sz() As String

    On Error Resume Next

    sz() = Split(ConfigString, "#?")
    uIsAlertEnabled = Val(sz(0))

End Function

Private Sub JSONSocket_OnConnect()

    g_Debug "JSONSocket.OnConnect()", LEMON_LEVEL_PROC

End Sub

Private Sub JSONSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "JSONSocket.OnConnectionRequest(): requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

    mSockets = mSockets + 1
    ReDim Preserve mJSONSocket(mSockets)
    Set mJSONSocket(mSockets) = New CJSONSocket
    mJSONSocket(mSockets).Accept requestID

End Sub

Private Sub JSONSocket_OnDataArrival(ByVal bytesTotal As Long)

    g_Debug "JSONSocket.OnDataArrival(): bytesTotal=" & g_HexStr(bytesTotal), LEMON_LEVEL_PROC

End Sub

Friend Function bSetHotkeys(Optional ByVal KeyCode As Long = 0) As Boolean

    ' /* return True if the prefs hotkey was registered ok */

    If g_ConfigGet("use_hotkey") = "0" Then
        ' /* hotkeys not enabled */
        uUnregisterHotkeys
        bSetHotkeys = True
        Exit Function

    End If

    If KeyCode = 0 Then
        g_Debug "bSetHotKeys(): registering existing hotkey (" & g_ConfigGet("hotkey_prefs") & ")", LEMON_LEVEL_INFO
        KeyCode = Val(g_ConfigGet("hotkey_prefs"))

    End If

Dim hSysKey As Long

    ' /* attempt to register the CTRL+keycode combo: if this fails, we fail */

    hSysKey = register_system_key(Me.hWnd, KeyCode, B_SYSTEM_KEY_CONTROL)
    g_Debug "bSetHotkeys(): register_system_key('prefs'): " & (hSysKey <> 0)
    If hSysKey = 0 Then _
        Exit Function

    ' /* registered okay, so unregister the existing hotkeys */

    uUnregisterHotkeys
    mSysKeyPrefs = hSysKey

    ' /* attempt to register the CTRL+SHIFT+keycode combo as well - don't fail if this fails though */

    hSysKey = register_system_key(Me.hWnd, KeyCode, B_SYSTEM_KEY_SHIFT Or B_SYSTEM_KEY_CONTROL)
    g_Debug "bSetHotkeys(): register_system_key('test'): " & (hSysKey <> 0)
    If hSysKey <> 0 Then _
        mSysKeyTest = hSysKey

    bSetHotkeys = True

End Function

Private Sub uUnregisterHotkeys()

    g_Debug "uUnregisterHotKeys(): unregister_system_key('prefs'): " & unregister_system_key(Me.hWnd, mSysKeyPrefs)
    g_Debug "uUnregisterHotKeys(): unregister_system_key('test'): " & unregister_system_key(Me.hWnd, mSysKeyTest)

    mSysKeyPrefs = 0
    mSysKeyTest = 0

End Sub

Friend Sub bUpdateAppList()
Dim pc As BControl

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("cb>apps", pc) Then _
            pc.Notify "update_list", Nothing

    End If

End Sub

Friend Sub bUpdateClassList(ByVal AppToken As Long)
Dim pc As BControl
Dim pm As CTempMsg

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("cb>apps", pc) Then
            Set pm = New CTempMsg
            pm.What = AppToken
            pc.Notify "update_classes", pm

        End If
    End If

End Sub

Private Sub uUpdateStyleList()
Dim pc As BControl

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("installed_styles", pc) Then _
            pc.Notify "update_list", Nothing

    End If

End Sub

Friend Sub bUpdateRemoteComputerList()
Dim pc As BControl

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("lb>forward", pc) Then _
            pc.Notify "update_list", Nothing

    End If

End Sub

Private Sub uDoSysInfoNotification()
Dim szMetric As String
Dim szMelon As String
Dim dFreq As Double
Dim hKey As Long
Dim dw As Long
Dim cb As Long

    ' /* empty working set */

    EmptyWorkingSet GetCurrentProcess()

    ' /* read melon version from registry */

    If RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\melon", hKey) = ERROR_SUCCESS Then
        If RegQueryValueEx(hKey, "DisplayVersion", 0, dw, ByVal 0&, cb) = ERROR_SUCCESS Then
            If dw = REG_SZ Then
                szMelon = String$(cb, 0)
                If RegQueryValueEx(hKey, "DisplayVersion", 0, ByVal 0&, ByVal szMelon, cb) = ERROR_SUCCESS Then _
                    szMelon = Left$(szMelon, cb - 1)

            End If
        End If
    End If

    With New BCPU
        .SetTo 1
        dw = g_GetCPUCount()

        dFreq = .Speed
        If dFreq > 1000# Then
            dFreq = dFreq / 1000
            szMetric = "GHz"

        Else
            szMetric = "MHz"

        End If

        g_PrivateNotify "", g_GetUserName() & " on " & g_GetComputerName(), _
                        g_GetOSName() & " " & g_GetServicePackName() & vbCrLf & _
                        IIf(dw > 1, CStr(dw) & "x", "") & .FullName & " @ " & Format$(dFreq, "0.0#") & " " & szMetric & vbCrLf & _
                        g_FileSizeToStringEx2(g_GetPhysMem(True), "GB", "", "0.0") & " (" & g_FileSizeToStringEx2(g_GetPageMem(True) + g_GetPhysMem(True), "GB", "", "0.0") & ") RAM" & vbCrLf & _
                        "Snarl " & App.Major & "." & App.Revision & " (" & App.Comments & ")" & vbCrLf & "melon " & IIf(szMelon <> "", szMelon, "??"), _
                        -1, _
                        g_MakePath(App.Path) & "etc\icons\snarl.png"

    End With

End Sub

'Private Sub myDownloadUpdateRequest_Completed()
'
'    mDownloadId = snShowMessage("Update Available", _
'                                "An update for Snarl is available.  To install it now click this notification with the left mouse button", _
'                                0, _
'                                g_MakePath(App.Path) & "etc\icons\snarl.png", Me.hWnd, WM_INSTALL_SNARL)
'
'    Set myDownloadUpdateRequest = Nothing
'    bDownloadPanelQuit
'
'End Sub
'
'Private Sub myDownloadUpdateRequest_TimedOut()
'
'    MsgBox "myDownloadUpdateRequest_TimedOut"
'    Set myDownloadUpdateRequest = Nothing
'    bDownloadPanelQuit
'
'End Sub

Private Sub theIdleTimer_Pulse()
Static b As Boolean

    If g_ConfigGet("away_when_fullscreen") = "1" Then
        ' /* track foreground app state */
        b = uIsFullScreenMode()
        If b <> g_IsPresence(SP_DND_FULLSCREEN_APP) Then
            ' /* full screen app state changed */
            If b Then
                g_SetPresence SP_DND_FULLSCREEN_APP

            Else
                g_ClearPresence SP_DND_FULLSCREEN_APP

            End If

            g_Debug "_theIdleTimer.Pulse(): " & Now() & " fullscreen app: " & g_IsPresence(SP_DND_FULLSCREEN_APP)

        End If
    End If


Dim n As Long

    If g_ConfigGet("away_when_screensaver") = "1" Then
        ' /* track screensaver state */

        If SystemParametersInfo(SPI_GETSCREENSAVERRUNNING, 0, n, 0) <> 0 Then
            If b <> g_IsPresence(SP_AWAY_SCREENSAVER_ACTIVE) Then
                ' /* screensaver state has changed */
                If b Then
                    g_SetPresence SP_AWAY_SCREENSAVER_ACTIVE
        
                Else
                    g_ClearPresence SP_AWAY_SCREENSAVER_ACTIVE
        
                End If
        
                g_Debug "_theIdleTimer.Pulse(): " & Now() & " screensaver: " & g_IsPresence(SP_AWAY_SCREENSAVER_ACTIVE)

            End If
        End If
    End If

    ' /* ignore if no idle timeout set */

    n = g_SafeLong(g_ConfigGet("idle_minutes"))
    If n > 30 Then _
        n = 30              ' // bounds-check

    n = n * 60000           ' // convert to ms
    If n < 1 Then _
        Exit Sub


Dim lii As LASTINPUTINFO

    lii.cbSize = Len(lii)
    If GetLastInputInfo(lii) = False Then _
        Exit Sub

    lii.dwTime = GetTickCount() - lii.dwTime
'    Debug.Print "_theIdleTimer.Pulse(): idle time is now " & CStr(lii.dwTime) & " needs to be " & CStr(n)

    b = (lii.dwTime > n)
    If b <> g_IsPresence(SP_AWAY_USER_IDLE) Then
        ' /* idle state has changed */
        If b Then
            g_SetPresence SP_AWAY_USER_IDLE

        Else
            g_ClearPresence SP_AWAY_USER_IDLE

        End If

        g_Debug "_theIdleTimer.Pulse(): " & Now() & " user idle: " & g_IsPresence(SP_AWAY_USER_IDLE)

    End If

End Sub

Private Sub theReadyTimer_Pulse()

    ' /* tell everyone we're open for business */

    g_Debug "Notifying ready to run..."
    PostMessage HWND_BROADCAST, g_GlobalMessage(), SNARL_BROADCAST_LAUNCHED, ByVal CLng(App.Major)

End Sub

Private Sub Timer1_Timer()
Dim pWindow As CSnarlWindow
Dim pt As POINTAPI
Dim i As Long

    If (g_NotificationRoster Is Nothing) Or (mMenuOpen) Then _
        Exit Sub

    GetCursorPos pt
    i = g_NotificationRoster.HitTest(pt.x, pt.y)

    If i > 0 Then
        Set pWindow = g_NotificationRoster.NotificationAt(i)

        ' /* existing? */

        If Not (mClickThruOver Is Nothing) Then
            If mClickThruOver.Id <> pWindow.Id Then
                ' /* different notification */
                mClickThruOver.MakeFuzzy False
                Set mClickThruOver = Nothing

            End If
        End If

'        Debug.Print pWindow.Window.hWnd & " " & pWindow.NotificationOnlyMode

        If pWindow.NotificationOnlyMode Then
            pWindow.MakeFuzzy True
            Set mClickThruOver = pWindow

        End If

    Else
        ' /* reset current */
        If Not (mClickThruOver Is Nothing) Then
            mClickThruOver.MakeFuzzy False
            Set mClickThruOver = Nothing

        End If

    End If

    ' /* R2.31 - purge unused memory pages if we appear to be idle */

'    If Abs(DateDiff("n", Now, gLastNotification)) > 5 Then
'        EmptyWorkingSet GetCurrentProcess()
'        gLastNotification = Now()           ' // fix timestamp so we don't constant do this...
'
'    End If

End Sub

Friend Sub bUpdateStylesList()
Dim pc As BControl

    If Not (mPanel Is Nothing) Then
        If mPanel.Find("installed_styles", pc) Then _
            pc.Notify "update_list", Nothing

    End If

End Sub

'Friend Sub bStartUpdateDownload()
'
'    If Not (myDownloadUpdateRequest Is Nothing) Then
'        MsgBox "bStartUpdateDownload(): Request is still processing.", vbCritical Or vbOKOnly, App.Title
'        Exit Sub
'
'    End If
'
'    Set myDownloadUpdateRequest = New CHTTPRequest
'    myDownloadUpdateRequest.GetFile "http://www.fullphat.net/updateinfo/" & gUpdateFilename, g_MakePath(App.Path) & gUpdateFilename, 60
'
'End Sub
'
'Friend Function bIsDownloadingUpdate() As Boolean
'
'    If (myDownloadUpdateRequest Is Nothing) Then _
'        Exit Function
'
'    bIsDownloadingUpdate = True
'
'End Function
'
'Friend Sub bDownloadPanelQuit()
'
'    ' /* the "An Update is Available" panel was closed - but was a download initiated? */
'
'    If (mPanel Is Nothing) Then _
'        Exit Sub
'
'Dim pc As BControl
'
'    ' /* "Check for Update" button */
'
'    If mPanel.Find("go_updates", pc) Then
'        If Me.bIsDownloadingUpdate Then
'            pc.SetText "Downloading Update..."
'            pc.SetEnabled False
'
'        Else
'            pc.SetText "Check for Update"
'            pc.SetEnabled True
'
'        End If
'    End If
'
'    If mPanel.Find("lbl_updates", pc) Then
'        If Me.bIsDownloadingUpdate Then
'            pc.SetText "You will be notified when the update is ready to be installed."
'
'        Else
'            pc.SetText " "
'
'        End If
'    End If
'
'End Sub

'Public Sub AddRemoteNotification(ByVal Token As Long, ByRef Socket As TRemoteConnection)
'
'    mRemoteNotifications = mRemoteNotifications + 1
'    ReDim Preserve mRemoteNotification(mRemoteNotifications)
'    With mRemoteNotification(mRemoteNotifications)
'        .Token = Token
'        Set .Socket = Socket
'
'    End With
'
'End Sub
'
'Private Function uFindRemoteNotification(ByVal Token As Long) As Long
'Dim i As Long
'
'    If mRemoteNotifications = 0 Then _
'        Exit Function
'
'    For i = 1 To mRemoteNotifications
'        If mRemoteNotification(i).Token = Token Then
'           uFindRemoteNotification = i
'           Exit Function
'
'        End If
'
'    Next i
'
'End Function
'
'Private Sub uRemoveRemoteNotification(ByVal Index As Long)
'Dim i As Long
'
'    If (Index < 1) Or (Index > mRemoteNotifications) Then _
'        Exit Sub
'
'    If Index < mRemoteNotifications Then
'        For i = Index To (mRemoteNotifications - 1)
'            LSet mRemoteNotification(i) = mRemoteNotification(i + 1)
'
'        Next i
'
'    End If
'
'    mRemoteNotifications = mRemoteNotifications - 1
'    ReDim Preserve mRemoteNotification(mRemoteNotifications)
'
'End Sub

Public Sub EnableJSON(ByVal Enabled As Boolean)

    g_Debug "frmAbout.EnableJSON(" & CStr(Enabled) & ")", LEMON_LEVEL_PROC_ENTER

    If Enabled Then
        g_Debug "creating JSON listener..."

        Set JSONSocket = New CSocket
        JSONSocket.Bind "9889", "127.0.0.1"
        JSONSocket.Listen

        g_Debug "listening on " & JSONSocket.LocalIP & ":" & JSONSocket.LocalPort & "..."

    Else
        g_Debug "stopping JSON listener..."
        If Not (JSONSocket Is Nothing) Then
            JSONSocket.CloseSocket
            Set JSONSocket = Nothing
            g_Debug "JSON listener stopped"

        Else
            g_Debug "wasn't started", LEMON_LEVEL_WARNING

        End If
    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Sub EnableSNP(ByVal Enabled As Boolean)
Dim szAddr() As String
Dim i As Long

    g_Debug "frmAbout.EnableSNP(" & CStr(Enabled) & ")", LEMON_LEVEL_PROC_ENTER

    If Enabled Then
        g_Debug "getting local ip address table..."
        ' /* get local ip addresses */
        szAddr() = Split(get_ip_address_table(), " ")

        If UBound(szAddr()) > -1 Then

            ' /* add SNP/tcp and GNTP listeners */

            For i = 0 To UBound(szAddr())
                If szAddr(i) <> "0.0.0.0" Then
                    uAddListener szAddr(i), False                   ' // SNP listener
                    uAddListener szAddr(i), True                    ' // GNTP listener

                End If

            Next i

        Else
            g_Debug "couldn't read local ip address table", LEMON_LEVEL_WARNING

        End If

        ' /* R2.4: native Growl/UDP support */

        g_Debug "creating Growl UDP socket..."

        Set GrowlUDPSocket = New CSocket
        With GrowlUDPSocket
            .Protocol = sckUDPProtocol
            .Bind 9887

        End With

    Else

        g_Debug "closing Growl UDP socket..."

        If Not (GrowlUDPSocket Is Nothing) Then
            GrowlUDPSocket.CloseSocket
            Set GrowlUDPSocket = Nothing

        End If

        g_Debug "stopping SNP listeners..."

        If mListenerCount Then
            For i = 1 To mListenerCount
                mListener(i).Quit

            Next i

        Else
            g_Debug "no listeners", LEMON_LEVEL_WARNING

        End If

        ReDim mListener(0)

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function DoExtensionConfig(ByVal Index As Long) As Boolean

    If (mPanel Is Nothing) Then _
        Exit Function

    If IsWindowEnabled(mPanel.hWnd) = 0 Then _
        Exit Function

Dim pExtList As BControl

    If mPanel.Find("lb>extensions", pExtList) Then _
        pExtList.SetValue CStr(Index)

Dim pExt As TExtension

    Set pExt = g_ExtnRoster.ItemAt(Index)
    DoExtensionConfig = pExt.DoPrefs(mPanel.hWnd)

End Function

Public Function DoStyleConfig(ByVal Index As Long) As Boolean

    If (mPanel Is Nothing) Then _
        Exit Function

    If IsWindowEnabled(mPanel.hWnd) = 0 Then _
        Exit Function

Dim pStyleList As BControl

    If mPanel.Find("installed_styles", pStyleList) Then _
        pStyleList.SetValue CStr(Index)

    If mPanel.Find("ftb>style", pStyleList) Then _
        pStyleList.Changed "1"

    DoStyleConfig = True

End Function

Public Sub DoAppConfig(ByVal AppName As String, Optional ByVal ClassName As String)

    NewDoPrefs 2

Dim i As Long

    If Not (g_AppRoster Is Nothing) Then
        i = g_AppRoster.IndexOf(AppName)
        If i Then
            ' /* select the application */
            prefskit_SetValue mPanel, "cb>apps", CStr(i)

            ' /* find the class */
            i = g_AppRoster.AppAt(i).IndexOf(ClassName)
            If i = 0 Then
                ' /* not found/null - select _all */
                prefskit_SetValue mPanel, "lb>classes", "1"

            Else
                ' /* select it */
                prefskit_SetValue mPanel, "lb>classes", CStr(i)

            End If

        Else
            g_Debug "frmAbout.DoAppConfig(): '" & AppName & "' not found", LEMON_LEVEL_CRITICAL

        End If

    Else
        g_Debug "frmAbout.DoAppConfig(): app roster not available", LEMON_LEVEL_CRITICAL

    End If

End Sub

Public Sub DoAppConfigBySignature(ByVal Signature As String)

    On Error Resume Next

    If (g_AppRoster Is Nothing) Then _
        Exit Sub

Dim i As Long

    i = Val(g_AppRoster.IndexOfSig(Signature))
    If i = 0 Then
        g_Debug "frmAbout.DoAppConfigBySignature(): '" & Signature & "' not in app roster"
        Exit Sub

    End If

    ' /* show the apps page */

    NewDoPrefs 2

    ' /* select the app */

    prefskit_SetValue mPanel, "cb>apps", CStr(g_AppRoster.IndexOfSig(Signature))

    ' /* do a configure... */

Dim pc As BControl

    If mPanel.Find("fb>cfg_class", pc) Then _
        mPanel.PageAt(2).ControlInvoked pc

End Sub

Friend Sub bReadyToRun()

    Set theReadyTimer = new_BTimer(2000, True)

End Sub

Private Sub uAddListener(ByVal IPAddr As String, ByVal IsGNTP As Boolean)

    mListenerCount = mListenerCount + 1
    ReDim Preserve mListener(mListenerCount)
    Set mListener(mListenerCount) = New CSnarlListener
    mListener(mListenerCount).Go IPAddr, IsGNTP

End Sub



Private Function uIsFullScreenMode() As Boolean
Static hWnd As Long
Static h As Long

    hWnd = uParentFromPoint(1, 1)

'    g_Debug g_ClassName(hWnd) & " " & _
            g_ClassName(uParentFromPoint(g_ScreenWidth() - 1, 1)) & " " & _
            g_ClassName(uParentFromPoint(1, g_ScreenHeight() - 1)) & " " & _
            g_ClassName(uParentFromPoint(g_ScreenWidth() - 1, g_ScreenHeight() - 1))

    If hWnd = uParentFromPoint(uScreenWidth() - 1, 1) Then
        If hWnd = uParentFromPoint(uScreenWidth() - 1, uScreenHeight() - 1) Then
            If hWnd = uParentFromPoint(1, uScreenHeight() - 1) Then

'                g_Debug "uIsFullScreenMode(): four points match: " & g_ClassName(hWnd) & " '" & g_WindowText(hWnd) & "'"

                h = GetWindow(hWnd, GW_HWNDPREV)
                Do While h
                    If uIsAppWindow(h) Then _
                        Exit Function

                    h = GetWindow(h, GW_HWNDPREV)

                Loop

'                g_Debug "uIsFullScreenMode(): no higher app window"
                uIsFullScreenMode = True

            End If
        End If
    End If

End Function

Private Function uGetTopLevel(ByVal hWnd As Long) As Long
Static h As Long

    uGetTopLevel = hWnd
    h = GetParent(uGetTopLevel)
    Do While h
        uGetTopLevel = h
        h = GetParent(uGetTopLevel)

    Loop

End Function

Private Function uParentFromPoint(ByVal x As Long, ByVal y As Long) As Long
Static h As Long

    h = WindowFromPoint(x, y)
    If IsWindow(h) = 0 Then _
        Exit Function

    uParentFromPoint = uGetTopLevel(h)

End Function

Private Function uScreenWidth(Optional ByVal VirtualScreen As Boolean = False) As Long

    uScreenWidth = GetSystemMetrics(SM_CXSCREEN)

End Function

Private Function uScreenHeight(Optional ByVal VirtualScreen As Boolean = False) As Long

    uScreenHeight = GetSystemMetrics(SM_CYSCREEN)

End Function

Private Function uIsAppWindow(ByVal hWnd As Long) As Boolean
Static lExStyle As Long
Static Style As Long

    ' /* more reliable version (although can pick up the 'wrong' window in cases
    '    where there's a choice (e.g. VB IDE and Platform SDK) - code taken from
    '    here: http://shell.franken.de/~sky/explorer-doc/taskbar_8cpp-source.html */

    ' /* modified 7-Sep-09 to also include dialog windows where the owner is an
    '    app window.  To filter these out, just exclude any hWnd where GW_OWNER is
    '    not zero */

    Style = GetWindowLong(hWnd, GWL_STYLE)
    lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

    If (Style And WS_VISIBLE) = 0 Then _
        Exit Function

    If (lExStyle And WS_EX_APPWINDOW) Then
        uIsAppWindow = True
        Exit Function

    End If

    If (lExStyle And WS_EX_TOOLWINDOW) = 0 Then
        If (GetParent(hWnd) = 0) And (GetWindow(hWnd, GW_OWNER) = 0) Then
            uIsAppWindow = True

        Else
            uIsAppWindow = uIsAppWindow(GetWindow(hWnd, GW_OWNER))

        End If
    End If

End Function

Private Function uFindForward(ByVal StyleAndScheme As String, ByVal List As String) As Long
Dim s() As String

    On Error Resume Next

    err.Clear
    s = Split(Replace$(LCase$(List), ": ", "/"), "|")
    If UBound(s) < 1 Then _
        Exit Function

    If err.Number Then _
        Exit Function

Dim i As Long

    For i = 0 To UBound(s)
        If s(i) = StyleAndScheme Then
            uFindForward = i + 1
            Exit Function

        End If
    Next i

End Function


