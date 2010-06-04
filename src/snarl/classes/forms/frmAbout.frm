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
      Height          =   1095
      Index           =   4
      Left            =   1980
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
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
      Left            =   5520
      Picture         =   "frmAbout.frx":10ED
      Stretch         =   -1  'True
      Top             =   2460
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
      Picture         =   "frmAbout.frx":1D53
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
      Height          =   315
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
      Top             =   540
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
      Top             =   780
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

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long

End Type

Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO) As Boolean

Dim mSysKeyPrefs As Long
Dim mSysKeyTest As Long

Dim mTrayIcon As BNotifyIcon

Dim m_About As String
Dim mTaskbarCreated As Long
Dim m_SelectedApp As String         ' // current selected application in listbox

Dim mPrefs As T_CONFIG
Dim mCurAlert As TAlert

Dim mPanel As BPrefsPanel
Dim mAppsPage As TAppsPage

Dim mJSONSocket() As CJSONSocket
Dim mSockets As Long

Dim WithEvents ioSocket As CSocket
Attribute ioSocket.VB_VarHelpID = -1

Dim mListener() As CSnarlListener
Dim mListenerCount As Long

Dim mClickThruOver As CSnarlWindow
Dim mMenuOpen As Boolean

'Dim WithEvents myDownloadUpdateRequest As CHTTPRequest
Dim mDownloadId As Long

Private Type T_REMOTENOTIFICATION
    Token As Long               ' // notification token
    Socket As CSnarlSocket      ' // remote socket

End Type

Dim mRemoteNotification() As T_REMOTENOTIFICATION
Dim mRemoteNotifications As Long

Dim WithEvents theIdleTimer As BTimer
Attribute theIdleTimer.VB_VarHelpID = -1


Implements MMessageSink
Implements KPrefsPanel
Implements MWndProcSink

Private Sub Form_Load()
Dim sz As String
Dim pm As OMMenu
Dim n As Integer

'Dim pss As SNARLSTRUCT
'Dim pss2 As SNARLSTRUCT2
'
'    MsgBox LenB(pss) & " " & LenB(pss2)

    On Error Resume Next

    g_HideFromView Me.hWnd

    ' /* register the hotkeys */

    Me.bSetHotkeys

    ' /* pre-load our 'About' text */

    g_Debug "_load: pre-loading readme..."
    n = FreeFile()
    Err.Clear
    Open g_MakePath(App.Path) & "read-me.rtf" For Input As #n
    If Err.Number = 0 Then
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
'        .AddItem .CreateItem("dnd", "Do Not Disturb", , , g_Prefs.do_not_disturb)
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

    mTaskbarCreated = RegisterWindowMessage("TaskbarCreated")
    g_Debug "_load: 'TaskbarCreated' = " & g_HexStr(mTaskbarCreated, 4)

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

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Label1.Font.Underline = True Then _
        Label1.Font.Underline = False

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
            uNewDoPrefs

        ElseIf LoWord(wParam) = mSysKeyTest Then

            uDoSysInfoNotification

        Else
            g_Debug "ISubClassed.WndProc(): Spurious WM_HOTKEY received: " & _
                    g_HexStr(HiWord(wParam), 4) & " " & g_HexStr(LoWord(wParam), 4), LEMON_LEVEL_WARNING

        End If

        MWndProcSink_WndProc = True


    Case MSG_SHOW_PREFS
        uNewDoPrefs


    Case MSG_QUIT
'        Unload Me
        PostQuitMessage 0
        MWndProcSink_WndProc = True


    Case WM_SNARL_TRAY_ICON
        Select Case lParam
        Case WM_RBUTTONUP
            If Not fIgnoreNext Then
                uDoMainMenu

            Else
                fIgnoreNext = False

            End If

        Case WM_LBUTTONDBLCLK
'            fIgnoreNext = True
            uNewDoPrefs

        End Select

    Case WM_ENTERMENULOOP
        mMenuOpen = True

    Case WM_EXITMENULOOP
        mMenuOpen = False

    Case mTaskbarCreated
        g_Debug "frmAbout.WndProc(): 'TaskbarCreated' received - adding icon..."
        uAddTrayIcon

'    Case Is > WM_USER
'        Debug.Print ">> " & g_HexStr(uMsg, 4)

    Case WM_CLOSE
        Unload Me
        MWndProcSink_WndProc = True

    Case WM_INSTALL_SNARL
        If LoWord(wParam) = SNARL_NOTIFICATION_ACK Then _
            ShellExecute hWnd, "open", g_MakePath(App.Path) & gUpdateFilename, vbNullString, vbNullString, SW_SHOW

    Case WM_REMOTENOTIFY

        Debug.Print "WM_REMOTENOTIFY: token=" & lParam & " notification=" & LoWord(wParam)

        dw = uFindRemoteNotification(lParam)
        If dw = 0 Then
            g_Debug "WM_REMOTENOTIFY: token " & g_HexStr(lParam) & " not in remote list", LEMON_LEVEL_CRITICAL

        Else
            mRemoteNotification(dw).Socket.Notify LoWord(wParam), CStr(lParam)
            Select Case LoWord(wParam)
            Case SNARL_NOTIFICATION_ACK, SNARL_NOTIFICATION_CLOSED
                uRemoveRemoteNotification dw

            End Select

        End If

        MWndProcSink_WndProc = True

    End Select

End Function


Private Sub uDoMainMenu()
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
        .AddItem .CreateItem("sticky", "Sticky Notifications", , , g_IsSticky())
        .AddSeparator

        .AddItem .CreateItem("dnd", "Do Not Disturb", , , g_IsDNDModeEnabled())
        .AddItem .CreateItem("missed", "Missed Notifications...")
        
        .AddSeparator
        .AddItem .CreateItem("restart", "Restart Snarl", , g_IsRunning)

        If g_IsRunning Then
            .AddItem .CreateItem("stop", "Stop Snarl")

        Else
            .AddItem .CreateItem("start", "Start Snarl")

        End If
        
        .AddItem .CreateItem("quit", "Quit Snarl")
        .AddSeparator
        .AddItem .CreateItem("prefs", "Settings...")
'        .AddItem .CreateItem("app_mgr", "App Manager...")
'        .AddItem .CreateItem("", "Snarl Apps", , , , , , g_AppRoster.SnarlAppsMenu())
        .AddSeparator
        .AddItem .CreateItem("about", "About Snarl")

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
            uNewDoPrefs

        Case "app_mgr"
            ShellExecute 0, "open", g_MakePath(App.Path) & "Snarl_App_Manager.exe", vbNullString, vbNullString, SW_SHOW

        Case "sticky"
            g_Prefs.sticky_snarls = Not g_Prefs.sticky_snarls
            g_WriteConfig

        Case "dnd"
            g_Prefs.do_not_disturb_ = Not g_Prefs.do_not_disturb_
            g_WriteConfig

        Case "missed"
            If Not (g_NotificationRoster Is Nothing) Then _
                g_NotificationRoster.ShowMissedPanel

        Case Else
            sz = g_SafeLeftStr(pi.Name, 3)
            szData = g_SafeRightStr(pi.Name, Len(pi.Name) - 3)

            Select Case sz
            Case "cfg"
                ' /* Snarl App -> Settings... szData is App Roster index */
                g_AppRoster.SnarlAppDo Val(szData), SNARL_APP_SHOW_PREFS

            Case "abt"
                ' /* Snarl App -> About... szData is App Roster index */
                g_AppRoster.SnarlAppDo Val(szData), SNARL_APP_SHOW_ABOUT

            End Select

        End Select
    End If

    If update_config Then _
        g_WriteConfig

End Sub

Private Sub uNewDoPrefs()

    If Not (mPanel Is Nothing) Then _
        Exit Sub

Dim pbm As MImage

    Set mPanel = New BPrefsPanel
    mPanel.SetHandler Me
'    mPanel.SetMargin 90

    load_image g_MakePath(App.Path) & "etc\icons\general.png", pbm
    mPanel.AddPage new_BPrefsPage("General", pbm, New TGeneralPage)

    load_image g_MakePath(App.Path) & "etc\icons\apps.png", pbm
    Set mAppsPage = New TAppsPage
    mPanel.AddPage new_BPrefsPage("Apps", pbm, mAppsPage)

    load_image g_MakePath(App.Path) & "etc\icons\display.png", pbm
    mPanel.AddPage new_BPrefsPage("Display", pbm, New TDisplayPage)

    load_image g_MakePath(App.Path) & "etc\icons\styles.png", pbm
    mPanel.AddPage new_BPrefsPage("Styles", pbm, New TStylesPage)

    load_image g_MakePath(App.Path) & "etc\icons\extensions.png", pbm
    mPanel.AddPage new_BPrefsPage("Extensions", pbm, New TExtPage)

    load_image g_MakePath(App.Path) & "etc\icons\network.png", pbm
    mPanel.AddPage new_BPrefsPage("Network", pbm, New TNetworkPage)

    load_image g_MakePath(App.Path) & "etc\icons\advanced.png", pbm
    mPanel.AddPage new_BPrefsPage("Advanced", pbm, New TAdvancedPage)

    load_image g_MakePath(App.Path) & "etc\icons\about.png", pbm
    mPanel.AddPage new_BPrefsPage("About", pbm, New TAboutPage)

    If gDebugMode Then
        load_image g_MakePath(App.Path) & "etc\icons\debug.png", pbm
        mPanel.AddPage new_BPrefsPage("Debug", pbm, New TDebugPage)

    End If

    mPanel.SetTitle "Snarl Preferences"
    mPanel.SetWidth 500

    mPanel.Go
    g_SetWindowIconToAppResourceIcon mPanel.hWnd

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


    ' /* find our current style and select it in the 'Display' sub page */

Dim i As Long
Dim px As TStyle
Dim j As Long

'    Debug.Print g_Prefs.default_style

    If Not (g_StyleRoster Is Nothing) Then
        i = g_StyleRoster.IndexOf(style_GetStyleName(g_ConfigGet("default_style")))
        If i Then
            Set px = g_StyleRoster.StyleAt(i)
            j = px.SchemeIndex(style_GetSchemeName(g_ConfigGet("default_style")))

        End If

        If (i > 0) And (j > 0) Then
            If mPanel.Find("default_style", pc) Then _
                pc.SetValue CStr(i)
    
            If mPanel.Find("default_scheme", pc) Then _
                pc.SetValue CStr(j)

        End If
    End If

    If mPanel.Find("melontype_contrast", pc) Then _
        pc.SetEnabled (g_Prefs.font_smoothing = E_MELONTYPE)

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

Private Function MMessageSink_Received(Message As melon.MMessage) As Boolean

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

    If (mTrayIcon Is Nothing) Or (Not g_Prefs.show_tray_icon) Then _
        Exit Sub

    hIcon = LoadImage(App.hInstance, 1&, IMAGE_ICON, 16, 16, 0)
    If hIcon = 0 Then _
        hIcon = Me.Icon.Handle

    mTrayIcon.Add "tray_icon", hIcon, "Snarl"

End Sub

Private Function uIsAlertEnabled(ByVal ConfigString As String) As Boolean
Dim sz() As String

    On Error Resume Next

    sz() = Split(ConfigString, "#?")
    uIsAlertEnabled = Val(sz(0))

End Function

Private Sub ioSocket_OnConnect()

    g_Debug "ioSocket.OnConnect()", LEMON_LEVEL_PROC

End Sub

Private Sub ioSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "ioSocket.OnConnectionRequest(): requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

    mSockets = mSockets + 1
    ReDim Preserve mJSONSocket(mSockets)
    Set mJSONSocket(mSockets) = New CJSONSocket
    mJSONSocket(mSockets).Accept requestID

End Sub

Private Sub ioSocket_OnDataArrival(ByVal bytesTotal As Long)

    g_Debug "ioSocket.OnDataArrival(): bytesTotal=" & g_HexStr(bytesTotal), LEMON_LEVEL_PROC

End Sub






Friend Function bSetHotkeys(Optional ByVal KeyCode As Long = 0) As Boolean

    ' /* return True if the prefs hotkey was registered ok */

    If Not g_Prefs.use_hotkey Then
        ' /* hotkeys not enabled */
        uUnregisterHotkeys
        bSetHotkeys = True
        Exit Function

    End If

    If KeyCode = 0 Then
        g_Debug "bSetHotKeys(): registering existing hotkey (" & CStr(g_Prefs.hotkey_prefs) & ")", LEMON_LEVEL_INFO
        KeyCode = g_Prefs.hotkey_prefs

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
Dim dFreq As Double
Dim dw As Long

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

        snShowMessage g_GetUserName() & " on " & g_GetComputerName(), _
                      g_GetOSName() & " " & g_GetServicePackName() & vbCrLf & _
                      IIf(dw > 1, CStr(dw) & "x", "") & .FullName & " @ " & Format$(dFreq, "0.0#") & " " & szMetric & vbCrLf & _
                      g_FileSizeToStringEx2(g_GetPhysMem(True), "GB", " ", "0.0#") & " physical, " & g_FileSizeToStringEx2(g_GetPageMem(True) + g_GetPhysMem(True), "GB", " ", "0.##") & " total RAM" & vbCrLf & _
                      "Snarl " & App.Major & "." & App.Revision & ", exec " & melonGetVersion(libmexec.globals), _
                      20, _
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

    ' /* ignore if no idle timeout set */

    If g_Prefs.idle_timeout <= 0 Then _
        Exit Sub

Dim lii As LASTINPUTINFO

    lii.cbSize = Len(lii)
    If GetLastInputInfo(lii) = False Then _
        Exit Sub

    lii.dwTime = GetTickCount() - lii.dwTime
'    g_Debug "_theIdleTimer.Pulse(): idle time is now " & CStr(lii.dwTime) & " needs to be " & CStr(g_Prefs.idle_timeout * 1000)

Static b As Boolean

    b = (lii.dwTime > (g_Prefs.idle_timeout * 1000))
    If b <> gIsIdle Then
        g_Debug "_theIdleTimer.Pulse(): " & Now() & " away mode: " & b
        gIsIdle = b

    End If

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

Public Sub AddRemoteNotification(ByVal Token As Long, ByRef Socket As CSnarlSocket)

    mRemoteNotifications = mRemoteNotifications + 1
    ReDim Preserve mRemoteNotification(mRemoteNotifications)
    With mRemoteNotification(mRemoteNotifications)
        .Token = Token
        Set .Socket = Socket

    End With

End Sub

Private Function uFindRemoteNotification(ByVal Token As Long) As Long
Dim i As Long

    If mRemoteNotifications = 0 Then _
        Exit Function

    For i = 1 To mRemoteNotifications
        If mRemoteNotification(i).Token = Token Then
           uFindRemoteNotification = i
           Exit Function

        End If

    Next i

End Function

Private Sub uRemoveRemoteNotification(ByVal Index As Long)
Dim i As Long

    If (Index < 1) Or (Index > mRemoteNotifications) Then _
        Exit Sub

    If Index < mRemoteNotifications Then
        For i = Index To (mRemoteNotifications - 1)
            LSet mRemoteNotification(i) = mRemoteNotification(i + 1)

        Next i

    End If

    mRemoteNotifications = mRemoteNotifications - 1
    ReDim Preserve mRemoteNotification(mRemoteNotifications)

End Sub

Public Sub EnableJSON(ByVal Enabled As Boolean)

    g_Debug "frmAbout.EnableJSON(" & CStr(Enabled) & ")", LEMON_LEVEL_PROC_ENTER

    If Enabled Then
        g_Debug "creating JSON listener..."

        Set ioSocket = New CSocket
        ioSocket.Bind "9889", "127.0.0.1"
        ioSocket.Listen

        g_Debug "listening on " & ioSocket.LocalIP & ":" & ioSocket.LocalPort & "..."

    Else
        g_Debug "stopping JSON listener..."
        If Not (ioSocket Is Nothing) Then
            ioSocket.CloseSocket
            Set ioSocket = Nothing
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
            For i = 0 To UBound(szAddr())
                If szAddr(i) <> "0.0.0.0" Then
                    mListenerCount = mListenerCount + 1
                    ReDim Preserve mListener(mListenerCount)
                    Set mListener(mListenerCount) = New CSnarlListener
                    mListener(mListenerCount).Go szAddr(i)

                End If
            Next i

        Else
            g_Debug "couldn't read local ip address table", LEMON_LEVEL_WARNING

        End If

    Else
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
