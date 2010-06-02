VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnarlWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const OFFSET_ICON = 0

Private Const TIMER_MAIN_TICK = 22
Private Const TIMER_ASYNC_FADE = 23
Private Const TIMER_TRACK_POINTER = 24
Private Const TIMER_PULSE = 25

Dim theView As mfxView
Dim theWindow As NitroWindow

Dim m_Alpha As Integer
Dim mOriginalTimeout As Long            ' // in seconds!

Dim mInfo As T_NOTIFICATION_INFO

Dim m_Current As Long
Dim m_Target As Long
Dim m_step As Long

    ' /* 37.33 */
Dim m_Disable As Long                   ' // >0 = timer paused, <=0 timer running
    ' /* 37.80
Dim m_IsVisible As Boolean

    ' /* 39.62 */
Dim mInstance As IStyleInstance

    ' /* 38.9 */
Dim theBitmap As mfxBitmap
Dim mCloseButtonVisible As Boolean

    ' /* 38.132 */
Dim mMenu As OMMenu                     ' // set by calling snChangeAttribute(SNARL_ATTRIBUTE_MENU)

    ' /* R2.2 */
Dim mStyleFlags As S_STYLE_FLAGS

Dim mFuzzy As Boolean
Dim mQuitting As Boolean

Dim mTick As Long                       '// timeout/100

    ' /* R2.3 */
Dim mClickThru As Boolean               '// set during Create() - taken from class settings and then modified by style flags
Dim mStylePath As String
Dim mClass As TAlert

Implements MWndProcSink

Friend Function Create(ByRef Class As TAlert, ByRef Info As T_NOTIFICATION_INFO, ByRef Instance As IStyleInstance, ByVal StyleFlags As S_STYLE_FLAGS, ByVal StylePath As String) As M_RESULT

    If (Class Is Nothing) Then
        Create = M_INVALID_ARGS
        Exit Function

    End If

    mClickThru = Class.IsClickThru
    mStylePath = StylePath
    Set mClass = Class

    LSet mInfo = Info

Dim pbm As mfxBitmap
Dim rc As RECT
Dim i As Long

    m_Alpha = 255 * (g_Prefs.global_opacity / 100)
    mOriginalTimeout = IIf(g_Prefs.sticky_snarls, 0, Info.Timeout)
    Set mInstance = Instance
    mStyleFlags = StyleFlags

    ' /* R2.2: styles can (now) also request click-through notifications */

    If (StyleFlags And S_STYLE_CLICK_THROUGH) Then _
        mClickThru = True

    ' /* ask our style instance to update its content */

    uUpdateStyleContent Info

Dim dwExStyle As Long

    dwExStyle = WS_EX_TOPMOST
    If mClickThru Then
        dwExStyle = dwExStyle Or WS_EX_TRANSPARENT
        If mOriginalTimeout = 0 Then
            g_Debug "CSnarlWindow.Create(): timeout cannot be zero if E_CLASS_NOTIFICATION_ONLY is set", LEMON_LEVEL_WARNING
            mOriginalTimeout = IIf(g_Prefs.default_duration > 0, g_Prefs.default_duration, 10)

        End If
    End If

    If nto_CreateWindow(theWindow, Me, N_POPUP_WINDOW, , , , dwExStyle) = M_OK Then
        Set theView = New mfxView

        ' /* configure our size from the style */

        Set theBitmap = uGetContent2()
        If Not (theBitmap Is Nothing) Then
            theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON
'            theView.SizeTo theBitmap.Width + 0, theBitmap.Height + 0

            mCloseButtonVisible = False
            uDrawContent 0
            theWindow.MakeVisible

            If (StyleFlags And S_STYLE_PULSE_NEEDED) Then _
                SetTimer theWindow.hWnd, TIMER_PULSE, 100, 0

        Else
            g_Debug "CSnarlWindow.Create(): returned style bitmap was NULL", LEMON_LEVEL_CRITICAL

        End If
    End If

End Function

Public Property Get Window() As NitroWindow

    Set Window = theWindow

End Property

Public Sub Show()
Dim bDone As Boolean
Dim pt As POINTAPI

    If Not (mInstance Is Nothing) Then
        If (mStyleFlags And S_STYLE_CUSTOM_SHOW) Then

            mInstance.Show True
            m_Alpha = 255
            m_Current = 255

            Do
                bDone = True
                mInstance.AdjustPosition pt.x, pt.y, m_Alpha, bDone

                ' /* if S_STYLE_WILL_RESIZE flag is set, ask the instance for a new bitmap each time */

                If (mStyleFlags And S_STYLE_WILL_RESIZE) Then
                    Set theBitmap = uGetContent2()
                    theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON

                End If

                uDrawContent m_Alpha, pt.x, pt.y

                ' /* keep going until the instance leaves bDone as True */

            Loop While Not bDone

            Me.SetTimeout mOriginalTimeout
            
Debug.Print "TIMEOUT " & mOriginalTimeout
            
            m_IsVisible = True

            Exit Sub

        End If

    End If

Debug.Print "SHOW: " & m_Alpha

    If Not (theWindow Is Nothing) Then
        uSetTargetAlpha m_Alpha

        Me.SetTimeout mOriginalTimeout

Debug.Print "TIMEOUT " & mOriginalTimeout

'        If m_Timeout > 0 Then _
            theWindow.AddTimer TIMER_MAIN_TICK, m_Timeout * 1000

        m_IsVisible = True

    End If

End Sub

Public Sub Quit()
'Static bBusy As Boolean
Dim i As Long

'    If bBusy Or mQuitting Then _
        Exit Sub

    If mQuitting Then _
        Exit Sub

    m_IsVisible = False
'    bBusy = True

    If Not (theWindow Is Nothing) Then
        mQuitting = True                                    '// set this notification as no longer valid...
        KillTimer theWindow.hWnd, TIMER_TRACK_POINTER
'        KillTimer theWindow.hWnd, TIMER_PULSE
        theWindow.AddExStyles WS_EX_TRANSPARENT             '// window becomes click-through while fading

        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.Remove mInfo.Token         '// remove from roster so space can be reused

        uSetTargetAlpha 0                                   '// async fade to zero
'        Do While m_step <> 0
'            DoEvents
'            Sleep 1
'            DoEvents
'
'        Loop

    End If

'    bBusy = False

End Sub

Public Sub Zap()

    m_step = 0
    nto_DestroyWindow theWindow
    Set theView = Nothing

End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static bPaused As Boolean
Static pt As POINTAPI
Static sz As String
Static f As Boolean

    Select Case uMsg
    Case WM_SETCURSOR
        SetCursor LoadCursor(0, IDC_ARROW)

        If LoWord(lParam) = HTERROR Then
            ' /* window is disabled which means a menu is active */
            Select Case HiWord(lParam)
            Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
                Debug.Print "-- enabling --"
                ' /* enable the window but don't process the message */
                EnableWindow hWnd, -1
                ReturnValue = -1
                MWndProcSink_WndProc = True

            End Select
        End If





    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
        ' /* V38.131: better way of working - button pressed now does a SetCapture() and we only respond
        '    if the mouse button is released over the same item */
        sz = uHitTest(lParam)
        f = True
        SetCapture hWnd

    Case WM_LBUTTONUP
        If f Then
            ReleaseCapture
            If sz = uHitTest(lParam) Then
                Select Case sz
                Case "close"
                    g_Debug "CSnarlWindow.WndProc(): close button invoked - not sending ACK"
                    If (g_IsPressed(VK_LCONTROL)) Or (g_IsPressed(VK_RCONTROL)) Then
                        ' /* CTRL held down... */
    
    '                    g_Debug "CTRL " & g_HexStr(m_hWndReply)
    
                        If (g_IsPressed(VK_LSHIFT)) Or (g_IsPressed(VK_RSHIFT)) Then
                            ' /* SHIFT held down: close all */
                            g_NotificationRoster.CloseMultiple 0
    
    '                        g_Debug "SHIFT"
    
                        ElseIf mInfo.hWndReply <> 0 Then
                            ' /* just from the provided window */
                            g_NotificationRoster.CloseMultiple mInfo.hWndReply
    
                        Else
                            ' /* no reply window so just close ourselves */
                            Hide
    
                        End If
    
                    Else
                        ' /* just close ourself */
                        bPaused = True                  ' // don't count down while we're doing this...
                        uNotify SNARL_NOTIFICATION_CLOSED
                        bPaused = False
                        Hide

                    End If
    
                Case "client"
                    g_Debug "CSnarlWindow.WndProc(): client area invoked"
                    bPaused = True                  ' // don't count down while we're doing this...
    
                    If mInfo.DefaultAck <> "" Then
                        ' /* V38.94 - if an ACK string has been specified, launch it now */
                        g_Debug "CSnarlWindow.WndProc(): launching ACK '" & mInfo.DefaultAck & "'...", LEMON_LEVEL_INFO
                        ShellExecute frmAbout.hWnd, vbNullString, mInfo.DefaultAck, vbNullString, vbNullString, SW_SHOW
    
                    End If

                    uNotify SNARL_NOTIFICATION_ACK
                    bPaused = False
                    Hide
    
                End Select
            Else
                g_Debug "CSnarlWindow.WndProc(): WM_LBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_LBUTTONUP ignored: outside window"

        End If

        sz = ""


    Case WM_RBUTTONUP
        If f Then
            ReleaseCapture
            If sz = uHitTest(lParam) Then
                Select Case sz
                Case "close"
                    g_Debug "CSnarlWindow.WndProc(): close button invoked (right-click) - not sending CLICKED"
                    bPaused = True
                    uDoSysMenu
                    bPaused = False
    
                Case "client"
                    bPaused = True
                    If Not (mMenu Is Nothing) Then
                        ' /* track menu */
                        uTrackMenu hWnd

                    Else
                        uNotify SNARL_NOTIFICATION_CLICKED

                    End If
                    bPaused = False
    
                End Select
            Else
                g_Debug "CSnarlWindow.WndProc(): WM_RBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_RBUTTONUP ignored: outside window"

        End If
        sz = ""


    Case WM_MBUTTONUP
        ' /* 38.133 - support the middle button */
        If f Then
            ReleaseCapture
            If sz = uHitTest(lParam) Then
                Select Case sz
                Case "close"
    
                Case "client"
                    bPaused = True
                    uNotify SNARL_NOTIFICATION_MIDDLE_BUTTON
                    bPaused = False

                End Select
            Else
                g_Debug "CSnarlWindow.WndProc(): WM_MBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_MBUTTONUP ignored: outside window"

        End If
        sz = ""


    Case WM_TIMER
        If (wParam = TIMER_MAIN_TICK) And (Not bPaused) And (Not mCloseButtonVisible) Then
            mTick = mTick - 1
            If mTick = 0 Then
                
Debug.Print "SNARL_NOTIFICATION_TIMED_OUT " & mTick
                
                bPaused = True
                uNotify SNARL_NOTIFICATION_TIMED_OUT
                bPaused = False
                Hide

            End If

        ElseIf wParam = TIMER_ASYNC_FADE Then
            If m_step < 0 Then
                ' /* fading down */
                
'Debug.Print "FADING DOWN"
                
                If m_Current <= m_Target Then
                    ' /* done! */
                    m_Current = m_Target
                    KillTimer hWnd, wParam
'                    g_NotificationRoster.Remove m_Id
                    m_step = 0

                    If mQuitting Then
                        KillTimer hWnd, TIMER_PULSE
                        nto_DestroyWindow theWindow
                        Set theView = Nothing

                    End If
                    Exit Function
                
                End If

            Else
                ' /* fading up */
                If m_Current >= m_Target Then
                    ' /* done! */
                    m_Current = m_Target
                    KillTimer hWnd, wParam
                    m_step = 0
                    Exit Function

                End If

            End If

            ' /* keep fading */
            m_Current = m_Current + m_step
            apply_view_to_window theView, hWnd, m_Current


        ElseIf (wParam = TIMER_TRACK_POINTER) And (Not f) Then
            GetCursorPos pt
            If Not theWindow.Frame.Contains(nto_NewPoint(pt.x, pt.y)) Then
                KillTimer hWnd, wParam
                mCloseButtonVisible = False
                uDrawContent m_Current

            End If

        ElseIf wParam = TIMER_PULSE Then
            uPulse

        End If

    Case WM_NCHITTEST
        mCloseButtonVisible = True
        uDrawContent m_Current
        SetTimer hWnd, TIMER_TRACK_POINTER, 100, 0

    Case WM_CAPTURECHANGED
        f = False


    Case WM_NCACTIVATE
        If wParam = 0 Then _
            EnableWindow hWnd, -1


'    Case WM_EXITMENULOOP
'        PostMessage hWnd, &H666, 0, ByVal 0&
'        GetCursorPos pt
'        If Not theWindow.Frame.Contains(nto_NewPoint(pt.x, pt.y)) Then _
'            EnableWindow hWnd, -1



'        Debug.Print "-- exit menu --"
'        EnableWindow hWnd, -1

'    Case WM_KEYDOWN
'        Debug.Print wParam & " " & g_HexStr(lParam)

    End Select

End Function

Private Function uHitTest(ByVal lParam As Long) As String
Dim pt As POINTAPI
Dim rc As RECT
Dim ps As SIZEL
Dim n As Integer

    MAKEPOINTS lParam, pt.x, pt.y
    GetClientRect theWindow.hWnd, rc

    If uDrawDropshadow() Then _
        n = g_Prefs.dropshadow_size

    If (pt.x >= rc.Left) And (pt.x <= rc.Right) And (pt.y >= rc.Top) And (pt.y <= rc.Bottom) Then
        ' /* returns either "close" or "client" depending on where the cursor is */

        If (pt.x >= n) And (pt.y >= n) And (pt.x <= (n + bm_Close.Width)) And (pt.y <= (n + bm_Close.Height)) Then
            uHitTest = "close"

        Else
            uHitTest = "client"

        End If

        Debug.Print uHitTest

    Else
        g_Debug "CSnarlWindow(): not inside window"

    End If

End Function

Public Sub Hide()
Dim bDone As Boolean
Dim pt As POINTAPI

    If Not (mInstance Is Nothing) Then
        If (mStyleFlags And S_STYLE_CUSTOM_HIDE) Then
            m_IsVisible = False
            mQuitting = True

            With theWindow
                .RemoveTimer TIMER_MAIN_TICK
                .RemoveTimer TIMER_TRACK_POINTER
                .AddExStyles WS_EX_TRANSPARENT          '// window becomes click-through while hiding

            End With

            g_NotificationRoster.Remove mInfo.Token     '// remove from roster so space can be reused
            mInstance.Show False

            Do
                bDone = True
                mInstance.AdjustPosition pt.x, pt.y, m_Alpha, bDone

                If (mStyleFlags And S_STYLE_WILL_RESIZE) Then
                    Set theBitmap = uGetContent2()
                    theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON

                End If
                uDrawContent m_Alpha, pt.x, pt.y

            Loop While Not bDone

            nto_DestroyWindow theWindow
            Set theView = Nothing

            Exit Sub

        End If

    End If

    theWindow.RemoveTimer TIMER_MAIN_TICK
'    g_RemoveSnarl theWindow.hWnd
    Quit

End Sub

Public Property Get Id() As Long

    Id = mInfo.Token

End Property

Friend Function Update(ByRef Info As T_NOTIFICATION_INFO) As M_RESULT

    ' /* tell the style handler to update its content */

    uUpdateStyleContent Info

    ' /* view is already set - don't bother resizing */

    Set theBitmap = uGetContent2()
    uDrawContent m_Current

    SetTimeout mOriginalTimeout                 ' // V37.80 - reset timeout value
    Update = M_OK

End Function

Public Sub uSetTargetAlpha(ByVal Alpha As Long)

Debug.Print "SET TARGET ALPHA " & Alpha

    If Alpha < m_Current Then
        m_step = -Val(g_ConfigGet("step_size")) * 8

    ElseIf Alpha > m_Current Then
        m_step = Val(g_ConfigGet("step_size")) * 4

    Else
        Exit Sub

    End If

    m_Target = Alpha
    SetTimer theWindow.hWnd, TIMER_ASYNC_FADE, 2, 0

End Sub

Public Sub Enable()

    m_Disable = m_Disable - 1

End Sub

Public Sub Disable()

    m_Disable = m_Disable + 1

End Sub

Public Function SetTimeout(ByVal Timeout As Long) As M_RESULT

Debug.Print "SETTIMEOUT"

    SetTimeout = M_FAILED
    If (theWindow Is Nothing) Then _
        Exit Function

    SetTimeout = M_INVALID_ARGS
    If (Timeout < 0) Or (Timeout > 65535) Then _
        Exit Function

    SetTimeout = M_OK
    theWindow.RemoveTimer TIMER_MAIN_TICK

    If Timeout = 0 Then _
        Exit Function

    mOriginalTimeout = Timeout
    mTick = Timeout * 10                    ' // timer frequency is 100ms...

    theWindow.AddTimer TIMER_MAIN_TICK, 100

End Function

Public Property Get IsVisible() As Boolean

    IsVisible = m_IsVisible

End Property

Public Function GetView() As mfxView

    Set GetView = theView

End Function

Private Sub uDrawContent(Optional ByVal Alpha As Long = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)

    If (theView Is Nothing) Or (theBitmap Is Nothing) Or (theWindow Is Nothing) Then _
        Exit Sub

'Dim pr As NRect

    If Alpha = -1 Then _
        Alpha = m_Current

    If uDrawDropshadow() Then
        If (x = -1) And (y = -1) Then
'            Set pr = theWindow.Frame
'            x = pr.Left - 10
'            y = pr.Top - 10

        Else
            ' /* specific position requested and dropshadow enabled */
            x = x - g_Prefs.dropshadow_size
            y = y - g_Prefs.dropshadow_size

        End If
    End If

    With theView
        .Clear

'        .SetHighColour rgba(255, 0, 0, 190)
'        .FillRect .Bounds

        .DrawScaledImage theBitmap, new_BPoint(OFFSET_ICON, OFFSET_ICON)
        If (mCloseButtonVisible) And (Not mClickThru) Then _
            .DrawScaledImage bm_Close, IIf(uDrawDropshadow(), new_BPoint(g_Prefs.dropshadow_size, g_Prefs.dropshadow_size), new_BPoint(0, 0))

    End With

    apply_view_to_window theView, theWindow.hWnd, Alpha, x, y

End Sub

Friend Sub bSetAlpha(ByVal Alpha As Integer)

    If Alpha < 0 Then
        Alpha = 0

    ElseIf Alpha > 255 Then
        Alpha = 255

    End If

    m_Current = Alpha
    uDrawContent m_Current

End Sub

Private Sub uDoSysMenu()
Dim pmi As OMMenuItem
Dim pm As OMMenu

    With New OMMenu
        ' /* R2.2 */
        .AddItem .CreateItem("stky", "Make Sticky", , (mOriginalTimeout > 0))
        .AddSeparator

        .AddItem .CreateItem("wlog", "Write content to Log")
        .AddItem .CreateItem("copy", "Copy content to Clipboard")

        Set pm = g_GetRemoteComputersMenu()
        .AddItem .CreateItem(" ", "Send Content to", , Not (pm Is Nothing), , , , pm)

        .AddSeparator

        .AddItem .CreateItem("clse", "Close")
        .AddItem .CreateItem("capp", "Close all from " & mClass.App.Name)
        .AddItem .CreateItem("call", "Close all")

        If g_IsPressed(VK_CONTROL) Then
            .AddSeparator
            .AddItem .CreateItem("scrn", "Capture Notification")

        End If

        EnableWindow theWindow.hWnd, 0

        Set pmi = .Track(theWindow.hWnd)

    End With

    If Not (pmi Is Nothing) Then
        EnableWindow theWindow.hWnd, -1         ' // item was selected so re-enable the window
        
        Select Case pmi.Name
        Case "clse"
            Hide

        Case "capp"
            g_NotificationRoster.CloseMultiple mClass.App.Token

        Case "call"
            g_NotificationRoster.CloseMultiple 0

        Case "scrn"
            uCapture

        Case "stky"
            Me.SetTimeout 0

        Case "wlog"
            g_WriteToLog mInfo.Title, mInfo.Text

        Case "copy"
            Clipboard.SetText mInfo.Title & vbCrLf & mInfo.Text

        Case Else
            If g_SafeLeftStr(pmi.Name, 3) = "ip>" Then
'                MsgBox g_SafeRightStr(pmi.Name, Len(pmi.Name) - 3)

'                With New SNBConnection
'                    If .Connect(g_SafeRightStr(pmi.Name, Len(pmi.Name) - 3)) = SNB_SUCCESS Then
'                        .ShowNotification "", mInfo.Title, mInfo.Text & vbCrLf & vbCrLf & "Sent from " & get_host_name(), mOriginalTimeout
'                        .Disconnect
'
'                    End If
'
'                End With

            End If

        End Select

    End If

End Sub

Public Function SetAck(ByVal Value As String) As M_RESULT

    mInfo.DefaultAck = Value
    SetAck = M_OK

End Function

Private Sub uCapture()
Dim szPath As String
Dim sz As String
Dim c As Long

    If Not g_GetSystemFolder(CSIDL_DESKTOP, szPath) Then _
        Exit Sub

    szPath = g_MakePath(szPath)

    c = 1
    sz = szPath & "snarl_" & CStr(c) & ".png"
    Do While g_Exists(sz)
        c = c + 1
        If c > 65535 Then _
            Exit Sub

        sz = szPath & "snarl_" & CStr(c) & ".png"

    Loop

    mCloseButtonVisible = False
    uDrawContent m_Current

    CaptureToFile sz

    mCloseButtonVisible = True
    uDrawContent m_Current

End Sub

Public Function CaptureToFile(ByVal Filename As String)

    If Not (theView Is Nothing) Then _
        theView.WriteToFile Filename, "image/png"

End Function

Public Function SetMenu(ByVal Value As String) As M_RESULT

    On Error Resume Next

    Set mMenu = Nothing
    SetMenu = M_OK

    If Value = "" Then _
        Exit Function

Dim sz() As String
Dim i As Long

    sz() = Split(Value, "|")
    Debug.Print "$$%: " & UBound(sz())

    If UBound(sz()) = -1 Then _
        Exit Function

    Set mMenu = New OMMenu

    For i = 0 To UBound(sz())
        uAddToMenu mMenu, sz(i)

    Next i

End Function

Private Sub uAddToMenu(ByRef Menu As OMMenu, ByVal Item As String)
Dim szText As String
Dim szData As String
Dim dw As Long
Dim i As Integer

    If Item = "" Then
        Menu.AddSeparator
        Exit Sub

    End If

    i = InStr(Item, "#?")
    If i = 0 Then
        szText = Item

    Else
        szText = g_SafeLeftStr(Item, i - 1)
        szData = g_SafeRightStr(Item, Len(Item) - i - 1)
        If Not g_IsNumeric(szData) Then
            g_Debug "CSnarlWindow.uAddToMenu(): can't add '" & Item & "' to menu: data part is not numeric", LEMON_LEVEL_CRITICAL
            Exit Sub

        End If

        dw = Val(szData)
        If (dw < 1) Or (dw > 65535) Then
            g_Debug "CSnarlWindow.uAddToMenu(): can't add '" & Item & "' to menu: data part must be 1-65535", LEMON_LEVEL_CRITICAL
            Exit Sub

        End If

    End If

    Menu.AddItem mMenu.CreateItem(szData, szText)

End Sub

Private Sub uTrackMenu(ByVal hWndOwner As Long)

    If (mMenu Is Nothing) Then _
        Exit Sub

Dim pmi As OMMenuItem

    EnableWindow hWndOwner, 0
    Set pmi = mMenu.Track(hWndOwner)
    If (pmi Is Nothing) Then _
        Exit Sub

    EnableWindow hWndOwner, -1          ' // item was selected so re-enable the window

'Debug.Print "@ " & m_hWndReply & " / " & m_uMsg

    If (IsWindow(mInfo.hWndReply) = 0) Or (mInfo.uReplyMsg = 0) Then
        g_Debug "CSnarlWindow.uTrackMenu(): can't send notification: invalid window or reply message", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim dw As Long

    SendMessageTimeout mInfo.hWndReply, mInfo.uReplyMsg, MAKELONG(SNARL_NOTIFICATION_MENU, Val(pmi.Name)), ByVal mInfo.Token, SMTO_ABORTIFHUNG, 500, dw

End Sub

Private Sub uNotify(ByVal Notification As Long)
Dim dw As Long

    If (mInfo.hWndReply) And (mInfo.uReplyMsg) Then
        g_Debug "CSnarlWindow.uNotify(): sending #" & CStr(Notification) & " to hWnd=" & g_HexStr(mInfo.hWndReply) & " msg=" & g_HexStr(mInfo.uReplyMsg) & " id=" & g_HexStr(mInfo.Token), LEMON_LEVEL_INFO
        SendMessageTimeout mInfo.hWndReply, mInfo.uReplyMsg, Notification, ByVal mInfo.Token, SMTO_ABORTIFHUNG, 500, dw

    Else
        g_Debug "CSnarlWindow.uNotify(): can't send #" & CStr(Notification) & ": hWnd=" & g_HexStr(mInfo.hWndReply) & " msg=" & g_HexStr(mInfo.uReplyMsg) & " id=" & g_HexStr(mInfo.Token), LEMON_LEVEL_WARNING

    End If

End Sub

' /*---------------------------------------------------------------------------------------
'   Update2() -- intelligent update
'
'   Called by the notification roster for single-instance, non-queued notifications when
'   a new notification request arrives and an existing notification is still on-screen.
'
'   This function updates the existing notification's content, then asks the style instance
'   for an updated position (the style instance can resize the content during the update
'   request) and re-positions it on screen at the requested location and size.
'
' ---------------------------------------------------------------------------------------*/

Friend Sub Update2(ByRef Info As T_NOTIFICATION_INFO)

    ' /* ask our style instance to update */

    uUpdateStyleContent Info

    ' /* get the new content */

Dim pb As mfxBitmap

    Set pb = uGetContent2()
    If (pb Is Nothing) Then _
        Exit Sub

    ' /* ask the instance if it wants to change the notification's position */

Dim pt As POINTAPI

    pt.x = theWindow.Frame.Left
    pt.y = theWindow.Frame.Top
    Me.AdjustPosition pt.x, pt.y

    ' /* from here this is generic to both instance types */

    SetTimeout mOriginalTimeout                 ' // reset timeout value

    ' /* re-position and re-size it */

    Set theBitmap = pb
    theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON

    ' /* make it visible */

    g_MoveWindow theWindow.hWnd, pt.x, pt.y
    uDrawContent m_Current

End Sub

Public Function NotificationOnlyMode() As Boolean

    NotificationOnlyMode = mClickThru

End Function

Public Sub MakeFuzzy(ByVal Enabled As Boolean)

    If mFuzzy = Enabled Then _
        Exit Sub

Static iPrevAlpha As Long
Static bmOld As mfxBitmap

'Dim pv As mfxView

    If Enabled Then

        Set bmOld = theBitmap.Duplicate
        theBitmap.MakeGreyscale
        uDrawContent

        iPrevAlpha = m_Current
        uSetTargetAlpha MIN(102, Fix(m_Current / 2))

'        uFadeTo MIN(102, Fix(m_Current / 2))
'
'
''        bSetAlpha MIN(60, Fix(m_Current / 2))
'
''        Set pv = New mfxView
''        pv.SizeTo theView.Width * 4, theView.Height * 4
''        pv.DrawScaledImage theView.ConvertToBitmap, , create_point(pv.Width, pv.Height)
''
''        theView.Clear
''        theView.DrawScaledImage pv.ConvertToBitmap, , create_point(theView.Width, theView.Height)

    Else

        Set theBitmap = bmOld.Duplicate
        Set bmOld = Nothing
        uDrawContent

        Debug.Print "resetting to " & iPrevAlpha
'        'bSetAlpha iPrevAlpha
'        uFadeTo iPrevAlpha

        uSetTargetAlpha iPrevAlpha

    End If

    mFuzzy = Enabled

End Sub

Private Sub uFadeTo(ByVal Alpha As Long)
Dim i As Long

    If Alpha < 0 Then
        Alpha = 0

    ElseIf Alpha > 255 Then
        Alpha = 255

    End If

    If m_Current > Alpha Then
        ' /* fade down */
        For i = m_Current To Alpha Step -3
            uDrawContent i

        Next i

    ElseIf m_Current < Alpha Then
        ' /* fade up */
        For i = m_Current To Alpha Step 3
            uDrawContent i

        Next i

    End If

    m_Current = Alpha

End Sub

Private Function uGetContent2() As mfxBitmap

    If (mInstance Is Nothing) Then _
        Exit Function

Dim pbm As mfxBitmap

    Set pbm = create_bitmap_from_image(mInstance.GetContent())
    If (pbm Is Nothing) Then _
        Exit Function

Dim pv As mfxView

    If uDrawDropshadow() Then
        Set pv = New mfxView
        pv.SizeTo pbm.Width, pbm.Height
        pv.DrawScaledImage pbm
        Set uGetContent2 = create_dropshadow(pv, , g_Prefs.dropshadow_size * 2, Fix((g_Prefs.dropshadow_strength / 100) * 255))

    Else
        Set uGetContent2 = pbm

    End If

End Function

Private Sub uUpdateStyleContent(ByRef Info As T_NOTIFICATION_INFO)

    On Error GoTo er

    If (mInstance Is Nothing) Then _
        Exit Sub

    With Info

        ' /* translate special markers */

        .Title = Replace$(.Title, "\n", vbCrLf)
        .Text = Replace$(.Text, "\n", vbCrLf)

        ' /* take into consideration and "do not change" markers */
        
        If .Title = "-" Then _
            .Title = mInfo.Title

        If .Text = "-" Then _
            .Text = mInfo.Text

        If .IconPath = "-" Then _
            .IconPath = mInfo.IconPath

        ' /* this actually only contains the scheme name */

        If .StyleToUse = "-" Then _
            .StyleToUse = mInfo.StyleToUse

    End With

    ' /* store the new values */

    With mInfo
        .Title = Info.Title
        .Text = Info.Text
        .IconPath = Info.IconPath
        .StyleToUse = Info.StyleToUse

    End With

    ' /* translate into a notification_info struct that the style type lib can understand */

Dim pInfo As notification_info

    With pInfo
        .Title = Info.Title
        .Text = Info.Text
        .Icon = Info.IconPath
        .Scheme = Info.StyleToUse

        If Info.Priority > 0 Then _
            .Flags = .Flags Or S_NOTIFICATION_IS_PRIORITY

    End With

    mInstance.UpdateContent pInfo
    Exit Sub

er:
    g_Debug "CSnarlWindow.uUpdateStyleContent(): " & Err.Description, LEMON_LEVEL_CRITICAL

End Sub

Public Sub AdjustPosition(ByRef x As Long, ByRef y As Long)

    If uDrawDropshadow() Then
        x = x + g_Prefs.dropshadow_size
        y = y + g_Prefs.dropshadow_size

    End If

    mInstance.AdjustPosition x, y, 255, True

    ' /* compensate for dropshadow */

    If uDrawDropshadow() Then
        x = x - g_Prefs.dropshadow_size
        y = y - g_Prefs.dropshadow_size

    End If

End Sub

Private Function uDrawDropshadow() As Boolean

    ' /* styles can veto the system drop shadow */

    uDrawDropshadow = (g_Prefs.use_dropshadow) And ((mStyleFlags And S_STYLE_NO_DROPSHADOW) = 0)

End Function

Private Sub uPulse()

    On Error GoTo ex

    If (mInstance Is Nothing) Then _
        Exit Sub

    ' /* only update if calling Pulse() returns True */

    If Not mInstance.Pulse() Then _
        Exit Sub

    Set theBitmap = uGetContent2()
    uDrawContent
    Exit Sub

ex: g_Debug "CSnarlWindow.uPulse(): " & Err.Description

End Sub

Public Function StylePath() As String

    StylePath = mStylePath

End Function

Public Function Frame() As BRect

    If Not (theWindow Is Nothing) Then _
        Set Frame = new_BRectFromRect(theWindow.Frame)

End Function

