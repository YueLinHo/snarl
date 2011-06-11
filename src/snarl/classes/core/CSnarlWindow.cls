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

Dim mCurrent As Long
Dim mTargetAlpha As Long
Dim mAlphaStep As Single

    ' /* 37.33 */
Dim m_Disable As Long                   ' // >0 = timer paused, <=0 timer running
    ' /* 37.80
Dim mVisible As Boolean

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
Dim mMenuOpen As Boolean

    ' /* R2.4 */

Dim mCloseRect As BRect
Dim mActionsRect As BRect

Dim mFadeStart As Long
Dim mOriginalAlpha As Long

Implements MWndProcSink

Friend Function Create(ByRef Info As T_NOTIFICATION_INFO, ByRef Instance As IStyleInstance, ByVal StyleFlags As S_STYLE_FLAGS, ByVal StylePath As String) As M_RESULT

    ' /* must have a class */

    If (Info.ClassObj Is Nothing) Then
        Create = M_INVALID_ARGS
        Exit Function

    End If

'    Info.StyleToUse = style_GetSchemeName(Info.StyleToUse)

    LSet mInfo = Info
    mClickThru = mInfo.ClassObj.IsClickThru
    mStylePath = StylePath

Dim pbm As mfxBitmap
Dim rc As RECT
Dim i As Long

    m_Alpha = 255 * (g_SafeLong(g_ConfigGet("global_opacity")) / 100)
    mOriginalAlpha = m_Alpha
    mOriginalTimeout = Info.Timeout
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
            g_Debug "CSnarlWindow.Create(): timeout cannot be zero if click-through", LEMON_LEVEL_WARNING
            mOriginalTimeout = IIf(Val(g_ConfigGet("default_duration")) > 0, Val(g_ConfigGet("default_duration")), 10)

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
            mCurrent = 255

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
            
'Debug.Print "TIMEOUT " & mOriginalTimeout
            
            mVisible = True

            Exit Sub

        End If

    End If

'Debug.Print "SHOW: " & m_Alpha

    If Not (theWindow Is Nothing) Then
        uSetTargetAlpha m_Alpha

        Me.SetTimeout mOriginalTimeout

Debug.Print "TIMEOUT " & mOriginalTimeout

'        If m_Timeout > 0 Then _
            theWindow.AddTimer TIMER_MAIN_TICK, m_Timeout * 1000

        mVisible = True

    End If

End Sub

Public Sub Quit()

    If mQuitting Then _
        Exit Sub

    Debug.Print "CSnarlWindow.Quit(): closing..."

    mVisible = False
    If Not (theWindow Is Nothing) Then
        mQuitting = True                                    '// set this notification as no longer valid...
        KillTimer theWindow.hWnd, TIMER_TRACK_POINTER
'        KillTimer theWindow.hWnd, TIMER_PULSE
        theWindow.AddExStyles WS_EX_TRANSPARENT             '// window becomes click-through while fading

        If Not (g_NotificationRoster Is Nothing) Then _
            g_NotificationRoster.Remove mInfo.Token         '// remove from roster so space can be reused

        uSetTargetAlpha 0                                   '// async fade to zero

    End If

End Sub

Public Sub Remove()

    Debug.Print "CSnarlWindow.Remove(): removing window..."

    mAlphaStep = 0

    If Not (theWindow Is Nothing) Then _
        KillTimer theWindow.hWnd, TIMER_ASYNC_FADE

    nto_DestroyWindow theWindow
    Set theView = Nothing

        ' /* R2.4: if this was a low priority notification, check the queue */

'                        If mInfo.Priority < 0 Then _
            g_NotificationRoster.QueueNext

End Sub

Private Sub Class_Terminate()

    Debug.Print "CSnarlWindow.Terminate"

End Sub

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static bPaused As Boolean
Static pt As POINTAPI
Static sz As String
Static f As Boolean
Static t As Long

    Select Case uMsg
    Case WM_SETCURSOR
        SetCursor LoadCursor(0, IDC_ARROW)

'    Debug.Print gNotificationMenuOpen

        If (LoWord(lParam) = HTERROR) Then
            ' /* window is disabled which means a menu is active */
            Select Case HiWord(lParam)
            Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
                Debug.Print "-- enabling --"
                ' /* enable the window but don't process the message */
'                gNotificationMenuOpen = False
                EnableWindow hWnd, -1
                ReturnValue = -1
                MWndProcSink_WndProc = True

            End Select

        ElseIf Not mCloseButtonVisible Then
            mCloseButtonVisible = True
            uDrawContent mCurrent

            If Not mClickThru Then _
                SetWindowPos hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE

            SetTimer hWnd, TIMER_TRACK_POINTER, 100, 0

        End If


    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
        ' /* V38.131: better way of working - button pressed now does a SetCapture() and we only respond
        '    if the mouse button is released over the same item */
        sz = uHitTest(lParam)
        Debug.Print "--> " & sz
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
                        If (g_IsPressed(VK_LSHIFT)) Or (g_IsPressed(VK_RSHIFT)) Then
                            ' /* SHIFT held down: close all */
                            g_NotificationRoster.CloseMultiple 0

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
                        uNotify SNARL_CALLBACK_CLOSED
                        bPaused = False
                        Hide

                    End If
    
                Case "client"
                    g_Debug "CSnarlWindow.WndProc(): client area invoked"
                    bPaused = True                  ' // don't count down while we're doing this...

                    If mInfo.DefaultAck <> "" Then
                        ' /* V38.94 - if an ACK string has been specified, launch it now */
                        g_ProcessAck mInfo.DefaultAck

                    Else
                        ' /* R2.4 DR7: only do the callback if no static callback defined */
                        uNotify SNARL_CALLBACK_INVOKED

                    End If

                    bPaused = False
                    Hide

                Case "action"
                    bPaused = True
                    uDoActionsMenu
                    bPaused = False

                End Select

            Else
                g_Debug "CSnarlWindow.WndProc(): WM_LBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_LBUTTONUP ignored: outside window"

        End If

        sz = ""


'    Case WM_LBUTTONDBLCLK
'        ' /* R2.4 beta 4: make sticky if not already */
'        If mOriginalTimeout > 0 Then _
'            Me.SetTimeout 0


    Case WM_RBUTTONUP
        If f Then
            ReleaseCapture
            If sz = uHitTest(lParam) Then
                Select Case sz
                Case "close"
                    g_Debug "CSnarlWindow.WndProc(): close button invoked (right-click)"
                    bPaused = True
                    uDoSysMenu
                    bPaused = False

                Case "client"
                    bPaused = True
                    If Not (mMenu Is Nothing) Then
                        ' /* track menu */
                        If Not uTrackMenu(hWnd) Then
                            bPaused = False
                            Exit Function

                        End If

                    Else
                        If (g_ConfigGet("allow_right_clicks") = "1") Or ((mInfo.IntFlags And NF_API_MASK) < 42) Then
                            ' /* this is deprecated as of R2.4 DR7 */
                            uNotify SNARL_CALLBACK_R_CLICK

                        Else
                            ' /* R2.4.1: pop open the system menu instead */
                            uDoSysMenu

                        End If

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

                    If (g_ConfigGet("allow_right_clicks") = "1") Or ((mInfo.IntFlags And NF_API_MASK) < 42) Then
                        ' /* this is deprecated as of R2.4 DR7 */

                        bPaused = True
                        uNotify SNARL_CALLBACK_M_CLICK
                        bPaused = False

                    End If

'                    ' /* R2.4: hide when middle button clicked? */
'                    If (mInfo.Flags And SNARL41_NOTIFICATION_AUTO_DISMISS) Then _
'                        Hide

                End Select
            Else
                g_Debug "CSnarlWindow.WndProc(): WM_MBUTTONUP ignored: hit-test mismatch"

            End If
        Else
            g_Debug "CSnarlWindow.WndProc(): WM_MBUTTONUP ignored: outside window"

        End If
        sz = ""


    Case WM_TIMER
        If (wParam = TIMER_MAIN_TICK) And (Not bPaused) And (Not mCloseButtonVisible) And (Not mFuzzy) Then
            mTick = mTick - 1
            If mTick = 0 Then
                bPaused = True
                uNotify SNARL_CALLBACK_TIMED_OUT
                Debug.Print "timed out"
                bPaused = False
                Hide

            End If

        ElseIf wParam = TIMER_ASYNC_FADE Then
            If mAlphaStep < 0 Then
                ' /* fading down */
                If mCurrent <= mTargetAlpha Then
                    ' /* done! */
                    mCurrent = mTargetAlpha
                    KillTimer hWnd, wParam
                    mAlphaStep = 0
                    
                    If mQuitting Then _
                        Remove

                    Exit Function

                End If

            Else
                ' /* fading up */
                If mCurrent >= mTargetAlpha Then
                    ' /* done! */
                    Debug.Print "*** Fading took: " & GetTickCount() - mFadeStart & " ms ***"
                    mCurrent = mTargetAlpha
                    KillTimer hWnd, wParam
                    apply_view_to_window theView, hWnd, mCurrent        ' // re-apply in case we've overshot
                    mAlphaStep = 0
                    Exit Function

                End If
            End If

'            Debug.Print GetTickCount() - t
'            t = GetTickCount()

            ' /* keep fading */
            mCurrent = mCurrent + mAlphaStep
            apply_view_to_window theView, hWnd, mCurrent

        ElseIf (wParam = TIMER_TRACK_POINTER) And (Not f) And (Not mMenuOpen) Then
            GetCursorPos pt
            If Not theWindow.Frame.Contains(nto_NewPoint(pt.x, pt.y)) Then
                KillTimer hWnd, wParam
                mCloseButtonVisible = False
                uDrawContent mCurrent

            End If

        ElseIf wParam = TIMER_PULSE Then
            uPulse

        End If


'    Case WM_NCHITTEST
'
'        Debug.Print "NCHITTEST"
'
'        mCloseButtonVisible = True
'        uDrawContent mCurrent
'        SetTimer hWnd, TIMER_TRACK_POINTER, 100, 0


    Case WM_CAPTURECHANGED
        f = False

    Case WM_NCACTIVATE
        If wParam = 0 Then _
            EnableWindow hWnd, -1

    Case WM_ENTERMENULOOP, WM_EXITMENULOOP
        mMenuOpen = (uMsg = WM_ENTERMENULOOP)

    End Select

End Function

Private Function uHitTest(ByVal lParam As Long) As String
Dim pp As BPoint
Dim pr As BRect

    ' /* returns:
    '       "close" if the pointer is over the close button
    '       "action" if the pointer is over the action button (and actions are defined)
    '       "client" if the pointer is over the notification
    '       "" if the pointer is outside the notification
    ' */

    Set pp = new_BPointFromInt32(lParam)
    Set pr = BW_Bounds(theWindow.hWnd)

    If Not pr.Contains(pp) Then
        g_Debug "CSnarlWindow.uHitTest(): not inside window"
        Exit Function

    End If
'mActions = 1

    If mCloseRect.Contains(pp) Then
        uHitTest = "close"

    ElseIf (mInfo.Actions.CountItems > 0) And (mActionsRect.Contains(pp)) Then
        uHitTest = "action"

    Else
        uHitTest = "client"

    End If

'    Debug.Print uHitTest

'Dim n As Integer
'
'    If uDrawDropshadow() Then _
'        n = Val(g_ConfigGet("dropshadow_size"))
'
'    If (pt.x >= rc.Left) And (pt.x <= rc.Right) And (pt.y >= rc.Top) And (pt.y <= rc.Bottom) Then
'
'        If (pt.x >= n) And (pt.y >= n) And (pt.x <= (n + bm_Close.Width)) And (pt.y <= (n + bm_Close.Height)) Then
'            uHitTest = "close"
'
'        Else
'            uHitTest = "client"
'
'        End If
'
'        Debug.Print uHitTest
'
'    Else
'        g_Debug "CSnarlWindow(): not inside window"
'
'    End If

End Function

Public Sub Hide()
Dim bDone As Boolean
Dim pt As POINTAPI

    If Not (mInstance Is Nothing) Then
        If (mStyleFlags And S_STYLE_CUSTOM_HIDE) Then
            mVisible = False
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

            Remove

            Exit Sub

        Else
            Debug.Print "CSnarlWindow.Hide(): no custom hide required"

        End If

    Else
        Debug.Print "CSnarlWindow.Hide(): no instance"

    End If

    theWindow.RemoveTimer TIMER_MAIN_TICK
    Quit

End Sub

Public Property Get Id() As Long

    Id = mInfo.Token

End Property

'Friend Function Update(ByRef Info As T_NOTIFICATION_INFO) As M_RESULT
''
'''    MsgBox Info.StyleName & " > " & Info.SchemeName, , "Update"
''
''    ' /* pre-set these if they're missing */
''
'''    If Info.StyleName = "" Then _
'''        Info.StyleName = mInfo.StyleName
'''
'''    If Info.SchemeName = "" Then _
'''        Info.SchemeName = mInfo.SchemeName
''
''    ' /* tell the style handler to update its content */
''
''    uUpdateStyleContent Info
''
''    ' /* view is already set - don't bother resizing */
''
''    Set theBitmap = uGetContent2()
''    uDrawContent mCurrent
''
''
''
''    SetTimeout mOriginalTimeout                 ' // V37.80 - reset timeout value
''    Update = M_OK
'
'    Update2 Info
'
'    Update = M_OK
'
'End Function

Public Sub uSetTargetAlpha(ByVal Alpha As Long)

    Debug.Print "CSnarlWindow.uSetTargetAlpha(): current=" & mCurrent & " requested=" & Alpha

    If (Alpha = 0) And (mCurrent = 0) And (mQuitting) Then
        ' /* R2.4 DR8: handles an unusual case whereby the notification is being
        '    removed before async fading has even begun.  */
        Remove
        Exit Sub

    ElseIf Alpha = mCurrent Then
        ' /* nothing to do */
        Exit Sub

    End If

    ' /* figure out granularity */

Dim s As Single

    s = Abs(mCurrent - Alpha)

'    Debug.Print "alpha delta = " & s

    s = (s / 300) * (10 + 6)

    If Alpha < mCurrent Then _
        s = -s

'    Debug.Print "step size = " & s

    mAlphaStep = CLng(s)

    ' /* kick off the fade process */

    mTargetAlpha = Alpha
    SetTimer theWindow.hWnd, TIMER_ASYNC_FADE, 10, 0
    mFadeStart = GetTickCount()

End Sub

Public Sub Enable()

    m_Disable = m_Disable - 1

End Sub

Public Sub Disable()

    m_Disable = m_Disable + 1

End Sub

Public Function SetTimeout(ByVal Timeout As Long) As M_RESULT

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

    IsVisible = mVisible

End Property

Public Function GetView() As mfxView

    Set GetView = theView

End Function

Private Sub uDrawContent(Optional ByVal Alpha As Long = -1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1)

    If (theView Is Nothing) Or (theBitmap Is Nothing) Or (theWindow Is Nothing) Then _
        Exit Sub

    If Alpha = -1 Then _
        Alpha = mCurrent

    If uDrawDropshadow() Then
        If (x = -1) And (y = -1) Then
'            Set pr = theWindow.Frame
'            x = pr.Left - 10
'            y = pr.Top - 10

        Else
            ' /* specific position requested and dropshadow enabled */
            x = x - Val(g_ConfigGet("dropshadow_size"))
            y = y - Val(g_ConfigGet("dropshadow_size"))

        End If
    End If

Dim lSize As Long
Dim pr As BRect
Dim cy As Long

    With theView
        .Clear

        lSize = MIN(bm_Close.Height, .Height)

        .DrawScaledImage theBitmap, new_BPoint(OFFSET_ICON, OFFSET_ICON)
        If (mCloseButtonVisible) And (Not mClickThru) Then

            ' /* draw close button */
            .DrawScaledImage bm_Close, mCloseRect.TopLeft, new_BPoint(mCloseRect.Width, mCloseRect.Height)

            ' /* R2.4: if the notification has actions, draw the actions button */
            If mInfo.Actions.CountItems > 0 Then _
                .DrawScaledImage bm_Actions, mActionsRect.TopLeft, new_BPoint(mActionsRect.Width, mActionsRect.Height)

        End If

        If (g_ConfigGet("show_timestamp") = "1") And (.Width > 79) Then
            If ((mCloseButtonVisible) And (Not mClickThru)) Or (mFuzzy) Then
                .SetFont "Tahoma", 8, True
                .TextMode = MFX_TEXT_ANTIALIAS
                .SetHighColour rgba(255, 255, 255, 230)
                .SetLowColour rgba(0, 0, 0, 190)
'                Set pr = .Bounds.InsetByCopy(2, 2)
'                If uDrawDropshadow() Then _
'                    pr.InsetBy Val(g_ConfigGet("dropshadow_size")), Val(g_ConfigGet("dropshadow_size"))

                ' /* timestamp overlay */
                Set pr = new_BRect(0, 0, .StringWidth(g_When(mInfo.DateStamp)) + 3, .StringHeight("A"))
                pr.OffsetBy .Width - pr.Width, .Height - pr.Height
                If uDrawDropshadow() Then _
                    pr.OffsetBy -Val(g_ConfigGet("dropshadow_size")), -Val(g_ConfigGet("dropshadow_size"))

                cy = pr.Height
                .FillRoundRect pr, 6, 6, MFX_SOLID_LOW
                .DrawString g_When(mInfo.DateStamp), pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

                ' /* R2.4b4: app name overlay */
                Set pr = new_BRect(0, 0, .StringWidth(mInfo.ClassObj.App.Name) + 3, .StringHeight("A"))
                pr.OffsetBy .Width - pr.Width, .Height - pr.Height - cy
                If uDrawDropshadow() Then _
                    pr.OffsetBy -Val(g_ConfigGet("dropshadow_size")), -Val(g_ConfigGet("dropshadow_size"))

                .FillRoundRect pr, 6, 6, MFX_SOLID_LOW
                .DrawString mInfo.ClassObj.App.Name, pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

'                ' /* timestamp overlay */
'                .FillRoundRect new_BRect(pr.Right - , pr.Bottom - , pr.Right + 2, pr.Bottom + 2), 6, 6, MFX_SOLID_LOW
'                .DrawString uTimeStamp(), pr, MFX_ALIGN_BOTTOM Or MFX_ALIGN_RIGHT
'
'                ' /* R2.4b4: app name overlay */
'                .FillRoundRect new_BRect(pr.Left - 2, pr.Bottom - .StringHeight("A"), pr.Left + .StringWidth(mInfo.ClassObj.App.Name), pr.Bottom + 2), 6, 6, MFX_SOLID_LOW
'                .DrawString mInfo.ClassObj.App.Name, pr, MFX_ALIGN_BOTTOM Or MFX_ALIGN_LEFT

            End If
        End If

Const rx = 6

Dim n As Long

        ' /* can only ever have one low priority notification visible at any one time
        '    (obviously this excludes visible non-windowed styles) */

'        If mInfo.Priority < 0 Then _
            n = g_NotificationRoster.QueueCount

        ' /* if we have queued notifications, draw the number in a badge */

        If n > 0 Then
            .SetFont "Tahoma", 8, True
            .TextMode = MFX_TEXT_ANTIALIAS
            .EnableSmoothing True

            ' /* figure out the badge size and position */

            Set pr = new_BRect(0, 0, MAX(.StringWidth(CStr(n)), .StringHeight("A")), .StringHeight("A"))
            pr.ExpandBy 8, 8
            pr.OffsetBy .Width - pr.Width - 1, 1

            If uDrawDropshadow() Then _
                pr.OffsetBy -Val(g_ConfigGet("dropshadow_size")), Val(g_ConfigGet("dropshadow_size"))

            ' /* background */

            .SetHighColour rgba(0, 0, 64, 190)
            .SetLowColour rgba(0, 0, 64, 100)
            .FillRoundRect pr, rx, rx, MFX_VERT_GRADIENT

            ' /* label */

            .SetHighColour rgba(255, 255, 255)
            .DrawString CStr(n), pr, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

            ' /* border */

            .SetHighColour rgba(255, 255, 255)
            .StrokeRoundRect pr.InsetByCopy(1, 1), rx, rx, 2

            ' /* edge and inset */

            .SetHighColour rgba(0, 0, 0, 150)
            .StrokeRoundRect pr, rx, rx, 1
            .StrokeRoundRect pr.InsetByCopy(3, 3), rx, rx, 1

        End If

    End With

    apply_view_to_window theView, theWindow.hWnd, Alpha, x, y

End Sub

Friend Sub bSetAlpha(ByVal Alpha As Integer)

    If Alpha < 0 Then
        Alpha = 0

    ElseIf Alpha > 255 Then
        Alpha = 255

    End If

    mCurrent = Alpha
    uDrawContent mCurrent

End Sub

Private Sub uDoSysMenu()
Dim pmi As OMMenuItem
Dim pm As OMMenu

    With New OMMenu
        ' /* R2.2 */
        .AddItem .CreateItem("stky", "Make Sticky", , (mOriginalTimeout > 0))
        .AddItem .CreateItem(" ", "Details", , , , , , uDetailsMenu())          ' // R2.3
        .AddSeparator

        If mInfo.ClassObj.App.HasConfig Then _
            .AddItem .CreateItem("acfg", mInfo.ClassObj.App.Name & " Settings...", , mInfo.ClassObj.App.HasConfig)

        .AddItem .CreateItem("$" & mInfo.ClassObj.App.Name & "#?" & mInfo.ClassObj.Name, "Notification Settings...")

        .AddSeparator

        .AddItem .CreateItem("wlog", "Write content to Log")
        .AddItem .CreateItem("copy", "Copy content to Clipboard")

        Set pm = g_GetRemoteComputersMenu()
        If Not (pm Is Nothing) Then _
            .AddItem .CreateItem(" ", "Send Content to", , Not (pm Is Nothing), , , , pm)

        .AddSeparator
        .AddItem .CreateItem("clse", "Close")
        .AddSeparator
        .AddItem .CreateItem("capp", "Close all from " & mInfo.ClassObj.App.Name)
        .AddItem .CreateItem("call", "Close all")

        If g_IsPressed(VK_CONTROL) Then
            .AddSeparator
            .AddItem .CreateItem("scrn", "Capture Notification")

        End If

        EnableWindow theWindow.hWnd, 0

        Set pmi = .Track(theWindow.hWnd)

    End With

Dim sz() As String

    If Not (pmi Is Nothing) Then
        EnableWindow theWindow.hWnd, -1         ' // item was selected so re-enable the window
        
        Select Case pmi.Name
        Case "clse"
            Hide

        Case "capp"
            g_NotificationRoster.CloseMultiple mInfo.ClassObj.App.Token

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

        Case "acfg"
            mInfo.ClassObj.App.DoSettings

        Case Else
            
            Select Case g_SafeLeftStr(pmi.Name, 1)
            Case "$"
                ' /* app settings */
                sz = Split(g_SafeRightStr(pmi.Name, Len(pmi.Name) - 1), "#?")
                frmAbout.DoAppConfig sz(0), sz(1)

            End Select
            
'            If g_SafeLeftStr(pmi.Name, 3) = "ip>" Then
'                MsgBox g_SafeRightStr(pmi.Name, Len(pmi.Name) - 3)

'                With New SNBConnection
'                    If .Connect(g_SafeRightStr(pmi.Name, Len(pmi.Name) - 3)) = SNB_SUCCESS Then
'                        .ShowNotification "", mInfo.Title, mInfo.Text & vbCrLf & vbCrLf & "Sent from " & get_host_name(), mOriginalTimeout
'                        .Disconnect
'
'                    End If
'
'                End With
'
'            End If

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
    uDrawContent mCurrent

    CaptureToFile sz

    mCloseButtonVisible = True
    uDrawContent mCurrent

End Sub

Public Function CaptureToFile(ByVal Filename As String)

    If Not (theView Is Nothing) Then _
        theView.WriteToFile Filename, "image/png"

End Function

Public Function SetMenu(ByVal Value As String) As M_RESULT

    ' /* 'Value' should be a string formatted as <Label>[#?<value>][|<Label>[#?<value>]]...
    '    <Label> defines the menu item label and can be any text string.  <Value> must be numeric
    '    and in the range of 1 to 65535.  <Value> does not need to be specified, in which
    '    case the item is added but no notification is sent if it is selected.  If <Label> is
    '    blank, a separator item is added (which cannot be selected) */

    ' /* R2.3 onwards: if <Label> is prefixed with '$' the following character is used as
    '    a control character.  The following are currently defined:
    '       'X' a checkmark is displayed against the menu item
    '       '\' the menu item is ghosted and cannot be selected
    '       'x'  a checkmark is displayed against the menu item and the menu item is ghosted and cannot be selected
    ' */

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
Dim i As Long

    i = InStr(Item, "#?")
    If i = 0 Then
        szText = Item

    Else
        szText = g_SafeLeftStr(Item, i - 1)
        szData = g_SafeRightStr(Item, Len(Item) - i - 1)

    End If

Dim itemEnabled As Boolean
Dim itemMarked As Boolean

    ' /* pre-set this */

    itemEnabled = True

    If szText = "" Then
        ' /* R2.3 seperator added even if <value> supplied with empty <Label> */
        Menu.AddSeparator
        Exit Sub

    ElseIf g_SafeLeftStr(szText, 1) = "$" Then
        ' /* control character */

        Select Case g_SafeMidStr(szText, 2, 1)
        Case "X"
            ' /* checkmark */
            itemMarked = True

        Case "\"
            ' /* disabled */
            itemEnabled = False

        Case "x"
            ' /* disabled and marked */
            itemMarked = True
            itemEnabled = False

        Case Else
            g_Debug "CSnarlWindow.uAddToMenu(): invalid control character '" & g_SafeMidStr(szText, 2, 1) & "'", LEMON_LEVEL_CRITICAL
            Exit Sub

        End Select

        szText = g_SafeRightStr(szText, Len(szText) - 2)

    End If

Dim dw As Long

    ' /* if <Value> is provided, make sure it's valid */

'    If szData <> "" Then
'        If Not g_IsNumeric(szData) Then
'            g_Debug "CSnarlWindow.uAddToMenu(): can't add '" & Item & "' to menu: data part is not numeric", LEMON_LEVEL_CRITICAL
'            Exit Sub
'
'        End If
'
'        dw = Val(szData)
'        If (dw < 1) Or (dw > 65535) Then
'            g_Debug "CSnarlWindow.uAddToMenu(): can't add '" & Item & "' to menu: data part must be 1-65535", LEMON_LEVEL_CRITICAL
'            Exit Sub
'
'        End If
'
'    End If

    Menu.AddItem mMenu.CreateItem(szData, szText, , itemEnabled, itemMarked)

End Sub

Private Function uTrackMenu(ByVal hWndOwner As Long) As Boolean

    If (mMenu Is Nothing) Then _
        Exit Function

Dim pmi As OMMenuItem

    EnableWindow hWndOwner, 0
    Set pmi = mMenu.Track(hWndOwner)
    If (pmi Is Nothing) Then _
        Exit Function

    EnableWindow hWndOwner, -1          ' // item was selected so re-enable the window

    If g_SafeLeftStr(pmi.Name, 1) = "!" Then
        ' /* R2.4: menu id's can be bang commands, paths, urls, etc. */
        g_ProcessAck pmi.Name

    Else
        ' /* send it as a SNARL_NOTIFICATION_MENU notification */
        uNotify SNARL_CALLBACK_MENU_SELECTED, pmi.Name

    End If

    uTrackMenu = True

End Function

Private Sub uNotify(ByVal Code As Integer, Optional ByVal Data As String)
Dim szResponse As String
Dim szData As String
Dim n As Integer

    If Not (mInfo.Socket Is Nothing) Then
        ' /* if we have a socket, then reply to that */

        Code = Code + 270               ' // convert into "proper" V42 notification

        If (mInfo.IntFlags And NF_IS_GNTP) Then
            ' /* R2.4.2: GNTP handling */

            Select Case Code
            Case SNARL_NOTIFY_INVOKED
                szData = "CLICKED"

            Case SNARL_NOTIFY_EXPIRED
                szData = "TIMEDOUT"

            Case SNARL_NOTIFY_CLOSED
                szData = "CLOSED"

            Case Else
                g_Debug "CSnarlWindow.uNotify(): GNTP does not support notification " & CStr(Code)
                Exit Sub

            End Select

            If gntp_CreateCallbackResponse(szData, mInfo.ClassObj.App.Name, mInfo.OriginalContent, szResponse) Then
                g_Debug "CSnarlWindow.uNotify(): sending GNTP callback response to " & mInfo.Socket.RemoteHostIP & ":" & CStr(mInfo.Socket.RemotePort)
                mInfo.Socket.SendData szResponse

            Else
                g_Debug "CSnarlWindow.uNotify(): failed to create GNTP response", LEMON_LEVEL_CRITICAL

            End If

            ' /* close the socket now, as per the GNTP specification */

            mInfo.Socket.CloseSocket

        Else
            ' /* SNP2 */

            ' /* most callbacks are deprecated as of V42 - INVOKED, EXPIRED and ACTION are the only
            '    ones which aren't, so we only do these for SNP */

            Select Case Code
            Case SNARL_NOTIFY_INVOKED, SNARL_NOTIFY_EXPIRED
                ' /* nothing special to do here */

            Case SNARL_NOTIFY_ACTION, SNARL_NOTIFY_MENU
                ' /* send the data verbatim */
                szData = "/" & Data

            Case Else
                g_Debug "CSnarlWindow.uNotify(): not sending deprecated command " & CStr(Code)
                Exit Sub

            End Select

            g_Debug "CSnarlWindow.uNotify(): dest=" & mInfo.Socket.RemoteHostIP & ":" & CStr(mInfo.Socket.RemotePort)

            mInfo.Socket.SendData "SNP/2.0/" & CStr(Code) & _
                                  "/" & uSNP2CallbackResponse(Code) & _
                                  "/" & CStr(mInfo.Token) & szData & _
                                  vbCrLf

        End If

    ElseIf (IsWindow(mInfo.hWndReply) <> 0) And (mInfo.uReplyMsg <> 0) Then
        ' /* otherwise, if we have a valid window and message, reply to that */

        n = LoWord(g_SafeLong(g_SafeRightStr(Data, Len(Data) - 1)))
        g_Debug "CSnarlWindow.uNotify(): sending " & CStr(Code) & " to 0x" & g_HexStr(mInfo.hWndReply) & _
                " using 0x" & g_HexStr(mInfo.uReplyMsg, 4) & " token=" & CStr(mInfo.Token) & " data=" & CStr(n), LEMON_LEVEL_INFO

        PostMessage mInfo.hWndReply, mInfo.uReplyMsg, MAKELONG(Code, n), ByVal mInfo.Token

    Else
        g_Debug "CSnarlWindow.uNotify(): no reply socket or window"

    End If

End Sub

Private Function uSNP2CallbackResponse(ByVal Id As Long) As String
Dim sz As String

    Select Case Id

    Case SNARL_NOTIFY_INVOKED                   '// notification was left-clicked
        sz = "INVOKED"

    Case SNARL_NOTIFY_ACTION                    '// user picked an action from the pop-up
        sz = "ACTION"

    Case SNARL_NOTIFY_EXPIRED
        sz = "EXPIRED"

    Case Else
        sz = "(undefined)"

    End Select

    uSNP2CallbackResponse = sz

End Function

Private Sub uQuickUpdate()

    Set theBitmap = uGetContent2()
    uDrawContent

End Sub

' /*---------------------------------------------------------------------------------------
'   Update() -- intelligent update
'
'   This function updates the existing notification's content, then asks the style instance
'   for an updated position (the style instance can resize the content during the update
'   request) and re-positions it on screen at the requested location and size.
'
' ---------------------------------------------------------------------------------------*/

Friend Function Update(ByRef Info As T_NOTIFICATION_INFO) As Boolean

    ' /* ask our style instance to update */

    uUpdateStyleContent Info

    ' /* get the new content */

Dim pb As mfxBitmap

    Set pb = uGetContent2()
    If (pb Is Nothing) Then _
        Exit Function

    ' /* ask the instance if it wants to change the notification's position */

Dim pt As POINTAPI

    pt.x = theWindow.Frame.Left
    pt.y = theWindow.Frame.Top
    Me.AdjustPosition pt.x, pt.y

    ' /* reset timeout */

    SetTimeout mOriginalTimeout

    ' /* re-position and re-size it */

    Set theBitmap = pb
    theView.SizeTo theBitmap.Width + OFFSET_ICON, theBitmap.Height + OFFSET_ICON
    g_MoveWindow theWindow.hWnd, pt.x, pt.y

    ' /* draw it */

    uDrawContent mCurrent
    Update = True

End Function

Public Function NotificationOnlyMode() As Boolean

    NotificationOnlyMode = mClickThru

End Function

Public Sub MakeFuzzy(ByVal Enabled As Boolean)

    If (mFuzzy = Enabled) Or (Not mVisible) Then _
        Exit Sub

Static bmOld As mfxBitmap

    mFuzzy = Enabled

    If mFuzzy Then
        Set bmOld = theBitmap.Duplicate
        theBitmap.MakeGreyscale
        uDrawContent

        ' /* R2.4: bring window to front - required? */

        SetWindowPos theWindow.hWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
        uSetTargetAlpha MIN(102, Fix(mOriginalAlpha / 2))

    Else

        Set theBitmap = bmOld.Duplicate
        Set bmOld = Nothing
        uDrawContent

        Debug.Print "CSnarlWindow.MakeFuzzy(): restoring to alpha " & mOriginalAlpha
        uSetTargetAlpha mOriginalAlpha

    End If

End Sub

Private Function uGetContent2() As mfxBitmap
Dim pbm As mfxBitmap

    On Error Resume Next

    If Not (mInstance Is Nothing) Then _
        Set pbm = create_bitmap_from_image(mInstance.GetContent())

    If (pbm Is Nothing) Then
        ' /* managed style didn't provide content so we switch to the built-in one */
        g_Debug "CSnarlWindow.uGetContent2(): selected style didn't provide usable content", LEMON_LEVEL_WARNING
        Set mInstance = New TInternalStyle
        mStyleFlags = S_STYLE_MULTI_INSTANCE
        uUpdateInstance
        Set pbm = create_bitmap_from_image(mInstance.GetContent())

    End If

    If (pbm Is Nothing) Then
        g_Debug "CSnarlWindow.uGetContent2(): FATAL: couldn't not get notification content", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* R2.4: apply scaling */

Dim pv As mfxView
Dim s As Single
Dim n As Long

    s = Val(g_ConfigGet("scaling"))
    If s <= 0 Then _
        s = 1               ' // must NOT be zero!

    ' /* button rectangles, scaled as well */

    Set mCloseRect = new_BRect(0, 0, 19 * s, 19 * s)

    ' /* create a temporary view */

    Set pv = New mfxView
    pv.SizeTo pbm.Width * s, pbm.Height * s
    pv.DrawScaledImage pbm, , new_BPoint(pv.Width, pv.Height)

    ' /* emblems */

Dim pt As BPoint
Const EMBLEM_SIZE = 14

    Set pt = new_BPoint(pv.Width - EMBLEM_SIZE - 3, pv.Height - EMBLEM_SIZE - 3)

    ' /* if the notification is high priority */

    If mInfo.Priority > 0 Then
        pv.DrawScaledImage bm_Priority, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), 230
        pt.OffsetBy -(EMBLEM_SIZE + 1), 0

    End If

    ' /* if the notification is sticky */

    If mOriginalTimeout = 0 Then
        pv.DrawScaledImage bm_IsSticky, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), 230
        pt.OffsetBy -(EMBLEM_SIZE + 1), 0

    End If

    ' /* if notification isn't click-through and the mouse pointer isn't
    '    over it, and it has actions or a default callback, draw the actions indicator emblem */
                    '//And (Not mCloseButtonVisible)
    If (Not mClickThru) And ((mInfo.Actions.CountItems > 0) Or (mInfo.DefaultAck <> "")) Then
        pv.DrawScaledImage bm_HasActions, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), 230
        pt.OffsetBy -(EMBLEM_SIZE + 1), 0

    End If

    ' /* if the notification was generated by a remote app */

    If (mInfo.IntFlags And NF_REMOTE) Then
        pv.DrawScaledImage bm_Remote, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), 230
        pt.OffsetBy -(EMBLEM_SIZE + 1), 0

    End If

    ' /* if the notification was generated by a secure app */

    If (mInfo.IntFlags And NF_SECURE) Then
        pv.DrawScaledImage bm_Secure, pt, new_BPoint(EMBLEM_SIZE, EMBLEM_SIZE), 230
        pt.OffsetBy -(EMBLEM_SIZE + 1), 0

    End If


    ' /* draw with or without dropshadow */

    If uDrawDropshadow() Then
        ' /* has a drop shadow */
        n = Val(g_ConfigGet("dropshadow_size"))
        mCloseRect.OffsetBy n, n
        Set uGetContent2 = create_dropshadow(pv, , n * 2, Fix((Val(g_ConfigGet("dropshadow_strength")) / 100) * 255))

    Else
        ' /* no drop shadow */
        Set uGetContent2 = pv.ConvertToBitmap()

    End If



    ' /* set up the correct locations for the close and actions buttons */

    If pbm.Height > 47 Then
        Set mActionsRect = mCloseRect.OffsetByCopy(0, mCloseRect.Height)

    Else
        Set mActionsRect = mCloseRect.OffsetByCopy(mCloseRect.Width, 0)

    End If

End Function

Private Sub uUpdateStyleContent(ByRef Info As T_NOTIFICATION_INFO)

    On Error GoTo er

    If (mInstance Is Nothing) Then _
        Set mInstance = New TInternalStyle


    With Info

        ' /* translate special markers */

'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)

        ' /* take into consideration and "do not change" markers */

        If .Title = Chr$(255) Then _
            .Title = mInfo.Title

        If .Text = Chr$(255) Then _
            .Text = mInfo.Text

        If .IconPath = Chr$(255) Then _
            .IconPath = mInfo.IconPath

        ' /* this actually only contains the scheme name */

        If .SchemeName = Chr$(255) Then _
            .SchemeName = mInfo.SchemeName

    End With

    ' /* store the new values */

    With mInfo
        .Title = Info.Title
        .Text = Info.Text
        .IconPath = Info.IconPath
        .SchemeName = Info.SchemeName
        .LastUpdated = Now()            ' // update this as well

    End With

    uUpdateInstance
    Exit Sub

er:
    g_Debug "CSnarlWindow.uUpdateStyleContent(): " & err.Description, LEMON_LEVEL_CRITICAL

End Sub

Private Sub uUpdateInstance()

    If (mInstance Is Nothing) Then _
        Exit Sub

    ' /* translate into a notification_info struct that the style type lib can understand */

Dim pInfo As notification_info

    With pInfo
'        .Title = Replace$(Info.Title, Chr$(255), "")
'        .Text = Replace$(Info.Text, Chr$(255), "")
'        .Icon = Replace$(Info.IconPath, Chr$(255), "")
'        .Scheme = Replace$(LCase$(Info.SchemeName), Chr$(255), "")

        If (mStyleFlags And S_STYLE_V42_CONTENT) Then
            .Title = mInfo.ClassObj.App.Name
            .Text = mInfo.OriginalContent

        Else
            .Title = mInfo.Title
            .Text = mInfo.Text

        End If

        .Icon = g_TranslateIconPath(mInfo.IconPath, mStylePath)
        .Scheme = LCase$(mInfo.SchemeName)

        If mInfo.Priority > 0 Then _
            .Flags = .Flags Or S_NOTIFICATION_IS_PRIORITY

    End With

    mInstance.UpdateContent pInfo

End Sub

Public Sub AdjustPosition(ByRef x As Long, ByRef y As Long)

    If uDrawDropshadow() Then
        x = x + Val(g_ConfigGet("dropshadow_size"))
        y = y + Val(g_ConfigGet("dropshadow_size"))

    End If

    mInstance.AdjustPosition x, y, 255, True

    ' /* compensate for dropshadow */

    If uDrawDropshadow() Then
        x = x - Val(g_ConfigGet("dropshadow_size"))
        y = y - Val(g_ConfigGet("dropshadow_size"))

    End If

End Sub

Private Function uDrawDropshadow() As Boolean

    ' /* styles can veto the system drop shadow */

    uDrawDropshadow = (gPrefs.use_dropshadow) And ((mStyleFlags And S_STYLE_NO_DROPSHADOW) = 0)

End Function

Private Sub uPulse()

    If Not (mInstance Is Nothing) Then
        If mInstance.Pulse() Then _
            uQuickUpdate

    End If

End Sub

Public Function StylePath() As String

    StylePath = mStylePath

End Function

Public Function Frame() As BRect

    If Not (theWindow Is Nothing) Then _
        Set Frame = new_BRectFromRect(theWindow.Frame)

End Function

Private Function uDetailsMenu() As OMMenu

    Set uDetailsMenu = New OMMenu
    With uDetailsMenu
        .AddItem .CreateItem("", "Received: " & Format$(mInfo.DateStamp, "d mmm yyyy") & " at " & Format$(mInfo.DateStamp, "ttttt"))
        .AddItem .CreateItem("$" & mInfo.ClassObj.App.Name & "#?" & mInfo.ClassObj.Name, "Generated by: " & mInfo.ClassObj.App.Name & " (" & mInfo.ClassObj.Description & " class)")

    End With

End Function

Public Function Title() As String

    Title = mInfo.Title

End Function

Public Function Text() As String

    Text = mInfo.Text

End Function

Public Sub QueueCountChanged()

    uDrawContent

End Sub

Public Function IsMergeCandidate(ByRef ComparisonClass As TAlert, ByVal ComparisonTitle As String) As Boolean

    If (ComparisonClass Is Nothing) Or ((mInfo.IntFlags And NF_MERGE) = 0) Then _
        Exit Function

    If (ComparisonClass.App.Signature = mInfo.ClassObj.App.Signature) And (ComparisonClass.Name = mInfo.ClassObj.Name) Then _
        IsMergeCandidate = (ComparisonTitle = mInfo.Title) Or (ComparisonTitle = "")

End Function

'Friend Sub MergeWith(ByRef Info As T_NOTIFICATION_INFO)
''Dim t As Long
'
''    t = (mTick \ 10) + Info.Timeout
'    Info.Text = mInfo.Text & vbCrLf & Info.Text
'    Info.Text = g_TrimLines(Info.Text)              ' // still apply overflow trimming...
'
'    g_KludgeNotificationInfo Info
'
'    Update Info                                     ' // update icon and title (if applicable)
'    SetTimeout mInfo.Timeout                        ' // set new timeout
'
'End Sub

Public Sub RethinkActions()

    If (mInfo.Actions.CountItems = 1) Or (mInfo.Actions.CountItems = 0) Then _
        uQuickUpdate

End Sub

Public Function HasActions() As Boolean

    HasActions = (mInfo.Actions.CountItems <> 0)

End Function

Private Sub uDoActionsMenu()

    If mInfo.Actions.CountItems = 0 Then
        g_Debug "CSnarlWindow.uDoActionsMenu(): no actions defined", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

'    If (IsWindow(mInfo.hWndReply) = 0) Or (mInfo.uReplyMsg = 0) Then
'        g_Debug "CSnarlWindow.uDoActionsMenu(): bad reply-to window or reply message", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If

    EnableWindow theWindow.hWnd, 0

Dim pmi As OMMenuItem
Dim i As Long

    With New OMMenu
        For i = 1 To mInfo.Actions.CountItems
            .AddItem .CreateItem(CStr(i), mInfo.Actions.TagAt(i).Name)
            If i < mInfo.Actions.CountItems Then _
                .AddSeparator

        Next i

        Set pmi = .Track(theWindow.hWnd, BW_Frame(theWindow.hWnd).TopLeft.OffsetByCopy(mActionsRect.Left, mActionsRect.Bottom))

    End With

    If (pmi Is Nothing) Then _
        Exit Sub

    ' /* get selected command */

Dim szCmd As String

    i = Val(pmi.Name)
    szCmd = mInfo.Actions.TagAt(i).Value

    If g_SafeLeftStr(szCmd, 1) = "@" Then
        uNotify SNARL_NOTIFY_ACTION, szCmd          ' /* dynamic callback */

    Else
        g_ProcessAck szCmd                          ' /* static callback - process but don't do a callback */

    End If
        
    EnableWindow theWindow.hWnd, -1           ' // item was selected so re-enable the window

'    If (mInfo.Flags And SNARL41_NOTIFICATION_AUTO_DISMISS) Then _
        Me.Hide

End Sub

'Public Sub ClearActions()
'
'    mActions = 0
'    ReDim mAction(0)
'
'End Sub

'Friend Function Update(ByRef Info As T_NOTIFICATION_INFO) As Boolean
''Dim pInfo As T_NOTIFICATION_INFO
''Dim sz As String
'
'    ' /* get our current content */
'
''    LSet pInfo = mInfo
'
'    ' /* update content with only what's provided */
'
''    With mInfo
''
''        If Args.Exists("title") Then
''            sz = Replace$(Args.ValueOf("title"), "\n", vbCrLf)
''            sz = Replace$(sz, "\+", .Title)                         ' // specfic to Update() - replace '\+' with existing title
''            .Title = g_TrimLines(sz)
''
''        End If
''
''        If Args.Exists("text") Then
''            sz = Replace$(Args.ValueOf("text"), "\n", vbCrLf)
''            sz = Replace$(sz, "\+", .Text)                          ' // specfic to Update() - replace '\+' with existing text
''            .Text = g_TrimLines(sz)
''
''        End If
''
''        If Args.Exists("icon") Then _
''            .IconPath = Args.ValueOf("icon")
''
''        If Args.Exists("timeout") Then _
''            .Timeout = g_SafeLong(Args.ValueOf("timeout"))          ' // g_SafeLong() returns 0 on error - not good...
''
''    End With
''
''    ' /* create the OriginalContent value */
''
''    uFixupContent
''    g_KludgeNotificationInfo pInfo
''    LSet mInfo = pInfo
'
'    ' /* tell the style to apply the changes and redraw */
'
'    Update = uUpdate(Info)
'
'End Function

