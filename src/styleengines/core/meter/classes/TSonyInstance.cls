VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSonyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const EM_TITLE_TEXT = 0
Const EM_MARGIN = 10

Public Style As TStyle

Dim mInfo As notification_info
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mIsNumeric As Boolean
Dim mIconOnly As Boolean

Const SHADOW_SIZE = 10

Dim mTitleRect As BRect
Dim mTextRect As BRect
Dim mSettings As T_METER_STYLE_SETTINGS

Implements IStyleInstance

Private Sub Class_Initialize()

    Set mView = New mfxView

    With mView
        .TextMode = MFX_TEXT_ANTIALIAS

    End With

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, Y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
'Static nStep As Integer
'Const STEP_SIZE = 16
'
'    If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'
'        If nStep = 0 Then _
'            nStep = STEP_SIZE
'
'        mRed = mRed + nStep
'        If mRed > 254 Then
'            nStep = -STEP_SIZE
'
'        ElseIf mRed < 1 Then
'            nStep = STEP_SIZE
'
'        End If
'
'        uDraw
'
'        IStyleInstance_Pulse = True
'
'    End If

End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    LSet mInfo = NotificationInfo
    LSet mSettings = gSettings.Sony

    ' /* clear cached icon */

    Set mIcon = Nothing
    mIconSize = 0

Dim pi As MImage

    ' /* cache icon */

    Set pi = load_image_obj(NotificationInfo.Icon)
    If is_valid_image(pi) Then
        Set mIcon = style_MakeSquareImage(pi)
        mIconSize = MIN(mIcon.Height, 48)

    End If

    ' /* pre-configuration */

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))
    mIsNumeric = (g_IsNumeric(mInfo.Text) And (Val(Style.ReadSetting("show_graphically"))))

    ' /* calculate notifcation size and internal measurements */

    uCalcSize

    uDraw

End Sub

Private Sub uDraw()
Dim m As Long

    ' /* draw the background */

    With mView
        .Clear
        .EnableSmoothing False

        ' /* background */

        .SetHighColour IIf(InStr(mInfo.Scheme, "psp"), g_GetPSPColour(), g_GetPS3Colour())
        .FillRect .Bounds

'        .SetHighColour rgba(0, 0, 0, 102)
'        .SetLowColour rgba(0, 0, 0, 0)
'        .FillRect .Bounds, MFX_VERT_GRADIENT
'        .SetHighColour rgba(0, 0, 0, 127)
'        .SetLowColour rgba(255, 255, 255, 127)
'        .StrokeFancyRect .Bounds


        ' /* for the PS3 darken background based on time of day */

        If InStr(mInfo.Scheme, "ps3") Then

            If gDbgMinute = -1 Then
                ' /* get time of day in minutes */
                m = (Hour(Now) * 60) + Minute(Now)

            Else
                Debug.Print "using debug setting"
                m = gDbgMinute

            End If

            Debug.Print "minute: " & m

            If m > 1080 Then
                ' /* evening (after 6pm) */
                .SetHighColour rgba(0, 0, 0, 255)
                .SetLowColour rgba(0, 0, 0, m * 0.15)

            ElseIf m < 360 Then
                ' /* early morning (before 6am) */
                .SetHighColour rgba(0, 0, 0, 255)
                .SetLowColour rgba(0, 0, 0, 255 - (m * 0.15))

            Else
                ' /* daytime (0600-1800) */
                m = Abs(720 - m)                        ' // brightness strength is based on distance from midday
                .SetHighColour rgba(0, 0, 0, m * 0.7)
                .SetLowColour rgba(0, 0, 0, 0)

            End If

            .FillRect .Bounds, MFX_VERT_GRADIENT

            .EnableSmoothing True
            .SetHighColour rgba(255, 255, 255, 63)
            .FillEllipse new_BRect(-300, -30, .Bounds.Right + 90 - 1, .Bounds.Top + 47)
            .SetHighColour rgba(255, 255, 255, 102)
            .StrokeEllipse new_BRect(-300, -30, .Bounds.Right + 90 - 1, .Bounds.Top + 47)
            .EnableSmoothing False

        End If

        ' /* priority notifications get a red gradient starting at the bottom edge and fading up */

        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 0, 0, 190)
            .SetLowColour rgba(255, 0, 0, 0)
            .SetHighlightPosition .Bounds.Right, .Bounds.Bottom
            .SetHighlightFocusScale 0.1, 0.1
            .FillRect .Bounds, MFX_CENTER_GRADIENT

        End If

        .SetHighColour rgba(0, 0, 0)
        .StrokeRect .Bounds

    End With

    ' /* draw the foreground */

    uDrawForeground

End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then
        g_Debug "TSonyInstance.uCalcSize(): view not created", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim cyNotification As Long

    cyNotification = 80
    Set mTitleRect = new_BRect(0, 0, 0, 0)
    Set mTextRect = new_BRect(0, 0, 0, 0)

    ' /* icon-only notifications are a fixed width and height */

    If mIconOnly Then
        mView.SizeTo 250, cyNotification
        Exit Sub

    End If

Dim prTemp As BRect
Dim pr As BRect

    ' /* measure title and text to see how big the notification needs to be */

    With mView

        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */

        Set pr = new_BRect(0, 0, 250, cyNotification)
        pr.InsetBy EM_MARGIN, EM_MARGIN

        ' /* if we have an icon then shrink accordingly */
        If mIconSize > 0 Then _
            pr.Right = pr.Right - (mIconSize + EM_MARGIN)

        ' /* measure title */
        If mInfo.Title <> "" Then
            .SetFont Style.ReadSetting("title_font"), Style.ReadSetting("title_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Title, prTemp, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width + 2, prTemp.Height + 2)

        End If

        ' /* measure text */
        If (mInfo.Text <> "") And (Not mIsNumeric) Then
            .SetFont Style.ReadSetting("text_font"), Style.ReadSetting("text_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Text, prTemp, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width + 2, prTemp.Height + 2)

        End If

        ' /* size the view */

        .SizeTo 250, MAX(mTitleRect.Height + mTextRect.Height + (2 * EM_MARGIN), cyNotification)

    End With

End Sub

Private Sub uDrawForeground()
Dim pt As BPoint

    ' /* icon */

    If mIconSize > 0 Then
        With mView
            If mIconOnly Then
                ' /* centre of the notification */
                Set pt = new_BPoint(Fix((.Width - mIconSize) / 2), Fix((.Height - mIconSize) / 2))

            Else
                ' /* top-right corner */
                Set pt = new_BPoint(EM_MARGIN, EM_MARGIN)

            End If

            ' /* draw the icon */
            .DrawScaledImage mIcon, pt, new_BPoint(mIconSize, mIconSize), 128

        End With

    End If

    If mIconOnly Then _
        Exit Sub

'    mView.SetHighColour mSettings.TextColour

Dim xOffset As Long

    If mIconSize > 0 Then _
        xOffset = mIconSize + EM_MARGIN

Dim pr As BRect

    ' /* title */

    With mView

        If mTitleRect.Bottom > 0 Then
            Set pr = mTitleRect.Duplicate

            If (mTextRect.Height = 1) And (Not mIsNumeric) Then
                ' /* no text so centre vertically */
                pr.OffsetBy EM_MARGIN + xOffset, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN + xOffset, EM_MARGIN

            End If

            .SetFont Style.ReadSetting("title_font"), Val(Style.ReadSetting("title_size")), True
            .SetHighColour rgba(0, 0, 0, 102)
            .DrawString mInfo.Title, pr.OffsetByCopy(2, 2)                  '// shadow text
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(96, 96, 96, 127)
            .DrawString mInfo.Title, pr, MFX_SIMPLE_OUTLINE                 '// actual text


        End If

        ' /* text */

        If mIsNumeric Then


        ElseIf mTextRect.Bottom > 1 Then
            Set pr = mTextRect.Duplicate
            If mTitleRect.Height = 1 Then
                ' /* no title so centre vertically */
                pr.OffsetBy EM_MARGIN + xOffset, Fix((.Height - mTextRect.Height) / 2)

            Else
                pr.OffsetBy EM_MARGIN + xOffset, EM_MARGIN + mTitleRect.Height

            End If

            .SetFont Style.ReadSetting("text_font"), Val(Style.ReadSetting("text_size"))
            .SetHighColour rgba(0, 0, 0, 102)
            .DrawString mInfo.Text, pr.OffsetByCopy(2, 2)                   '// shadow text
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(96, 96, 96, 127)
            .DrawString mInfo.Text, pr, MFX_SIMPLE_OUTLINE                  '// actual text


'Dim pr2 As BRect
'Dim dw As Single
'Const RX = 5
'
'        With mView
''            .EnableSmoothing False
'
'            Set pr = new_BRect(0, 0, .Width, 7)
'            pr.InsetBy EM_MARGIN, 0
'            pr.OffsetBy 0, Fix((.Height - pr.Height) / 2)
'            If mIconSize > 0 Then _
'                pr.Left = pr.Left + xOffset
'
'            ' /* draw bar border */
'
'            .SetHighColour rgba(0, 0, 0)
'            .FillRoundRect pr.InsetByCopy(-1, -1), RX, RX, MFX_VERT_GRADIENT
'
''            .SetHighColour rgba(255, 255, 255, 90)
''            .SetLowColour rgba(255, 255, 255, 120)
'
'            .SetHighColour rgba(90, 90, 90)
'            .SetLowColour rgba(120, 120, 120)
'            .FillRoundRect pr, RX, RX, MFX_VERT_GRADIENT
'
'            ' /* take a copy of this - we'll need it if we draw the % value below the bar */
'            Set pr2 = pr.Duplicate
'
'            ' /* draw bar inner */
'            .SetHighColour rgba(96, 96, 96)
''            .FillRoundRect pr.insetbycopy(2, 2), rx, rx
'
'            ' /* if text is non-numeric (during a Scheme Preview
'            '    notification for example), treat as 50% */
'            If Not g_IsNumeric(mInfo.Text) Then
'                dw = 50
'
'            Else
'                dw = Val(mInfo.Text)
'
'            End If
'
'            ' /* constrain value */
'            If dw > 100 Then
'                dw = 100
'
'            ElseIf dw < 0 Then
'                dw = 0
'
'            End If
'
'            If dw > 0 Then
'                ' /* figure out how big the meter should be and draw it */
'                pr.Right = pr.Left + Fix(pr.Width * (dw / 100))
'                .SetHighColour rgba(179, 197, 80)
'                .FillRoundRect pr, RX, RX
'
'                ' /* draw glass effect over meter */
'                .SetHighColour rgba(255, 255, 255, 180)
'                pr.Bottom = pr.Top + 2
'                .FillRect pr.InsetByCopy(2, 0)
'
'            End If
'
'            ' /* draw the actual value (not the passed text) below */
'
''            If mSettings.ShowMeterPercent Then
''                pr2.OffsetBy 0, pr2.Height + 2
''                pr2.Bottom = pr2.Bottom + 20
''
''                .SetFont mSettings.zzTextFont, mSettings.zzTextFontSize, True
''
''                .SetHighColour rgba(0, 0, 0, 102)
''                .DrawString CStr(dw) & "%", pr2.OffsetToCopy(2, 2), MFX_ALIGN_H_CENTER
''
''                .SetHighColour rgba(255, 255, 255)
''                .DrawString CStr(dw) & "%", pr2, MFX_ALIGN_H_CENTER
''
''            End If

        End If
    End With

End Sub

