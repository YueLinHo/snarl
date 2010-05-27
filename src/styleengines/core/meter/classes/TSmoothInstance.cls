VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSmoothInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const EM_TITLE_TEXT = 3
Const EM_MARGIN = 10
Const EM_LEFT_MARGIN = 48

Dim mStyle As TStyle

Dim mInfo As notification_info
Dim mIconSize As Long
Dim mIcon As mfxBitmap
Dim mView As mfxView

Dim mIconOnly As Boolean
Dim mIsNumeric As Boolean
Dim mTitleRect As BRect
Dim mTextRect As BRect

Implements IStyleInstance
Implements IMeterStyle

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, Y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    If (mView Is Nothing) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS

    End If

    LSet mInfo = NotificationInfo

    ' /* clear cached icon */

    Set mIcon = Nothing
    mIconSize = 0

Dim pi As MImage

    ' /* cache icon */

    Set pi = load_image_obj(NotificationInfo.Icon)
    If is_valid_image(pi) Then _
        Set mIcon = create_bitmap_from_image(pi)

    ' /* pre-configuration */

    mIconOnly = ((mInfo.Title = "") And (mInfo.Text = "")) 'Or (InStr(mInfo.Scheme, "icon only") > 0))
    mIsNumeric = (g_IsNumeric(mInfo.Text) And (Val(mStyle.ReadSetting("show_graphically"))))

    ' /* calculate notifcation size and internal measurements */

    uCalcSize

    ' /* now we have the final notification height, calculate the size of the icon */

    If is_valid_image(mIcon) Then _
        mIconSize = MIN(mIcon.Height, mView.Height + (2 * EM_MARGIN))

    uDraw

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

        Set pr = new_BRect(0, 0, 250 - 1, cyNotification - 1)
        pr.InsetBy EM_MARGIN, EM_MARGIN
        pr.Left = EM_LEFT_MARGIN

        ' /* measure title */

        If mInfo.Title <> "" Then
            .SetFont mStyle.ReadSetting("title_font"), mStyle.ReadSetting("title_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Title, prTemp, prTemp
            Set mTitleRect = new_BRect(0, 0, prTemp.Width + 0, prTemp.Height + 0)

        End If

        ' /* measure text */

        If (mInfo.Text <> "") And (Not mIsNumeric) Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Bottom = 16384
            .MeasureString mInfo.Text, prTemp, prTemp
            Set mTextRect = new_BRect(0, 0, prTemp.Width + 0, prTemp.Height + 0)

        End If

        ' /* size the view */

        .SizeTo 250, MAX(mTitleRect.Height + mTextRect.Height + (2 * EM_MARGIN) + EM_TITLE_TEXT, cyNotification)

    End With

End Sub

Private Sub uDraw()

    ' /* draw the background */

    With mView
        .Clear
        .EnableSmoothing False

        ' /* background */

        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            ' /* priority notifications always get a red background */
            .SetHighColour rgba(223, 0, 0)

        Else
            .SetHighColour mStyle.GetSchemeColour(mInfo.Scheme, "background")

        End If

        .FillRect .Bounds

        .SetHighColour rgba(255, 255, 255, 96)
        .SetLowColour rgba(255, 255, 255, 0)
        .FillRect .Bounds, MFX_VERT_GRADIENT

        .SetHighColour rgba(0, 0, 0, 0)
        .SetLowColour rgba(0, 0, 0, 96)
        .FillRect .Bounds, MFX_VERT_GRADIENT

    End With

    ' /* draw the foreground */

    uDrawForeground

End Sub

Private Sub uDrawForeground()

    ' /* icon */

    If mIconSize > 0 Then _
        mView.DrawScaledImage mIcon, new_BPoint(-(mIconSize * 0.25), Fix((mView.Height - mIconSize) / 2)), new_BPoint(mIconSize, mIconSize), 128

    If mIconOnly Then _
        Exit Sub

Dim prLabel As BRect
Dim pr As BRect
Dim dw As Single
Dim xOffs As Long

    With mView

        ' /* title */

        If mIconSize > 0 Then
            xOffs = EM_LEFT_MARGIN

        Else
            xOffs = EM_MARGIN

        End If

        If mTitleRect.Bottom > 0 Then
            Set pr = mTitleRect.Duplicate
            If (mTextRect.Height = 1) And (Not mIsNumeric) Then
                ' /* no text so centre vertically */
                pr.OffsetBy xOffs, Fix((.Height - mTitleRect.Height) / 2)

            Else
                pr.OffsetBy xOffs, EM_MARGIN

            End If

            .SetFont mStyle.ReadSetting("title_font"), Val(mStyle.ReadSetting("title_size")), True
            .SetHighColour rgba(255, 255, 255)
            .DrawString mInfo.Title, pr

        End If

        ' /* text */

        If mIsNumeric Then
            ' /* draw a bar instead */
            Set pr = new_BRect(xOffs, 0, .Width - EM_MARGIN, 10 - 1)
            pr.OffsetBy 0, Fix((.Height - pr.Height) / 2)
            .SetHighColour rgba(0, 0, 0, 96)
            .EnableSmoothing True
            .FillRoundRect pr, 4, 4
            Set prLabel = pr.Duplicate

            dw = Val(mInfo.Text)
            dw = MAX(MIN(dw, 100), 0)
            If dw > 4 Then
                .SetHighColour rgba(255, 255, 255, 202)
                pr.InsetBy 2, 2
                pr.Right = pr.Left + Fix(pr.Width * (dw / 100))
                .FillRoundRect pr, 4, 4

            End If

            ' /* display value? */

            If Val(mStyle.ReadSetting("show_percent")) Then
                .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
                .SetHighColour rgba(255, 255, 255)
                prLabel.OffsetBy 0, prLabel.Height
                prLabel.Bottom = prLabel.Top + .StringHeight("A")
                .DrawString CStr(dw) & "%", prLabel, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER

            End If


            .EnableSmoothing False

        ElseIf mTextRect.Bottom > 1 Then
            ' /* have some text to draw */

            Set pr = mTextRect.Duplicate
            If mTitleRect.Height = 1 Then
                ' /* no title so centre vertically */
                pr.OffsetBy xOffs, Fix((.Height - mTextRect.Height) / 2)

            Else
                pr.OffsetBy xOffs, EM_MARGIN + mTitleRect.Height
                .SetHighColour rgba(255, 255, 255)
                .StrokeLine new_BRect(pr.Left, pr.Top - 1, .Bounds.Right - EM_MARGIN, pr.Top - 1)
                pr.OffsetBy 0, EM_TITLE_TEXT

            End If

            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            .SetHighColour rgba(255, 255, 255)
            .DrawString mInfo.Text, pr

        End If

    End With

End Sub


