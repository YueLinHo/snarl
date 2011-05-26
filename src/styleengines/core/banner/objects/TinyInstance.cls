VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TinyInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FONT_NAME = "Arial" ' "Tahoma"
Private Const FONT_SIZE_TITLE = 7
Private Const FONT_SIZE_TEXT = 7

Dim theView As mfxView
Dim mShowing As Boolean
Dim mYPos As Single
Dim mInfo As notification_info

'Const PRE_DEF_WIDTH = 256 ' 128
'Const IS_DOUBLE_HEIGHT = 0

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
Const STEP_SIZE = 1 / 16

    ' /* on entry:
    '       x, y are set to the position Snarl has determined is best for the notification
    '       Alpha is set to the maximum opacity this notification can reach
    '       Done is set to True
    '
    '   In order for Snarl to keep polling AdjustPosition we must set 'Done' to False each
    '   time until we're done positioning the notification
    '
    ' */

    If InStr(mInfo.Scheme, "left") Then
        x = 64

    ElseIf InStr(mInfo.Scheme, "right") Then
        x = GetSystemMetrics(SM_CXSCREEN) - (theView.Width + 64)

    Else
        x = Fix((GetSystemMetrics(SM_CXSCREEN) - theView.Width) / 2)

    End If



    If InStr(mInfo.Scheme, "top") Then

        If mShowing Then
            If mYPos < 0 Then
                ' /* keep going... */
                mYPos = mYPos + STEP_SIZE
                y = ROUNDUP(mYPos)
                Done = False

            End If

        Else
            ' /* hiding */
            If mYPos < theView.Height Then
                ' /* keep going... */
                mYPos = mYPos + (STEP_SIZE * 2)
                y = -ROUNDUP(mYPos)
                Done = False

            End If

        End If

    Else

        If mShowing Then
            If mYPos > (g_ScreenHeight - theView.Height) Then
                ' /* keep going... */
                mYPos = mYPos - STEP_SIZE
                y = ROUNDUP(mYPos)
                Done = False

            End If

        Else
            ' /* hiding */
            If mYPos < (g_ScreenHeight + theView.Height) Then
                ' /* keep going... */
                mYPos = mYPos + (STEP_SIZE * 2)
                y = ROUNDUP(mYPos)
                Done = False

            End If

        End If
    
    End If


End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
Dim rc As RECT

    mShowing = Visible
    g_GetWorkArea rc

    If InStr(mInfo.Scheme, "bottom") Then
        mYPos = IIf(mShowing, g_ScreenHeight() + theView.Height, g_ScreenHeight() - theView.Height)

    Else
        mYPos = IIf(mShowing, -theView.Height, 0)

    End If

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Const MARGIN_SIZE = 3
Const STAR_SIZE = 12
Const TEXT_GAP = 0
Const rx = 9

    LSet mInfo = NotificationInfo

    ' /* we're (now) a V42 style so notificationinfo->title = App name; notificationinfo->text = content */

Dim pContent As BPackedData
Dim szTitle As String
Dim szText As String
Dim szSubText As String
Dim cStars As Integer
Dim s As Single

    Set pContent = New BPackedData
    With pContent
        .SetTo NotificationInfo.Text

        szTitle = .ValueOf("title")
        szText = .ValueOf("text")
        szSubText = .ValueOf("label-subtext")

        If .Exists("value-percent") Then
            s = Val(.ValueOf("value-percent"))
            If (s < 0) Or (s > 100) Then _
                s = -1

        Else
            s = -1

        End If

        If s > -1 Then
            cStars = Fix(s / 20)

        Else
            cStars = -1

        End If

    End With

    ' /* get icon and calc text offsets from it */

Dim pIcon As mfxBitmap
Dim cxIcon As Long
Dim xOffset As Long
Dim cyFont As Long

    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))

    If is_valid_image(pIcon) Then

        Select Case gSettings.GetValueWithDefault("tiny.iconsize", "1")
        Case "1"
            ' /* small */
            cxIcon = 16

        Case "2"
            ' /* medium */
            cxIcon = 24

        Case Else
            ' /* large */
            cxIcon = 32

        End Select

        cxIcon = MIN(pIcon.Width, cxIcon)
        xOffset = cxIcon + 3                    ' // horizontal text offset

    End If

    Set theView = New mfxView

Dim xIcon As Long
Dim pr As BRect
Dim cy As Long
Dim yMargin As Long
Dim ppd As BPackedData
Dim sz As String

Dim cyContent As Long               ' // the sum total height of title, text, subtext and meter with corresponding gaps

    With theView

        ' /* determine the required height */

        .SetFont FONT_NAME, FONT_SIZE_TITLE ', True

        If gSettings.GetValueWithDefault("tiny.cleartype", "0") = "1" Then
            .TextMode = MFX_TEXT_CLEARTYPE

        Else
            .TextMode = MFX_TEXT_ANTIALIAS

        End If

        ' /* title */

        sz = gSettings.GetValueWithDefault("tiny.font")
        If sz <> "" Then
            Set ppd = New BPackedData
            If ppd.SetTo(sz) Then _
                .SetFont ppd.ValueOf("name"), Val(ppd.ValueOf("size")) + 1 ', (ppd.ValueOf("bold") = "1")

        End If

        cyFont = .StringHeight("A")
        If szTitle <> "" Then _
            cyContent = cyContent + cyFont + TEXT_GAP

        ' /* text */

        sz = gSettings.GetValueWithDefault("tiny.font")
        If sz <> "" Then
            Set ppd = New BPackedData
            If ppd.SetTo(sz) Then _
                .SetFont ppd.ValueOf("name"), Val(ppd.ValueOf("size"))

        End If

        cyFont = .StringHeight("A")
        If szText <> "" Then _
            cyContent = cyContent + cyFont + TEXT_GAP

        ' /* meter */

        If cStars > -1 Then _
            cyContent = cyContent + STAR_SIZE + TEXT_GAP

        ' /* subtext - uses same font as text but drawn in 90% alpha */

        If szSubText <> "" Then _
            cyContent = cyContent + cyFont + TEXT_GAP


        ' /* tidy content size up */

        If cyContent > 0 Then _
            cyContent = cyContent - TEXT_GAP

        cyContent = cyContent + (2 * MARGIN_SIZE)

        ' /* must be at least big enough for the icon */

        cy = MAX(cyContent, cxIcon)
        .SizeTo 320, cy

        If (szText = "") And (szTitle = "") Then
            ' /* special case: icon only */
            xIcon = Fix((.Width - cxIcon) / 2)

        Else
            xIcon = 4

        End If

        If cyContent < cy Then
            yMargin = Fix((cy - cyContent) / 2)

        Else
            yMargin = MARGIN_SIZE

        End If


        ' /* background */

        .EnableSmoothing True
        .SetHighColour gSettings.GetValueWithDefault("tiny.background", CStr(rgba(0, 0, 0)))
        .FillRoundRect .Bounds, rx, rx

        .SetHighColour rgba(255, 255, 255, 127)
        .SetLowColour rgba(0, 0, 0, 127)
        .StrokeFancyRoundRect .Bounds, rx, rx

        ' /* border */

        .SetHighColour gSettings.GetValueWithDefault("tiny.border", CStr(rgba(252, 243, 0)))
        .StrokeRoundRect .Bounds, rx, rx, 2

        ' /* shading */

Dim dCol As Long

        dCol = Val(gSettings.GetValueWithDefault("tiny.shading", CStr(rgba(0, 0, 0))))

        If gSettings.GetValueWithDefault("tiny.invertshading", "0") = "1" Then
            .SetHighColour g_SetRGB32Alpha(dCol, 128)
            .SetLowColour g_SetRGB32Alpha(dCol, 0)

        Else
            .SetHighColour g_SetRGB32Alpha(dCol, 0)
            .SetLowColour g_SetRGB32Alpha(dCol, 128)

        End If

        .FillRoundRect .Bounds, rx, rx, MFX_VERT_GRADIENT

'        .SetHighColour rgba(0, 0, 0)
'        .FillRoundRect .Bounds, RX, RX
'        .FillRect new_BRect(0, 0, .Width - 1, (RX / 2) - 1)         ' // square-up the top

'        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetHighColour rgba(255, 0, 0, 0)
'            .SetLowColour rgba(255, 0, 0)
'            .FillRoundRect .Bounds, rx, rx, MFX_VERT_GRADIENT
''            .SetHighColour rgba(255, 0, 0)
''            .StrokeRoundRect .Bounds, RX, RX
'
'        End If

        ' /* icon */

        If Not (pIcon Is Nothing) Then _
            .DrawScaledImage pIcon, new_BPoint(xIcon, Fix((.Height - cxIcon) / 2)), new_BPoint(cxIcon, cxIcon), 240

        ' /* labels */

        .SetHighColour gSettings.GetValueWithDefault("tiny.foreground", CStr(rgba(255, 255, 255)))


        ' /* create the base rect */

        Set pr = .Bounds.InsetByCopy(MARGIN_SIZE, 0)
        pr.Left = pr.Left + xOffset
        pr.OffsetBy 0, yMargin


        ' /* title */

        sz = gSettings.GetValueWithDefault("tiny.font")
        If sz <> "" Then
            Set ppd = New BPackedData
            If ppd.SetTo(sz) Then _
                .SetFont ppd.ValueOf("name"), Val(ppd.ValueOf("size")) + 1 ', (ppd.ValueOf("bold") = "1")

        End If

        sz = .GetFormattedText(szTitle, pr.Width, True)
        If sz <> "" Then
            .DrawString sz, pr, MFX_ALIGN_LEFT
            pr.OffsetBy 0, .StringHeight("A") + TEXT_GAP

        End If


        ' /* text */

        sz = gSettings.GetValueWithDefault("tiny.font")
        If sz <> "" Then
            Set ppd = New BPackedData
            If ppd.SetTo(sz) Then _
                .SetFont ppd.ValueOf("name"), Val(ppd.ValueOf("size"))

        End If

        sz = .GetFormattedText(szText, pr.Width, True)
        If sz <> "" Then
            .DrawString sz, pr, MFX_ALIGN_LEFT
            pr.OffsetBy 0, .StringHeight("A") + TEXT_GAP

        End If

Dim pImg As mfxBitmap
Dim pImgG As mfxBitmap
Dim i As Integer

        ' /* meter */

        If cStars > -1 Then
            Set pImg = load_image_obj(g_MakePath(App.Path) & "star.png")
            Set pImgG = load_image_obj(g_MakePath(App.Path) & "star.png")
            pImgG.MakeGreyscale

            For i = 0 To 4
                If i <= cStars - 1 Then
                    .DrawScaledImage pImg, pr.TopLeft.OffsetByCopy(i * (STAR_SIZE + 1), 0), new_BPoint(STAR_SIZE, STAR_SIZE)

                Else
                    .DrawScaledImage pImgG, pr.TopLeft.OffsetByCopy(i * (STAR_SIZE + 1), 0), new_BPoint(STAR_SIZE, STAR_SIZE)

                End If

            Next i

            pr.OffsetBy 0, STAR_SIZE + TEXT_GAP

        End If


        ' /* subtext */

        dCol = Val(gSettings.GetValueWithDefault("tiny.foreground", CStr(rgba(255, 255, 255))))
        .SetHighColour g_SetRGB32Alpha(dCol, 230)

        sz = .GetFormattedText(szSubText, pr.Width, True)
        If sz <> "" Then
            .DrawString sz, pr, MFX_ALIGN_LEFT
            pr.OffsetBy 0, .StringHeight("A") + TEXT_GAP

        End If



'        If szTitle = "" Then
'            pr.OffsetBy 0, Fix((pr.Height - cyFont) / 2)
'
'        Else
'            pr.OffsetBy 0, cyFont + 1
'
'        End If
'
'        .DrawString sz, pr

        ' /* glass effect */

        If gSettings.GetValueWithDefault("tiny.useglass", "1") = "1" Then
            .SetHighColour rgba(255, 255, 255, 63)
            .FillRoundRect new_BRect(0, 0, .Bounds.Right, Fix(.Bounds.Height / 2)), rx, rx

        End If

    End With

End Sub

