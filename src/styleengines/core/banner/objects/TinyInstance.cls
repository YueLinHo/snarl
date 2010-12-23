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

Const PRE_DEF_WIDTH = 256 ' 128
Const IS_DOUBLE_HEIGHT = 0

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
            If mYPos < -1 Then
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
Const RX = 5

    With NotificationInfo
        .Title = Replace$(.Title, vbCrLf, " ")
        .Text = Replace$(.Text, vbCrLf, " ")

    End With

    LSet mInfo = NotificationInfo

    ' /* get icon and calc text offsets from it */

Dim cxIcon As Long
Dim xOffset As Long
Dim pIcon As mfxBitmap
Dim cyFont As Long

    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))

    If is_valid_image(pIcon) Then
        cxIcon = MIN(pIcon.Width, 24)                   ' // can't be bigger than 24px
        xOffset = cxIcon + 2

    End If

    Set theView = New mfxView

Dim xIcon As Long
Dim pr As BRect
Dim cy As Long
Dim yMargin As Long
Dim ppd As BPackedData
Dim sz As String

    With theView

        ' /* determine the required height */

        .SetFont FONT_NAME, FONT_SIZE_TITLE, True

        If gSettings.GetValueWithDefault("tiny.cleartype", "0") = "1" Then
            .TextMode = MFX_TEXT_CLEARTYPE

        Else
            .TextMode = MFX_TEXT_ANTIALIAS

        End If

        sz = gSettings.GetValueWithDefault("tiny.font")
        If sz <> "" Then
            Set ppd = New BPackedData
            If ppd.SetTo(sz) Then _
                .SetFont ppd.ValueOf("name"), ppd.ValueOf("size"), (ppd.ValueOf("bold") = "1")

        End If

        cyFont = .StringHeight("A")

        If mInfo.Title <> "" Then _
            cy = cy + cyFont

        If mInfo.Text <> "" Then _
            cy = cy + cyFont


        cy = MAX(cy, 24)

        cy = cy + (2 * 2)               ' // 2px margin

        If (mInfo.Text = "") And (mInfo.Title = "") Then
            ' /* special case: icon only */
            .SizeTo 48 + (2 * 2), cy
            xIcon = Fix((.Width - 24) / 2)

        Else

            Select Case gSettings.GetValueWithDefault("tiny.size", "1")
            Case "1"
                ' /* tiny */
                .SizeTo 192, cy

            Case "3"
                ' /* large */
                .SizeTo 512, cy

            Case Else
                ' /* medium */
                .SizeTo 384, cy

            End Select

            xIcon = 2

        End If

        yMargin = Fix((.Height - (cyFont * 2)) / 2)
'        yMargin = 2

        ' /* background */

        .EnableSmoothing True
        .SetHighColour gSettings.GetValueWithDefault("tiny.background", CStr(rgba(0, 0, 0)))
        .FillRoundRect .Bounds, RX, RX

        .SetHighColour rgba(255, 255, 255, 16)
        .SetLowColour rgba(255, 255, 255, 0)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

'        .SetHighColour rgba(0, 0, 0)
'        .FillRoundRect .Bounds, RX, RX
'        .FillRect new_BRect(0, 0, .Width - 1, (RX / 2) - 1)         ' // square-up the top

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 0, 0, 0)
            .SetLowColour rgba(255, 0, 0)
            .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT
'            .SetHighColour rgba(255, 0, 0)
'            .StrokeRoundRect .Bounds, RX, RX

        End If

        ' /* icon */

        If Not (pIcon Is Nothing) Then _
            .DrawScaledImage pIcon, new_BPoint(xIcon, 2), new_BPoint(cxIcon, cxIcon), 240

        ' /* labels */

        .SetHighColour gSettings.GetValueWithDefault("tiny.foreground", CStr(rgba(255, 255, 255)))

        ' /* title */

        Set pr = .Bounds.InsetByCopy(2, yMargin)
        pr.Left = pr.Left + xOffset
        sz = .GetFormattedText(mInfo.Title, pr.Width, True)
        If mInfo.Text = "" Then _
            pr.OffsetBy 0, Fix((pr.Height - cyFont) / 2)

        .DrawString sz, pr, IIf((mInfo.Text = "") And Not (is_valid_image(pIcon)), MFX_ALIGN_H_CENTER, MFX_ALIGN_LEFT)

        ' /* text */

        Set pr = .Bounds.InsetByCopy(2, yMargin)
        pr.Left = pr.Left + xOffset
        sz = .GetFormattedText(mInfo.Text, pr.Width, True)
        If mInfo.Title = "" Then
            pr.OffsetBy 0, Fix((pr.Height - cyFont) / 2)

        Else
            pr.OffsetBy 0, cyFont

        End If

        .DrawString sz, pr

        .SetHighColour rgba(0, 0, 0, 102)
        .StrokeRoundRect .Bounds, RX, RX

    End With

End Sub

