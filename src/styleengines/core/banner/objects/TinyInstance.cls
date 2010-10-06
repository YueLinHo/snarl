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

Private Const FONT_NAME = "Tahoma"
Private Const FONT_SIZE_TITLE = 7
Private Const FONT_SIZE_TEXT = 7

Dim theView As mfxView
Dim mShowing As Boolean
Dim mYPos As Single

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
Const STEP_SIZE = 0.125

    ' /* on entry:
    '       x, y are set to the position Snarl has determined is best for the notification
    '       Alpha is set to the maximum opacity this notification can reach
    '       Done is set to True
    '
    '   In order for Snarl to keep polling AdjustPosition we must set 'Done' to False each
    '   time until we're done positioning the notification
    '
    ' */

    x = GetSystemMetrics(SM_CXSCREEN) - (theView.Width + 32)

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
            mYPos = mYPos + STEP_SIZE
            y = -ROUNDUP(mYPos)
            Done = False

        End If

    End If

End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)

    mShowing = Visible

    If mShowing Then
        mYPos = -theView.Height

    Else
        mYPos = 0

    End If

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Const RX = 12

    ' /* get icon and calc text offsets from it */

Dim cxIcon As Long
Dim xOffset As Long
Dim pIcon As mfxBitmap

    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))

    If Not (pIcon Is Nothing) Then
        If (pIcon.Width > 0) And (pIcon.Height > 0) Then
            cxIcon = MIN(pIcon.Width, 24)                   ' // can't be bigger than 24px
            xOffset = cxIcon + 4

        End If
    End If

    Set theView = New mfxView

Dim nHeight As Long
Dim nWidth As Long
Dim pr As BRect
Dim cy As Long

Dim z As Long

    ' /* determine overall width and height based on scheme */

    Select Case NotificationInfo.Scheme
    Case "extra small"
        ' /* width flexes to keep height down */

        nHeight = 4             ' // 2 x 2px border

        With theView
            ' /* calc title size */
            .SetFont FONT_NAME, FONT_SIZE_TITLE, True
            .MeasureString NotificationInfo.Title, new_BRect(0, 0, 999, 999), pr
            cy = pr.Height
            nHeight = nHeight + cy
            nWidth = pr.Width

            ' /* calc text size */
            .SetFont FONT_NAME, FONT_SIZE_TEXT, True
            .MeasureString NotificationInfo.Text, new_BRect(0, 0, 999, 999), pr
            nHeight = nHeight + pr.Height
            nWidth = MAX(nWidth, pr.Width) + xOffset + 4 + 2

        End With

    Case "small"
        ' /* width is fixed - height is variable */

        nWidth = 150

        With theView
            ' /* calc title size */
            .SetFont FONT_NAME, FONT_SIZE_TITLE, True
            .MeasureString NotificationInfo.Title, new_BRect(xOffset + 2, 0, nWidth - 2, 999), pr
            cy = pr.Height
            nHeight = nHeight + cy

            ' /* calc text size */
            .SetFont FONT_NAME, FONT_SIZE_TEXT, True
            .MeasureString NotificationInfo.Text, new_BRect(xOffset + 2, 0, nWidth - 2, 999), pr
            z = pr.Height
            nHeight = nHeight + pr.Height

        End With

    End Select


    With theView
        .SizeTo nWidth, nHeight
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        .SetHighColour rgba(0, 0, 0)
        .FillRoundRect .Bounds, RX, RX
        .FillRect new_BRect(0, 0, .Width - 1, (RX / 2) - 1)         ' // square-up the top

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(0, 0, 0)
            .SetLowColour rgba(255, 0, 0)
'            .FillRoundRect new_BRect(0, Fix(.Height / 2), .Width - 1, .Height - 1), RX, RX, MFX_VERT_GRADIENT
            .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

        Else

        End If

        .EnableSmoothing False

        ' /* icon */

        If Not (pIcon Is Nothing) Then _
            .DrawScaledImage pIcon, new_BPoint(2, 2), new_BPoint(cxIcon, cxIcon), 220


        ' /* labels */

'        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetHighColour rgba(255, 0, 0)
'            .SetLowColour rgba(255, 255, 255)
'
'        Else
            .SetHighColour rgba(255, 255, 255)
            .SetLowColour rgba(0, 0, 0)
'
'        End If

        .SetFont FONT_NAME, FONT_SIZE_TITLE, True

        ' /* title */

        Set pr = .Bounds.InsetByCopy(2, 2)
        pr.Left = pr.Left + xOffset
        .DrawString NotificationInfo.Title, pr, MFX_SIMPLE_OUTLINE

        ' /* text */

        Set pr = .Bounds.InsetByCopy(2, 2)
        pr.Left = pr.Left + xOffset
        pr.Top = cy

        .SetFont FONT_NAME, FONT_SIZE_TEXT, True
        .DrawString NotificationInfo.Text, pr, MFX_SIMPLE_OUTLINE

    End With

End Sub

