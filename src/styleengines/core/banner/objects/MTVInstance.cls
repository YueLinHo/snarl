VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTVInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FONT_NAME = "Trebuchet MS"
Private Const FONT_SIZE_TITLE = 20
Private Const FONT_SIZE_TEXT = 12

Dim theView As mfxView
Dim mShowing As Boolean
Dim mYPos As Long

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)

    ' /* on entry:
    '
    '       x, y are set to the position Snarl has determined is best for the notification
    '       Alpha is set to the maximum opacity this notification can reach
    '       Done is set to True
    '
    '   on exit:
    '       In order for Snarl to keep polling AdjustPosition() we must set 'Done' to False
    '       each time until we're done
    '
    ' */

    x = 0

    If mShowing Then

        y = GetSystemMetrics(SM_CYSCREEN) - mYPos

        If mYPos < theView.Height Then
            mYPos = mYPos + 6
            Done = False

'        Else
'            y = theView.Height

        End If

    Else
        y = GetSystemMetrics(SM_CYSCREEN) - mYPos
        
        If mYPos > 0 Then
            mYPos = mYPos - 6
            Done = False

'        Else
'            mYPos = 0

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

    If Visible Then
        mYPos = 0

    Else
        mYPos = theView.Height

    End If

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim cy As Long
Dim cxIcon As Long
Dim xOffset As Long
Dim pr As BRect
Dim hf As Long

Dim fSoftIcon As Boolean
Dim pIcon As mfxBitmap

Const MARGIN = 4

    fSoftIcon = (NotificationInfo.Scheme = LCase$(SCHEME_1))
    xOffset = MARGIN

    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If Not (pIcon Is Nothing) Then
        cxIcon = MIN(pIcon.Height, 128)
        If Not fSoftIcon Then _
            xOffset = xOffset + cxIcon + (2 * MARGIN)

    End If

    Set theView = New mfxView

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), 128 + 16
        .EnableSmoothing False
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        Set pr = .Bounds.Duplicate
        .SetHighColour Val(gSettings.GetValueWithDefault("mv_col1", CStr(rgba(38, 38, 38, 124))))
        .SetLowColour Val(gSettings.GetValueWithDefault("mv_col2", CStr(rgba(30, 30, 30, 124))))
        .FillRect pr, MFX_VERT_GRADIENT

        pr.Bottom = pr.Top
'        .SetHighColour rgba(16, 16, 16)
'        .StrokeLine pr

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 0, 0, 0)
            .SetLowColour rgba(255, 0, 0)
            .FillRect new_BRect(0, Fix(.Bounds.Height / 2), .Bounds.Right, .Bounds.Bottom), MFX_VERT_GRADIENT

        End If

        ' /* notification icon */

        .SetHighColour Val(gSettings.GetValueWithDefault("mv_col3", CStr(rgba(96, 8, 96, 124))))
        .SetLowColour Val(gSettings.GetValueWithDefault("mv_col4", CStr(rgba(96, 8, 96, 124))))
        .FillRect new_BRect(0, 0, cxIcon + (MARGIN * 2) - 1, .Bounds.Bottom), MFX_HORIZ_GRADIENT

'        If fSoftIcon Then
'            .DrawScaledImage pIcon, new_BPoint(MARGIN, 36 + 8), new_BPoint(cxIcon, cxIcon), 255
'
'        Else
'            .DrawScaledImage pIcon, new_BPoint(MARGIN, 8), new_BPoint(cxIcon, cxIcon), 255
'
'        End If

        .DrawScaledImage pIcon, new_BPoint(MARGIN, 8), new_BPoint(cxIcon, cxIcon), IIf(fSoftIcon, 128, 255)

        ' /* title */

        Set pr = .Bounds.Duplicate
        pr.Left = xOffset
        pr.Bottom = pr.Top + 36

'        .SetHighColour rgba(255, 0, 0)
'        .StrokeRect pr

        .SetHighColour rgba(255, 255, 255)
        .SetLowColour rgba(76, 76, 76)
        .SetFont FONT_NAME, FONT_SIZE_TITLE, True
        .DrawString NotificationInfo.Title, pr, MFX_SIMPLE_OUTLINE Or MFX_ALIGN_V_CENTER

        ' /* text */

        pr.Top = pr.Bottom
        pr.Bottom = .Height
        pr.Left = xOffset

'        .SetHighColour rgba(255, 0, 0)
'        .StrokeRect pr

        .SetFont FONT_NAME, FONT_SIZE_TEXT, True
        .DrawString NotificationInfo.Text, pr, MFX_SIMPLE_OUTLINE Or MFX_ALIGN_TOP

        ' /* now, convert this into a bitmap and redraw but with a constant alpha */

Dim pb As mfxBitmap

        Set pb = .ConvertToBitmap()
        .Clear
        .DrawScaledImage pb, , , 230

    End With

End Sub
