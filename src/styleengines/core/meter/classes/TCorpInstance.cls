VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCorpInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /*
    '
    '   TCorpInstance.cls -- Represents a single instance of a Corporate Style notification
    '
    '   © 2010 full phat products.  Some Rights Reserved.
    '   Licenced as per the Simplified BSD (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   This style utilises the new (introduced with Snarl R2.4) S_STYLE_V42_CONTENT flag
    '   to allow for more advanced notification content.  Enabling this flag causes Snarl
    '   to send the entire unabridged packed data created by sn41[EZ]Notify() and
    '   sn41[EZ]Update() in notification_info.Text.
    '
    ' */

Const EM_MARGIN = 8
Const EM_TITLE_TEXT = 3
Const PRIMARY_FONT_SIZE = 18

Dim mStyle As TStyle
Dim mInfo As notification_info
Dim mView As mfxView
Dim mRed As Integer

Dim mTitle As String
Dim mTitleRect As BRect
Dim mText As String
Dim mTextRect As BRect
Dim mIcon As mfxBitmap
Dim mIconSize As Long
Dim mIconPos As BPoint
Dim mPercent As Single
Dim mMeterRect As BRect

Dim mBkCol As Long

Dim mLabel() As T_LABEL
Dim mLabels As String
Dim mPrimary As String
Dim mPrimaryRect As BRect
Dim mContent As BPackedData
Dim mBaseTextCol As Long            ' // 24-bit RGB
Dim mCenterIconAndBadge As Boolean

Implements IStyleInstance
Implements IMeterStyle

Private Sub Class_Initialize()

    mRed = 160
    mBkCol = rgba(64, 64, 64)

End Sub

Private Sub IMeterStyle_SetStyle(Style As TStyle)

    Set mStyle = Style

End Sub

Private Sub IStyleInstance_AdjustPosition(X As Long, Y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
Static nStep As Integer
Const STEP_SIZE = 8

    If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then

        If nStep = 0 Then _
            nStep = STEP_SIZE

        mRed = mRed + nStep
        If mRed > 254 Then
            nStep = -STEP_SIZE

        ElseIf mRed < 161 Then
            nStep = STEP_SIZE

        End If

        uDraw
        IStyleInstance_Pulse = True

    End If

End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

'    MsgBox NotificationInfo.Text

    ReDim mLabel(0)
    mLabels = 0

    If (mView Is Nothing) Then
        Set mView = New mfxView
        mView.TextMode = MFX_TEXT_ANTIALIAS
        mView.TextMode = MFX_TEXT_CLEARTYPE

    End If

    ' /* notificationinfo->scheme is translated to lowercase by Snarl */

    LSet mInfo = NotificationInfo

    ' /* cache icon */

    mIconSize = 0
    Set mIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If is_valid_image(mIcon) Then _
        mIconSize = MIN(mIcon.Height, 48)

    ' /* pre-configuration */

Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

    If ppd.Exists("title") Then _
        mTitle = ppd.ValueOf("title")

    If ppd.Exists("text") Then _
        mText = ppd.ValueOf("text")

'    ' /* sub-text */
'
'    If ppd.Exists("label-subtext") Then _
'        mSubtext = ppd.ValueOf("label-subtext")

    ' /* primary text - maximum 5 chars! */

    If ppd.Exists("style-corp-tagline") Then _
        mPrimary = g_SafeLeftStr(ppd.ValueOf("style-corp-tagline"), 5)

    ' /* colour */

    If NotificationInfo.Scheme = "light" Then
        mBaseTextCol = vbBlack

    Else
        mBaseTextCol = vbWhite

    End If

    ' /* background colour */

    If ppd.Exists("colour-background") Then
        If ppd.ValueOf("colour-background") <> "" Then _
            mBkCol = g_RGBAFromString(ppd.ValueOf("colour-background"))

    ElseIf NotificationInfo.Scheme = "light" Then
        mBkCol = rgba(247, 247, 247)

    Else
        mBkCol = rgba(64, 64, 64)

    End If

    ' /* percent */

    If ppd.Exists("value-percent") Then
        mPercent = Val(ppd.ValueOf("value-percent"))
        If (mPercent < 0#) Or (mPercent > 100#) Then _
            mPercent = -1#

    ElseIf (mStyle.ReadSetting("show_graphically") = "1") And (g_IsNumeric(mText)) Then
        mPercent = Val(mText)
        mText = ""
        If (mPercent < 0#) Or (mPercent > 100#) Then _
            mPercent = -1#

    Else
        mPercent = -1#                          ' // means we don't need a meter...

    End If

    Set mContent = ppd

    ' /* calculate notifcation size and internal measurements */

    uCalcSize
    uDraw

End Sub

Private Sub uCalcSize()

    If (mView Is Nothing) Then _
        Exit Sub

Dim cyNotification As Long

    cyNotification = (2 * EM_MARGIN) + mIconSize
    Set mTitleRect = new_BRect(0, 0, 0, 0)
    Set mTextRect = new_BRect(0, 0, 0, 0)

    mView.SizeTo 260, cyNotification

Dim Y As Long
Dim prTemp As BRect
Dim pr As BRect
Dim fNeedGap As Boolean
Dim cxPrimary As Long
Dim xPos As Long
Dim cxMaster As Long

    ' /* measure title and text to see how big the notification needs to be */

    With mView

        .TextMode = MFX_TEXT_ANTIALIAS
        cxMaster = .Width

        Y = EM_MARGIN

        ' /* create the base bounding rect - this is used to get the title and text height
        '    based on a fixed width */

        Set pr = new_BRect(0, 0, .Width - 1, cyNotification - 1).InsetByCopy(EM_MARGIN, EM_MARGIN)

        ' /* measure primary text */

        If mPrimary <> "" Then
            .SetFont mStyle.ReadSetting("text_font"), PRIMARY_FONT_SIZE
            cxPrimary = .StringWidth(mPrimary) + EM_TITLE_TEXT
            Set mPrimaryRect = new_BRect(EM_MARGIN, Y, EM_MARGIN + cxPrimary - 1, Y + .StringHeight(mPrimary) - 1)
            If mIconSize > 0 Then _
                mPrimaryRect.OffsetBy 48 + EM_TITLE_TEXT, 0

        End If

        ' /* get the correct left edge */

        xPos = EM_MARGIN + cxPrimary
        If mIconSize > 0 Then _
            xPos = xPos + 48 + EM_TITLE_TEXT

        ' /* measure title */

        If mTitle <> "" Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size"), True
            Set prTemp = pr.Duplicate
            prTemp.Left = xPos
            prTemp.Bottom = 16384
            .MeasureString mTitle, prTemp, prTemp
            Set mTitleRect = new_BRect(xPos, Y, .Width - EM_MARGIN - 1, Y + prTemp.Height - 1)
            fNeedGap = True

            Y = Y + prTemp.Height

        End If

        ' /* measure text */

        If mText <> "" Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            Set prTemp = pr.Duplicate
            prTemp.Left = xPos
            prTemp.Bottom = 16384
            .MeasureString mText, prTemp, prTemp

            If fNeedGap Then _
                Y = Y + EM_TITLE_TEXT

            Set mTextRect = new_BRect(xPos, Y, .Width - EM_MARGIN - 1, Y + prTemp.Height - 1)

            fNeedGap = True
            Y = Y + prTemp.Height

        End If

        ' /* meter */

        If mPercent > -1 Then
            If fNeedGap Then _
                Y = Y + EM_TITLE_TEXT

            Set mMeterRect = new_BRect(xPos, Y, .Width - EM_MARGIN - 1, Y + 10 - 1)
            fNeedGap = True
            Y = Y + mMeterRect.Height

        Else
            Set mMeterRect = Nothing

        End If

        ' /* label-* */

        .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")

Dim szn As String
Dim szv As String

        With mContent
            .Rewind
            Do While .GetNextItem(szn, szv)
                If (g_SafeLeftStr(szn, 6) = "label-") And (szv <> "") Then
                    Set prTemp = pr.Duplicate
                    prTemp.Left = xPos
                    prTemp.Bottom = 16384
                    mView.MeasureString szv, prTemp, prTemp

                    If fNeedGap Then _
                        Y = Y + EM_TITLE_TEXT

                    mLabels = mLabels + 1
                    ReDim Preserve mLabel(mLabels)
                    With mLabel(mLabels)
                        .Text = szv
                        Set .Frame = new_BRect(xPos, Y, mView.Width - EM_MARGIN - 1, Y + prTemp.Height - 1)

                    End With

                    fNeedGap = True
                    Y = Y + prTemp.Height

                End If
            Loop
        End With

        Y = Y + EM_MARGIN
        mCenterIconAndBadge = False

        If (mText = "") And (mTitle = "") And (mLabels = 0) And (mMeterRect Is Nothing) And (mPrimary = "") Then
            ' /* special case: icon only */
'            cxMaster = 120

        ElseIf (mText = "") And (mTitle = "") And (mLabels = 0) And (mMeterRect Is Nothing) And (mIcon Is Nothing) Then
            ' /* special case: tagline only */
            cyNotification = mPrimaryRect.Height + (2 * EM_MARGIN)

        ElseIf (mText = "") And (mTitle = "") And (mLabels = 0) And (mMeterRect Is Nothing) Then
            ' /* special case: icon & badge only */
            cyNotification = MAX(mIconSize, mPrimaryRect.Height) + (2 * EM_MARGIN)
            mCenterIconAndBadge = True

        End If

        ' /* size the view */

        .SizeTo cxMaster, MAX(Y, cyNotification)

        ' /* icon position */

        Set mIconPos = new_BPoint(EM_MARGIN + Fix((48 - mIconSize) / 2), Fix((.Height - mIconSize) / 2))

        ' /* primary text */

        If Not (mPrimaryRect Is Nothing) Then _
            mPrimaryRect.OffsetTo mPrimaryRect.Left, Fix((.Height - mPrimaryRect.Height) / 2)

    End With

End Sub

Private Sub uDraw()
Const RX = 10

    With mView
        .Clear
        .EnableSmoothing True
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        If (mInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(mRed, 0, 0)
            .FillRoundRect .Bounds, RX, RX

        Else
            .SetHighColour mBkCol
            .FillRoundRect .Bounds, RX, RX

        End If

'        .SetHighColour rgba(255, 255, 255, 20)
'        .SetLowColour rgba(255, 255, 255, 0)
'        .FillRoundRect new_BRect(0, 0, .Bounds.Right, Fix(.Bounds.Height / 2) - 1), RX, RX, MFX_VERT_GRADIENT
'
'        .SetHighColour rgba(255, 255, 255, 20)
'        .SetLowColour rgba(255, 255, 255, 0)
'        .FillRoundRect new_BRect(0, 0, .Bounds.Right, .Bounds.Bottom), RX, RX, MFX_VERT_GRADIENT

        If mInfo.Scheme = "light" Then
            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 32)

        Else
            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 128)

        End If

        .FillRoundRect new_BRect(0, 0, .Bounds.Right, .Bounds.Bottom), RX, RX, MFX_VERT_GRADIENT


'        .SetHighColour rgba(255, 255, 255, 200)
'        .StrokeLine new_BRect(RX, 0, .Bounds.Right - RX, 0)

'        .SetHighColour rgba(255, 255, 255, 150)
'        .StrokeRoundRect .Bounds.InsetByCopy(0, 0), RX, RX

        ' /* title */

        If mTitle <> "" Then
'            .SetHighColour rgba(255, 0, 0)
'            .StrokeRect mTitleRect

            .SetFont mStyle.ReadSetting("text_font"), Val(mStyle.ReadSetting("text_size")), True
            .SetHighColour g_MakeRGBA(mBaseTextCol, 250)

            If (mText = "") And (mLabels = 0) And (mMeterRect Is Nothing) Then _
                mTitleRect.OffsetTo mTitleRect.Left, Fix((.Height - mTitleRect.Height) / 2)

            .DrawString mTitle, mTitleRect

        End If

        ' /* text */

        If mText <> "" Then
'            .SetHighColour rgba(255, 255, 0)
'            .StrokeRect mTextRect

            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            .SetHighColour g_MakeRGBA(mBaseTextCol, 200)

            If (mTitle = "") And (mLabels = 0) And (mMeterRect Is Nothing) Then _
                mTextRect.OffsetTo mTextRect.Left, Fix((.Height - mTextRect.Height) / 2)

            .DrawString mText, mTextRect

        End If

Dim cxIconBadge As Long

        If mCenterIconAndBadge Then _
            cxIconBadge = mIconSize + EM_TITLE_TEXT + mPrimaryRect.Width

        ' /* icon */

        If Not (mIcon Is Nothing) Then
        
            If mCenterIconAndBadge Then
                mIconPos.OffsetTo Fix((.Width - cxIconBadge) / 2), mIconPos.Y

            ElseIf (mTitle = "") And (mLabels = 0) And (mMeterRect Is Nothing) And (mText = "") And (mPrimary = "") Then
                ' /* just icon */
                mIconPos.OffsetTo Fix((.Width - mIconSize) / 2), mIconPos.Y

            End If

            .DrawScaledImage mIcon, mIconPos, new_BPoint(mIconSize, mIconSize)

        End If

        ' /* primary */

        If mPrimary <> "" Then
'            .SetHighColour rgba(0, 0, 255)
'            .StrokeRect mPrimaryRect

            If mCenterIconAndBadge Then
                mPrimaryRect.OffsetTo Fix((.Width - cxIconBadge) / 2) + mIconSize + EM_TITLE_TEXT, mPrimaryRect.Top

            ElseIf (mTitle = "") And (mLabels = 0) And (mMeterRect Is Nothing) And (mText = "") And (mIcon Is Nothing) Then
                ' /* just badge */
                mPrimaryRect.OffsetTo Fix((.Width - mPrimaryRect.Width) / 2), mPrimaryRect.Top

            End If

            .SetHighColour rgba(255, 0, 0)
'            .StrokeRect mPrimaryRect

            .SetFont mStyle.ReadSetting("text_font"), PRIMARY_FONT_SIZE
            .SetHighColour g_MakeRGBA(mBaseTextCol, 250)
            .DrawString mPrimary, mPrimaryRect, MFX_ALIGN_V_CENTER

        End If

        ' /* meter */

Dim pr As BRect

        If Not (mMeterRect Is Nothing) Then

            If (mTitle = "") And (mText = "") And (mLabels = 0) Then _
                mMeterRect.OffsetTo mMeterRect.Left, Fix((.Height - mMeterRect.Height) / 2)

            .SetHighColour rgba(0, 0, 0, 60)
            .FillRoundRect mMeterRect, 4, 4

            .SetHighColour rgba(255, 255, 255, 90)
            Set pr = mMeterRect.InsetByCopy(1, 1)
            pr.Right = pr.Left + ((mPercent / 100) * pr.Width) - 1
            If pr.Width > 3 Then _
                .FillRoundRect pr, 4, 4

        End If

        ' /* subtext */

Dim i As Long

        If mLabels > 0 Then
            .SetFont mStyle.ReadSetting("text_font"), mStyle.ReadSetting("text_size")
            .SetHighColour g_MakeRGBA(mBaseTextCol, 250)

            For i = 1 To mLabels
                Set pr = mLabel(i).Frame.Duplicate
                If (mTitle = "") And (mText = "") And (mMeterRect Is Nothing) Then _
                    pr.OffsetTo pr.Left, Fix((.Height - pr.Height) / 2)

                .DrawString mLabel(i).Text, pr

            Next i

        End If

'        .SetHighColour rgba(255, 255, 255)
'        .DrawString "'" & mText & "'", .Bounds

    End With

End Sub

