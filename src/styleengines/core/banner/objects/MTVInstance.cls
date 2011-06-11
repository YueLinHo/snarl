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

Private Const FONT_NAME = "Arial Rounded MT Bold"
Private Const FONT_SIZE_TITLE = 26
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

        Else
            ' /* done */
            y = GetSystemMetrics(SM_CYSCREEN) - theView.Height
            Done = True

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
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.SetTo NotificationInfo.Text

Dim fAdvanced As Boolean

    fAdvanced = (ppd.Exists("value-rating")) Or (ppd.Exists("value-percent")) Or (ppd.Exists("label-artist")) Or (ppd.Exists("label-album"))

Dim pIcon As mfxBitmap
Dim xOffset As Long
Dim cxIcon As Long

Const TEXT_MARGIN = 4
Const MARGIN = 16
Const H_SPLIT = 500

    xOffset = MARGIN
    cxIcon = 128 '// IIf(NotificationInfo.Scheme = LCase$(SCHEME_2), 64, 128)
    Set pIcon = style_MakeSquareImage(load_image_obj(NotificationInfo.Icon))
    If Not (pIcon Is Nothing) Then _
        xOffset = xOffset + cxIcon + MARGIN


Dim pb As mfxBitmap

    With New mfxView
        .EnableSmoothing True
        .SizeTo 500, 500
        .SetHighColour rgba(255, 255, 255, 10)
        .SetHighColour rgba(255, 255, 255, 255)
        .FillRect .Bounds ', MFX_VERT_GRADIENT
        Set pb = .ConvertToBitmap

    End With

Dim pr As BRect
Dim i As Integer

    Set theView = New mfxView

    With theView
        .SizeTo GetSystemMetrics(SM_CXSCREEN), cxIcon + 16
        .EnableSmoothing False
        .TextMode = MFX_TEXT_ANTIALIAS

        ' /* background */

        .SetHighColour rgba(34, 34, 34)
        .FillRect new_BRect(0, 0, .Bounds.Right - H_SPLIT, .Bounds.Bottom)

        .SetHighColour rgba(34, 34, 34)
        .SetLowColour rgba(0, 102, 127)
        .FillRect new_BRect(.Bounds.Right - H_SPLIT, 0, .Bounds.Right, .Bounds.Bottom), MFX_HORIZ_GRADIENT

        .SetHighColour rgba(0, 0, 0, 32)

        For i = 0 To .Bounds.Bottom Step 4
            .StrokeLine new_BRect(0, i, .Bounds.Right, i)

        Next i

'        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
'            .SetHighColour rgba(255, 0, 0, 0)
'            .SetLowColour rgba(255, 0, 0)
'            .FillRect new_BRect(0, Fix(.Bounds.Height / 2), .Bounds.Right, .Bounds.Bottom), MFX_VERT_GRADIENT
'
'        End If

        ' /* notification icon */

        .DrawScaledImage pIcon, new_BPoint(MARGIN, 8), new_BPoint(cxIcon, cxIcon), 250


        ' /* TESTING TESTING TESTING */
'fAdvanced = True
'        ppd.SetTo "title::Last Train To Trancentral#?value-rating::4/5#?label-artist::The KLF#?label-album::The White Room#?label-duration::3:47#?label-copyright::© 1991 KLF COMMUNICATIONS"
        ' /* TESTING TESTING TESTING */


Dim cy As Long
Dim j As Integer
Const TEXT_FIX = 2
Dim sz As String

        If fAdvanced Then
            ' /* work bottom up... */

            i = -1

            If ppd.Exists("value-rating") Then
                i = uRatingPercent(ppd.ValueOf("value-rating"))
                Debug.Print "rating % = " & CStr(i)

            ElseIf ppd.Exists("value-percent") Then
                i = g_SafeLong(ppd.ValueOf("value-percent"))
                If (i < 0) Or (i > 100) Then _
                    i = -1

            End If

            ' /* starting point */
            Set pr = new_BRect(xOffset, .Bounds.Bottom - TEXT_MARGIN - 16 - 1, .Bounds.Right - MARGIN - 1, .Bounds.Bottom - TEXT_MARGIN - 1)

            ' /* rating stars */
            If i > 0 Then
                Set pb = load_image_obj(g_MakePath(App.Path) & IIf(g_IsIDE, "bin\", "") & "star.png")
                i = Fix(i / 20)
                For j = 0 To (i - 1)
                    .DrawScaledImage pb, new_BPoint(pr.Left + (17 * j), pr.Top), new_BPoint(16, 16)

                Next j
                pr.OffsetBy 0, -17

            End If

            .SetFont FONT_NAME, FONT_SIZE_TEXT - 2
            .SetHighColour rgba(0, 204, 254)

            ' /* copyright */
            If ppd.ValueOf("label-copyright") <> "" Then
                Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
                pr.OffsetBy 0, -cy
                .DrawString ppd.ValueOf("label-copyright"), pr, MFX_ALIGN_TOP
                cy = pr.Height - TEXT_FIX

            End If

            .SetFont FONT_NAME, FONT_SIZE_TEXT

            ' /* album */
            If ppd.ValueOf("label-album") <> "" Then
                Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
                pr.OffsetBy 0, -cy
                .DrawString ppd.ValueOf("label-album"), pr, MFX_ALIGN_TOP
                cy = pr.Height - TEXT_FIX

            End If

            .SetFont FONT_NAME, FONT_SIZE_TEXT, True

            ' /* artist */
            If ppd.ValueOf("label-artist") <> "" Then
                Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
                pr.OffsetBy 0, -cy
                .DrawString ppd.ValueOf("label-artist"), pr, MFX_ALIGN_TOP
                cy = pr.Height - TEXT_FIX

            End If


            ' /* title */
            If ppd.ValueOf("title") <> "" Then _
                sz = ppd.ValueOf("title")

            If ppd.ValueOf("label-duration") <> "" Then _
                sz = sz & IIf(sz <> "", " ", "") & "(" & ppd.ValueOf("label-duration") & ")"

            If sz <> "" Then
                .SetFont FONT_NAME, FONT_SIZE_TITLE, True
                Set pr = new_BRect(xOffset, pr.Bottom - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, pr.Bottom - 1)
                pr.OffsetBy 0, -cy
                .SetHighColour rgba(255, 255, 255)
                .DrawString sz, pr, MFX_ALIGN_TOP

            End If


        Else
            ' /* work bottom up, so text part first... */
            If ppd.ValueOf("text") <> "" Then
                .SetFont FONT_NAME, FONT_SIZE_TEXT, True
                Set pr = new_BRect(xOffset, .Bounds.Bottom - TEXT_MARGIN - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, .Bounds.Bottom - TEXT_MARGIN - 1)
                .SetHighColour rgba(0, 204, 254)
                .DrawString ppd.ValueOf("text"), pr, MFX_ALIGN_TOP
                cy = pr.Height - 4

            End If

            ' /* title */
            If ppd.ValueOf("title") <> "" Then
                .SetFont FONT_NAME, FONT_SIZE_TITLE, True
                Set pr = new_BRect(xOffset, .Bounds.Bottom - TEXT_MARGIN - 1 - .StringHeight("A"), .Bounds.Right - MARGIN - 1, .Bounds.Bottom - TEXT_MARGIN - 1)
                pr.OffsetBy 0, -cy
                .SetHighColour rgba(255, 255, 255)
                .DrawString ppd.ValueOf("title"), pr, MFX_ALIGN_TOP

            End If
        End If

        ' /* highlight */

'Dim i As Long


'        .DrawScaledImage pb, new_BPoint(-150, -150), , 70, 52

        ' /* now, convert this into a bitmap and redraw but with a constant alpha */

        Set pb = .ConvertToBitmap()
        .Clear

        If NotificationInfo.Scheme = LCase$(SCHEME_2) Then
            ' /* compact */
            .SizeTo .Width, 96
            .DrawScaledImage pb, , new_BPoint(.Width, .Height), 250

        Else
            ' /* normal */
            .DrawScaledImage pb, , , 250

        End If

    End With

End Sub

Private Function uRatingPercent(ByVal Rating As String) As Integer
Dim s() As String

    uRatingPercent = -1

    On Error Resume Next
    
    s = Split(Rating, "/")
    If UBound(s) <> 1 Then _
        Exit Function

Dim a As Single
Dim b As Single

    a = Val(s(0))
    b = Val(s(1))

    If b <= 0 Then _
        Exit Function

    a = a / b
    Debug.Print "rating = " & a

    If a > 1 Then _
        Exit Function

    uRatingPercent = a * 100

End Function


