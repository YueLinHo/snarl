VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TInternalStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TInternalStyle.cls
    '/
    '/  Description:    Implements our internal styles (and associated style engine)
    '/
    '/  Created:        11-May-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const ENG_DATE = "18-Oct-2011"

Private Const STYLE_NAME = "Default"
Private Const DEF_STYLE_DATE = "25-Jul-2011"
Private Const DEF_STYLE_COPY = "© 2011 full phat products"

Private Const IP_STYLE_NAME = "IP Forwarder"
Private Const IP_STYLE_DATE = "12-Dec-2010"
Private Const IP_STYLE_COPY = "© 2010 full phat products"

Private Const RUN_STYLE_NAME = "Run File"
Private Const RUN_STYLE_DATE = "13-Jan-2011"
Private Const RUN_STYLE_COPY = "© 2011 full phat products"

Private Const IMAGE_STYLE_NAME = "Image"
Private Const IMAGE_STYLE_DATE = "21-Jan-2011"
Private Const IMAGE_STYLE_COPY = "© 2011 full phat products"

Private Const BORDER_SIZE = 6
Private Const ICON_SIZE = 32

Public IsIPForwarder As Boolean
Public IsFileRunner As Boolean
Public IsImage As Boolean

Dim theView As mfxView
Dim mPanel As TRunFilePanel

Implements IStyleInstance
Implements IStyleEngine

Private Sub Class_Initialize()

    Set theView = New mfxView

End Sub

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = 2

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim pStyle As TInternalStyle

    If StyleName = IP_STYLE_NAME Then
'        Set pStyle = New TInternalStyle
'        pStyle.IsIPForwarder = True
'        Set IStyleEngine_CreateInstance = pStyle

    ElseIf StyleName = RUN_STYLE_NAME Then
        Set pStyle = New TInternalStyle
        pStyle.IsFileRunner = True
        Set IStyleEngine_CreateInstance = pStyle

    ElseIf StyleName = IMAGE_STYLE_NAME Then
        Set pStyle = New TInternalStyle
        pStyle.IsImage = True
        Set IStyleEngine_CreateInstance = pStyle

    Else
        Set IStyleEngine_CreateInstance = New TInternalStyle

    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = ENG_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Provides the built-in default styles"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long

    If StyleName = "Run File" Then
        Set mPanel = New TRunFilePanel
        IStyleEngine_GetConfigWindow = mPanel.Go(True)

    ElseIf StyleName = "Image" Then
        Set mPanel = New TRunFilePanel
        IStyleEngine_GetConfigWindow = mPanel.Go(False)

    End If

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

'    g_RunFileLoadSchemes

End Function

Private Function IStyleEngine_LastError() As String
End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = "[Internal Styles]"

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Minor

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)

    If (Index < 1) Or (Index > IStyleEngine_CountStyles()) Then _
        Exit Sub

        With Style
            .Path = App.Path
            .Schemes = ""
            .SupportEmail = "info@fullphat.net"
            .URL = "http://www.getsnarl.info"

        Select Case Index

        Case 1
            .Copyright = DEF_STYLE_COPY
            .Date = DEF_STYLE_DATE
            .Description = "Built-in style which uses the Default Style settings (see the [Display] page)."
            .Flags = S_STYLE_MULTI_INSTANCE
            .IconPath = g_MakePath(App.Path) & "etc\icons\default_style.png"
            .Major = 1
            .Minor = 1
            .Name = STYLE_NAME


'        Case 2
'            .Copyright = IP_STYLE_COPY
'            .Date = IP_STYLE_DATE
'            .Description = "Forwards notifications to the computers selected in the 'Network' tab"
'            .Flags = S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME
'            .IconPath = g_MakePath(App.Path) & "etc\icons\network.png"
'            .Major = 0
'            .Minor = 8
'            .Name = IP_STYLE_NAME

        Case 2
            .Copyright = RUN_STYLE_COPY
            .Date = RUN_STYLE_DATE
            .Description = "Executes a file, passing the notification content as arguments"
            .Flags = S_STYLE_IS_WINDOWLESS Or S_STYLE_IS_CONFIGURABLE Or S_STYLE_V42_CONTENT
            .IconPath = g_MakePath(App.Path) & "etc\icons\runfile.png"
            .Major = 2
            .Minor = 0
            .Name = RUN_STYLE_NAME
            .Schemes = uGetFileSchemes("runfile")

        Case 3
            .Copyright = IMAGE_STYLE_COPY
            .Date = IMAGE_STYLE_DATE
            .Description = "Runs a user-defined file"
            .Flags = S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_CONFIGURABLE
            .IconPath = g_MakePath(App.Path) & "etc\icons\image_style.png"
            .Major = 0
            .Minor = 1
            .Name = IMAGE_STYLE_NAME
            .Schemes = uGetFileSchemes("image")

        End Select

    End With

End Sub

Private Sub IStyleEngine_TidyUp()

End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    If Not IsIPForwarder Then _
        Set IStyleInstance_GetContent = theView.ConvertToBitmap()

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim prf As TRunFileScheme

    If IsIPForwarder Then
        uIPFStyleUpdate NotificationInfo

    ElseIf IsFileRunner Then
        ' /* notificationinfo->scheme is the name */

        With gRunFiles
            .Rewind
            Do While .GetNextTag(prf) = B_OK
                If LCase$(prf.Name) = NotificationInfo.Scheme Then
                    uDoFileRunner NotificationInfo, prf
                    Exit Sub

                End If

            Loop

            g_Debug "TInternalStyle.UpdateContent(): scheme '" & NotificationInfo.Scheme & "' not found", LEMON_LEVEL_CRITICAL

        End With


    ElseIf IsImage Then
        ' /* Not enabled in R2.4 - for further development */
        uImageStyleUpdate NotificationInfo

    Else
        uDefStyleUpdate NotificationInfo

    End If

End Sub

Private Sub uDoFileRunner(ByRef NotificationInfo As notification_info, ByRef Scheme As TRunFileScheme)
Dim ppd As BPackedData
Dim s As String
Dim t As String
Dim u As String

    ' /* title is sender name, text is original info */

    Select Case Scheme.Version

    Case 1
        g_Debug "TInternalStyle.uDoFileRunner(): sending V1 content to '" & Scheme.Target & "'..."

        Set ppd = New BPackedData
        ppd.SetTo NotificationInfo.Text

        ' /* format is <file_to_run> <title> <text> <sender> [data-] */

        With ppd
            .Rewind

            ' /* loop through and for any item beginnning with either "data-" or
            '    "value-" add it to the args */

            Do While .GetNextItem(s, t)
                If g_SafeLeftStr(s, 5) = "data-" Then
                    u = u & " " & g_Quote(s & "=" & t)

                ElseIf g_SafeLeftStr(s, 6) = "value-" Then
                    u = u & " " & g_Quote(s & "=" & t)

                End If

            Loop

        End With

        ' /* launch the file using the "open" verb.  we must also escape any CRLF's back to "\n"'s
        '    as these can play havoc with MS-DOS batch files */

        ShellExecute 0, "open", Scheme.Target, g_Quote(Replace$(ppd.ValueOf("title"), vbCrLf, "\n")) & " " & _
                                               g_Quote(Replace$(ppd.ValueOf("text"), vbCrLf, "\n")) & " " & _
                                               g_Quote(NotificationInfo.Title) & _
                                               u, vbNullString, SW_SHOW

        ' /* FOR FUTURE DEVELOPMENT: setting notification_info->Flags to -1 tells
        '    Snarl that the original notification should be displayed on-screen */

'        NotificationInfo.Flags = -1


    Case 2
        ' /* replace anything surrounded by %..% with the equivalent from the notification content */
        g_Debug "TInternalStyle.uDoFileRunner(): template is '" & Scheme.Template & "'"
        Set ppd = New BPackedData
        ppd.SetTo NotificationInfo.Text
        s = uReplaceMarkers(Scheme.Template, ppd)
        g_Debug "TInternalStyle.uDoFileRunner(): fixed-up template is '" & s & "'..."
        ShellExecute 0, "open", Scheme.Target, s, vbNullString, SW_SHOW


    Case 3
        ' /* &/= format */
        s = g_URLEncode(NotificationInfo.Text)
        s = Replace$(s, "#?", "&")
        s = Replace$(s, "::", "=")
        s = s & "&app=" & NotificationInfo.Title
        g_Debug "TInternalStyle.uDoFileRunner(): sending V3 content to '" & Scheme.Target & "'..."
        ShellExecute 0, "open", Scheme.Target, g_Quote(Scheme.Template & s), vbNullString, SW_SHOW

    End Select

End Sub

Private Sub uIPFStyleUpdate(ByRef NotificationInfo As notification_info)

'    g_ForwardNotification NotificationInfo

End Sub

Private Sub uDefStyleUpdate(ByRef NotificationInfo As notification_info)
Dim cx As Long
Dim cy As Long
Dim pr As BRect
Dim bm As mfxBitmap

    On Error Resume Next

Const RX = 20

Dim pbm As MImage

Dim cyTitle As Long
Dim cyText As Long
Dim prFrame As BRect

Dim ppdFont As BPackedData
Dim lCol As Long
Dim lWidth As Long

    Set ppdFont = New BPackedData
    cx = BORDER_SIZE

    With theView

        ' /* get icon */

        load_image NotificationInfo.Icon, pbm
        If melonIsValidImage(pbm) Then _
            cx = cx + ICON_SIZE + BORDER_SIZE

        ' /* calculate height */

        lWidth = g_SafeLong(gStyleDefaults.ValueOf("width"))
        If lWidth < 80 Then _
            lWidth = 80

        .SizeTo lWidth, 110
        .TextMode = MFX_TEXT_ANTIALIAS

        Set prFrame = .Bounds.Duplicate
        With prFrame
            .InsetBy BORDER_SIZE, BORDER_SIZE
            .Left = cx
            .Bottom = 4096

        End With

        If NotificationInfo.Title <> "" Then
            ppdFont.SetTo gStyleDefaults.ValueOf("title-font")
            .SetFont ppdFont.ValueOf("name"), CLng(ppdFont.ValueOf("size")), (ppdFont.ValueOf("bold") = "1")
            .MeasureString NotificationInfo.Title, prFrame, pr
            cyTitle = pr.Height

        End If

        If NotificationInfo.Text <> "" Then
            ppdFont.SetTo gStyleDefaults.ValueOf("text-font")
            .SetFont ppdFont.ValueOf("name"), CLng(ppdFont.ValueOf("size")), (ppdFont.ValueOf("bold") = "1")
            .MeasureString NotificationInfo.Text, prFrame, pr
            cyText = pr.Height

        End If

        cy = (BORDER_SIZE * 2) + cyTitle + cyText
        If (NotificationInfo.Title <> "") And (NotificationInfo.Text <> "") Then _
            cy = cy + 4

        cy = MAX((BORDER_SIZE * 2) + ICON_SIZE, cy)         ' // must be at least (ICON_SIZE + (2 x BORDER_SIZE)) high...
        .SizeTo lWidth, cy
        .EnableSmoothing True

        ' /* draw background */

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour gStyleDefaults.ValueOf("background-colour-priority")

        Else

            .SetHighColour gStyleDefaults.ValueOf("background-colour")

        End If

        .FillRoundRect .Bounds, RX, RX

         .SetHighColour rgba(0, 0, 0, 0)
         .SetLowColour rgba(0, 0, 0, 63)
'            .SetLowColour rgba(206, 52, 0)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT

        .SetHighColour rgba(0, 0, 0, 146)
        .StrokeRoundRect .Bounds, RX, RX

        ' /* icon */

        If Not (pbm Is Nothing) Then _
            .DrawScaledImage pbm, new_BPoint(BORDER_SIZE, BORDER_SIZE), new_BPoint(ICON_SIZE, ICON_SIZE)

        Set pr = .Bounds.Duplicate
        pr.InsetBy BORDER_SIZE, BORDER_SIZE
        pr.Left = cx

'        .SetHighColour rgba(255, 0, 0)
'        .StrokeRect pr

        ' /* title */

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            lCol = CLng(gStyleDefaults.ValueOf("title-colour-priority"))

        Else
            lCol = CLng(gStyleDefaults.ValueOf("title-colour"))

        End If

        .SetHighColour g_SetRGB32Alpha(lCol, g_OpacityFromPercent(gStyleDefaults.ValueOf("title-opacity")))

        ppdFont.SetTo gStyleDefaults.ValueOf("title-font")
        .SetFont ppdFont.ValueOf("name"), CLng(ppdFont.ValueOf("size")), (ppdFont.ValueOf("bold") = "1")
        .DrawString NotificationInfo.Title, pr

        pr.OffsetBy 0, cyTitle

        ' /* etched line */

        If (NotificationInfo.Title <> "") And (NotificationInfo.Text <> "") Then
            .SetHighColour rgba(127, 127, 127, 80)
            .StrokeLine new_BRect(pr.Left, pr.Top + 1, pr.Right, pr.Top + 1)
            .SetHighColour rgba(255, 255, 255, 80)
            .StrokeLine new_BRect(pr.Left, pr.Top + 2, pr.Right, pr.Top + 2)
            pr.OffsetBy 0, 4

        End If


        ' /* text */

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            lCol = CLng(gStyleDefaults.ValueOf("text-colour-priority"))

        Else
            lCol = CLng(gStyleDefaults.ValueOf("text-colour"))

        End If

        .SetHighColour g_SetRGB32Alpha(lCol, g_OpacityFromPercent(gStyleDefaults.ValueOf("text-opacity")))

        ppdFont.SetTo gStyleDefaults.ValueOf("text-font")
        .SetFont ppdFont.ValueOf("name"), CLng(ppdFont.ValueOf("size")), (ppdFont.ValueOf("bold") = "1")
        .DrawString NotificationInfo.Text, pr

    End With

End Sub

Private Function uGetFileSchemes(ByVal Name As String) As String
Dim pt As BTagItem

    g_RunFileLoadSchemes

    With gRunFiles
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            uGetFileSchemes = uGetFileSchemes & pt.Name & "|"

        Loop

    End With

    uGetFileSchemes = g_SafeLeftStr(uGetFileSchemes, Len(uGetFileSchemes) - 1)

End Function

Private Sub uImageStyleUpdate(ByRef NotificationInfo As notification_info)
Dim pb As mfxBitmap

    Set pb = load_image_obj(NotificationInfo.Scheme)
    If Not is_valid_image(pb) Then _
        Exit Sub

    With theView
        .SizeTo pb.Width, pb.Height
        .EnableSmoothing False
        .SetHighColour rgba(0, 0, 0, 1)
        .FillRect .Bounds
        .DrawScaledImage pb

        Set pb = load_image_obj(NotificationInfo.Icon)
        If is_valid_image(pb) Then _
            .DrawScaledImage pb, new_BPoint(0, 0), new_BPoint(48, 48)

    End With

End Sub

Private Function uReplaceMarkers(ByVal str As String, ByRef Content As BPackedData) As String
Dim bWantEndMarker As Boolean
Dim szMarker As String
Dim i As Long
Dim j As Long

    If str = "" Then _
        Exit Function

    j = 1
    i = InStr(str, "%")

    Do While i
        If bWantEndMarker Then
            ' /* found end marker... */
            szMarker = mId$(str, j, i - j)
            uReplaceMarkers = uReplaceMarkers & Content.ValueOf(szMarker)
            i = i + 1
            j = i

        Else
            ' /* found start marker... */
            uReplaceMarkers = uReplaceMarkers & mId$(str, j, i - j)
            i = i + 1
            j = i

        End If
        
        bWantEndMarker = Not bWantEndMarker
        i = InStr(i, str, "%")

    Loop

    If bWantEndMarker Then _
        Debug.Print "syntax error"
    
    If Len(str) >= j Then _
        uReplaceMarkers = uReplaceMarkers & Right$(str, Len(str) - j + 1)

End Function

