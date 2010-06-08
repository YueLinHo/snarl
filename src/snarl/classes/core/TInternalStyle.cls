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
    '/  Description:    Implements our single internal style (and associated style engine)
    '/
    '/  Created:        11-May-2009
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const ENG_DATE = "8-Mar-2010"

Private Const STYLE_NAME = "Default"
Private Const DEF_STYLE_DATE = "6-Apr-2009"
Private Const DEF_STYLE_COPY = "© 2009 full phat products"

Private Const IP_STYLE_NAME = "IP Forwarder"
Private Const IP_STYLE_DATE = "8-Mar-2010"
Private Const IP_STYLE_COPY = "© 2010 full phat products"

Private Const BORDER_SIZE = 6
Private Const ICON_SIZE = 32

Public IsIPForwarder As Boolean

Dim theView As mfxView

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
        Set pStyle = New TInternalStyle
        pStyle.IsIPForwarder = True
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

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    ' /* TO-DO: check if SNB is installed okay? */

End Function

Private Function IStyleEngine_LastError() As String

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = "<snarl>"

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = 1

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)

    If (Index < 1) Or (Index > 2) Then _
        Exit Sub

        With Style
            .Path = App.Path
            .Schemes = ""
            .SupportEmail = "info@fullphat.net"
            .URL = "http://www.fullphat.net"

        Select Case Index

        Case 1
            .Copyright = DEF_STYLE_COPY
            .Date = DEF_STYLE_DATE
            .Description = "Built-in default style"
            .Flags = S_STYLE_MULTI_INSTANCE
            .IconPath = g_MakePath(App.Path) & "etc\icons\default_style.png"
            .Major = 1
            .Minor = 0
            .Name = STYLE_NAME


        Case 2
            .Copyright = IP_STYLE_COPY
            .Date = IP_STYLE_DATE
            .Description = "Forwards notifications to the computers selected in the 'Network' tab"
            .Flags = S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME
            .IconPath = g_MakePath(App.Path) & "etc\icons\network.png"
            .Major = 0
            .Minor = 1
            .Name = IP_STYLE_NAME

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

    If IsIPForwarder Then
        uIPFStyleUpdate NotificationInfo
    
    Else
        uDefStyleUpdate NotificationInfo

    End If

End Sub

Private Sub uIPFStyleUpdate(ByRef NotificationInfo As notification_info)

    Debug.Print "-----> forwarder <-------"
    If (gRemoteComputers Is Nothing) Then
        Debug.Print "--- error ---"
        Exit Sub

    End If

Static i As Long

    With gRemoteComputers
        Debug.Print "remote computers: " & .CountEntries

        If .CountEntries Then
            For i = 1 To .CountEntries
                Debug.Print .EntryAt(i).Name & " / " & .EntryAt(i).Value

                gForwarders = gForwarders + 1
                ReDim Preserve gForwarder(gForwarders)
                Set gForwarder(gForwarders) = New TForwarder
                gForwarder(gForwarders).Init .EntryAt(i).Name, NotificationInfo, gForwarderUID
                gForwarderUID = gForwarderUID + 4

            Next i
        End If

    End With

End Sub

Private Sub uDefStyleUpdate(ByRef NotificationInfo As notification_info)
Dim cx As Long
Dim cy As Long
Dim pr As BRect
Dim bm As mfxBitmap

    On Error Resume Next

Const RX = 20

    cx = BORDER_SIZE

Dim pbm As MImage

Dim cyTitle As Long
Dim cyText As Long
Dim prFrame As BRect

    With theView

        ' /* get icon */

        load_image NotificationInfo.Icon, pbm
        If melonIsValidImage(pbm) Then _
            cx = cx + ICON_SIZE + BORDER_SIZE

        ' /* calculate height */

        .SizeTo 250, 110
        .TextMode = MFX_TEXT_ANTIALIAS

        Set prFrame = .Bounds.Duplicate
        prFrame.Left = cx

        .SetFont "Tahoma", 8, True
        .MeasureString NotificationInfo.Title, prFrame, pr
        cyTitle = pr.Height

        .SetFont "Tahoma", 8, False
        .MeasureString NotificationInfo.Text, prFrame, pr
        cyText = pr.Height

        cy = (BORDER_SIZE * 2) + cyTitle + cyText
        cy = MAX((BORDER_SIZE * 2) + ICON_SIZE, cy)         ' // must be at least (ICON_SIZE + (2 x BORDER_SIZE)) high...
        .SizeTo 250, cy

        ' /* draw background */

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(236, 59, 0)
            .SetLowColour rgba(206, 52, 0)
            .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT
            .SetHighColour rgba(255, 0, 0)
            .StrokeRoundRect .Bounds, RX, RX

        Else
            .SetHighColour rgba(235, 235, 235)
            .SetLowColour rgba(205, 205, 205)
            .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT
            .SetHighColour rgba(176, 176, 176)
            .StrokeRoundRect .Bounds, RX, RX

        End If



        ' /* icon */

        If Not (pbm Is Nothing) Then _
            .DrawScaledImage pbm, new_BPoint(BORDER_SIZE, BORDER_SIZE), new_BPoint(ICON_SIZE, ICON_SIZE)

        Set pr = .Bounds.Duplicate
        pr.InsetBy BORDER_SIZE, BORDER_SIZE
        pr.Left = cx

'        .SetHighColour rgba(255, 0, 0)
'        .StrokeRect pr

        If (NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) Then
            .SetHighColour rgba(255, 255, 255, 210)

        Else
            .SetHighColour rgba(36, 36, 36, 210)

        End If

        ' /* title */

        .SetFont "Tahoma", 8, True
        .DrawString NotificationInfo.Title, pr

        ' /* text */

        .SetFont "Tahoma", 8
        pr.OffsetBy 0, cyTitle
        .DrawString NotificationInfo.Text, pr

    End With

End Sub


