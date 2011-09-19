VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAppsPopUpWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TAppsPopUpWindow.cls
    '/
    '/  Description:    Implements the floating daemon-picker window
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Event Closed()
Public Event Selected(ByVal Signature As String)

Const CLASS_NAME = "w>SnarlDaemons"
Dim mhWnd As Long

Dim mView As mfxView
Dim mApps As BTagList

Const CXICON = 32
Const CXMARGIN = 4
Const CYITEM = 36
Dim mSelected As String

Implements BWndProcSink

Public Sub Show()
Dim dw As Long

    dw = GetMessagePos()

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , , WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

    Set mView = New mfxView
    With mView
        .SetFont "Arial", 9, True
        .TextMode = MFX_TEXT_ANTIALIAS

    End With

    Set mApps = new_BTagList()

Dim cx As Long
Dim i As Long

    ' /* figure out required width */

    cx = 50

    With g_AppRoster
        If .CountApps Then
            For i = 1 To .CountApps
                If .AppAt(i).IncludeInMenu Then
                    mApps.Add .AppAt(i)
                    cx = MAX(cx, mView.StringWidth(.AppAt(i).Name))

                End If
            Next i
        End If

    End With

    If mApps.CountItems = 0 Then _
        Exit Sub

    ' /* set up the view */

Dim pr As BRect
Dim pa As TApp

    mView.SizeTo cx + CXICON + 4 + (CXMARGIN * 2) + 64, (mApps.CountItems * CYITEM) + ((mApps.CountItems - 1) * 2) + (CXMARGIN * 2)

    Set pr = mView.Bounds
    pr.InsetBy CXMARGIN, CXMARGIN
    pr.Bottom = pr.Top + (CYITEM - 1)

    With mApps
        .Rewind
        Do While .GetNextTag(pa) = B_OK
            Set pa.PopupFrame = pr.Duplicate
            pr.OffsetBy 0, CYITEM + 2

        Loop

    End With

    uRedraw

Dim pt As POINTAPI

    pt.x = GET_X_LPARAM(dw) - Fix(mView.Width / 2)
    pt.y = GET_Y_LPARAM(dw) - Fix(mView.Height / 2)

    If pt.x + mView.Width + 16 > g_ScreenWidth Then _
        pt.x = g_ScreenWidth() - (mView.Width + 16)

    If pt.y + mView.Height + 16 > g_ScreenHeight() Then _
        pt.y = g_ScreenHeight() - (mView.Height + 16)

    apply_view_to_window mView, mhWnd, , pt.x, pt.y
    g_ShowWindow mhWnd, True, True
    SetCapture mhWnd

End Sub

Public Sub Hide()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static szPressed As String
Dim sz As String

    Select Case uMsg
    Case WM_CAPTURECHANGED
        PostMessage hWnd, &H401, 0, ByVal 0&

    Case WM_MOUSEMOVE
        sz = uHitTest(new_BPointFromInt32(lParam))
        If sz <> mSelected Then
            mSelected = sz
            uRedraw
            apply_view_to_window mView, hWnd

        End If

    Case WM_LBUTTONDOWN
        szPressed = uHitTest(new_BPointFromInt32(lParam))

    Case WM_LBUTTONUP
        If (uHitTest(new_BPointFromInt32(lParam)) = szPressed) And (szPressed <> "") Then _
            RaiseEvent Selected(szPressed)

        PostMessage hWnd, &H401, 0, ByVal 0&
        szPressed = ""

    Case &H401
        RaiseEvent Closed
        Me.Hide

    End Select

End Function

Private Sub uRedraw()
Dim pa As TApp
Const RX = 14
        
    With mView
        .Clear
        .EnableSmoothing True
        .SetHighColour rgba(24, 24, 24, 190)
        .FillRoundRect .Bounds, RX, RX
        .SetHighColour rgba(255, 255, 255, 190)
        .StrokeRoundRect .Bounds.InsetByCopy(1, 1), RX, RX

'        .SetHighColour rgba(244, 244, 244, 190)
'        .FillRoundRect .Bounds, RX
'        .SetHighColour rgba(0, 0, 0, 96)
'        .StrokeRoundRect .Bounds, RX, RX
        
        .SetHighColour rgba(0, 0, 0, 64)
        .SetLowColour rgba(0, 0, 0, 0)
        .FillRoundRect .Bounds, RX, RX, MFX_VERT_GRADIENT


    End With

    With mApps
        .Rewind
        Do While .GetNextTag(pa) = B_OK
            If pa.Signature = mSelected Then
''                mView.SetHighColour rgba(255, 255, 255, 60)
''                mView.EnableSmoothing False
'                mView.SetHighColour rgba(0, 192, 255, 190)
'                mView.SetLowColour rgba(0, 192, 255, 90)
'                mView.SetHighColour rgba(128, 128, 128, 64)

                mView.SetHighColour rgba(255, 255, 255, 96)
                mView.FillRoundRect pa.PopupFrame.InsetByCopy(-1, -1), RX, RX ', MFX_VERT_GRADIENT
'                mView.SetHighColour rgba(0, 0, 0, 90)
'                mView.SetLowColour rgba(255, 255, 255, 90)
'                mView.StrokeFancyRect pa.PopupFrame

''                mView.SetHighColour rgba(, 192, 255, 90)
''                mView.SetLowColour rgba(, 192, 255, 0)
''                mView.EnableSmoothing True
                mView.SetHighColour rgba(0, 0, 0, 90)
                mView.StrokeRoundRect pa.PopupFrame.InsetByCopy(-1, -1), RX, RX

            End If

            mView.SetHighColour rgba(255, 255, 255)
'            mView.SetHighColour rgba(0, 0, 0, 250)
            mView.DrawString pa.Name, pa.PopupFrame.OffsetByCopy(CXICON + 4, 0), MFX_ALIGN_LEFT Or MFX_ALIGN_V_CENTER
            mView.DrawScaledImage pa.CachedIcon, pa.PopupFrame.TopLeft.OffsetByCopy(0, Fix((CYITEM - CXICON) / 2)), new_BPoint(CXICON, CXICON)

        Loop

    End With

End Sub

Private Function uHitTest(ByRef Point As BPoint) As String
Dim pa As TApp

    With mApps
        .Rewind
        Do While .GetNextTag(pa) = B_OK
            If pa.PopupFrame.Contains(Point) Then
                uHitTest = pa.Signature
                Exit Function

            End If
        Loop
    End With

End Function
