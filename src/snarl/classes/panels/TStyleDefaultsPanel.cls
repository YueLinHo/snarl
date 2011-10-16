VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleDefaultsPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mhWndOwner As Long

Public Event Closed()

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    If Control.GetName() = "toolbar" Then
        Select Case Val(Value)
        Case 1
            g_DoSchemePreview2 "Default", "", False, -1
            g_DoSchemePreview2 "Default", "", True, -1

        Case 3
            Debug.Print "(*)"

        End Select
    End If

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    EnableWindow mhWndOwner, -1
    g_ShowWindow mhWndOwner, True, True
    RaiseEvent Closed

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub Go(ByVal hWndPanel As Long)
Dim rcThis As RECT
Dim rcOwner As RECT
Dim dw As Long
Dim sz As String
Dim sza As String
Dim pp As BPrefsPage
Dim pm As CTempMsg
Dim pc As BControl

    mhWndOwner = hWndPanel

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "Style Defaults"
        .SetWidth 430

        ' /* page */

        Set pp = new_BPrefsPage("", , Me, 430)
        With pp
            .SetMargin 0
            Set pm = New CTempMsg
            pm.Add "height", 260
            Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
            BTabStrip_AddPage pc, "General", new_BPrefsPage("defs-general", , New TNetSubPage)
            BTabStrip_AddPage pc, "Title", new_BPrefsPage("defs-title", , New TNetSubPage)
            BTabStrip_AddPage pc, "Text", new_BPrefsPage("defs-text", , New TNetSubPage)
            .Add pc

            .Add new_BPrefsControl("fancytoolbar", "toolbar", "Test Settings||Restore Defaults")

        End With

        .AddPage pp

        ' /* page */

        .Go
        g_SetWindowIconToAppResourceIcon .hWnd

        dw = GetWindowLong(.hWnd, GWL_STYLE)
        dw = dw And (Not WS_MINIMIZEBOX)
        SetWindowLong .hWnd, GWL_STYLE, dw

        SetWindowLong .hWnd, GWL_HWNDPARENT, hWndPanel
        EnableWindow hWndPanel, 0

        GetWindowRect hWndPanel, rcOwner
        GetWindowRect .hWnd, rcThis
        g_RectNormalise rcThis

        g_MoveWindow .hWnd, _
                     rcOwner.Left + Fix(((rcOwner.Right - rcOwner.Left) - rcThis.Right) / 2), _
                     rcOwner.Top + Fix(((rcOwner.Bottom - rcOwner.Top) - rcThis.Bottom) / 2)

        g_ShowWindow .hWnd, True, True

    End With

End Sub


