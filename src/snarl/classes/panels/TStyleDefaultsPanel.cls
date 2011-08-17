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
Dim sz As String

    sz = g_SafeRightStr(Control.GetName, Len(Control.GetName) - 6)
    gStyleDefaults.Update sz, Value
    gStyleDefaults.Save

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    If Control.GetName() = "test" Then _
        g_DoSchemePreview2 "Default", "", False, -1

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

    mhWndOwner = hWndPanel

'        .AddItem .CreateItem("frwd", "Forward to remote computer...")
'        .AddItem .CreateItem("subs", "Subscribe to remote SNP server...")
'        .AddItem .CreateItem("subg", "Subscribe to remote GNTP server...")

    Set mPanel = New BPrefsPanel
    With mPanel
        .SetHandler Me
        .SetTitle "Style Defaults"
        .SetWidth 480

        ' /* page */

        Set pp = new_BPrefsPage("", , Me, 480)

        With pp
            .SetMargin 90

'            .Add new_BPrefsControl("label", "", "These settings apply to the default style, however other styles may also (and are encouraged to) support these settings.")

            .Add new_BPrefsControl("banner", "", "General")

            .Add new_BPrefsControl("fancyedit", "style:width", gStyleDefaults.ValueOf("width"), "Width:")
            .Add new_BPrefsControl("fancycolourpicker", "style:background-colour", , "Background:", gStyleDefaults.ValueOf("background-colour"))

            .Add new_BPrefsControl("banner", "", "Notification Title")
            .Add new_BPrefsControl("fancyfontpicker", "style:title-font", gStyleDefaults.ValueOf("title-font"), "Font:", gStyleDefaults.ValueOf("title-font"))
            .Add new_BPrefsControl("fancycolourpicker", "style:title-colour", , "Colour:", gStyleDefaults.ValueOf("title-colour"))
'            .Add new_BPrefsControl("fancycycle", "style:title-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("title-weight"))
            Set pm = New CTempMsg
            With pm
                .Add "min", 0&
                .Add "max", 100&
                .Add "freq", 10&
                .Add "suffix", "%"

            End With
            .Add new_BPrefsControl("fancyslider", "style:title-opacity", "", "Opacity:", gStyleDefaults.ValueOf("title-opacity"), pm)

            .Add new_BPrefsControl("banner", "", "Notification Text")
            .Add new_BPrefsControl("fancyfontpicker", "style:text-font", gStyleDefaults.ValueOf("text-font"), "Font:", gStyleDefaults.ValueOf("text-font"))
            .Add new_BPrefsControl("fancycolourpicker", "style:text-colour", , "Colour:", gStyleDefaults.ValueOf("text-colour"))
'            .Add new_BPrefsControl("fancycycle", "style:text-weight", "Normal|Bold", "Weight:", gStyleDefaults.ValueOf("text-weight"))
            Set pm = New CTempMsg
            With pm
                .Add "min", 0&
                .Add "max", 100&
                .Add "freq", 10&
                .Add "suffix", "%"

            End With
            .Add new_BPrefsControl("fancyslider", "style:text-opacity", "", "Opacity:", gStyleDefaults.ValueOf("text-opacity"), pm)

            .Add new_BPrefsSeparator
            .Add new_BPrefsControl("fancybutton2", "test", "Test")

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


