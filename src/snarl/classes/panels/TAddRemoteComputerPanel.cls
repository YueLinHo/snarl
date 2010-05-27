VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAddRemoteComputerPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage
Dim mhWndPanel As Long

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl

    mPage.Add new_BPrefsControl("ipaddr", "ia>addr", , "IP Address:")
    
    Set pc = new_BPrefsControl("fancybutton2", "fb2>done", "Ok")
    pc.SizeTo 80, 0             ' // a zero height here means "let control decide"

    mPage.Add pc

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim pc As BControl

    Select Case Control.GetName()
    Case "fb2>done"
        If mPage.Find("ia>addr", pc) Then _
            g_AddRemoteComputer pc.GetText()

        mPanel.Quit

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 100
    mPage.SetWidth 220

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

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

    EnableWindow mhWndPanel, -1
    g_ShowWindow mhWndPanel, True, True

End Sub

Private Sub KPrefsPanel_Ready()

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)

End Sub

Public Sub Go(ByVal hWndPanel As Long)

    mhWndPanel = hWndPanel
    Set mPanel = New BPrefsPanel
    mPanel.SetHandler Me

    mPanel.AddPage new_BPrefsPage("", , Me)
    mPanel.SetTitle "Add Remote Computer"
    mPanel.SetWidth 240

    mPanel.Go
    g_SetWindowIconToAppResourceIcon mPanel.hWnd

Dim dw As Long

    dw = GetWindowLong(mPanel.hWnd, GWL_STYLE)
    dw = dw And (Not WS_MINIMIZEBOX)
    SetWindowLong mPanel.hWnd, GWL_STYLE, dw

Dim rcConfig As RECT
Dim rc As RECT

    SetWindowLong mPanel.hWnd, GWL_HWNDPARENT, hWndPanel
    EnableWindow hWndPanel, 0

    GetWindowRect hWndPanel, rc
    GetWindowRect mPanel.hWnd, rcConfig
    g_RectNormalise rcConfig

    g_MoveWindow mPanel.hWnd, _
                 rc.Left + Fix(((rc.Right - rc.Left) - rcConfig.Right) / 2), _
                 rc.Top + Fix(((rc.Bottom - rc.Top) - rcConfig.Bottom) / 2)

    g_ShowWindow mPanel.hWnd, True, True

End Sub



