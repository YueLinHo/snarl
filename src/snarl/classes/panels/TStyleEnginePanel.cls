VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleEnginePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage
Dim mhWndPanel As Long

Dim mSelectedEngine As String

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

    uUpdateList

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    Set pm = New CTempMsg
    pm.Add "item-height", 24&
    Set pc = new_BPrefsControl("listbox", "lb>engines", , , , pm)
    pc.SizeTo 0, 128
    mPage.Add pc

    Set pc = new_BPrefsControl("label", "lbl>desc")
    pc.SizeTo 0, 26
    mPage.Add pc

    mPage.Add new_BPrefsControl("label", "", "Installed in:")
    Set pc = new_BPrefsControl("label", "lbl>path")
    pc.SizeTo 0, 26
    mPage.Add pc

    mPage.Add new_BPrefsControl("fancytoolbar", "ftb>engine", "Show in Explorer||Install...")

    mPage.Add new_BPrefsControl("fancybutton2", "fb2>loadunload", "Unload")
'    mPage.Add new_BPrefsControl("fancybutton2", "fb2>install", "Install...")

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pEngineInfo As T_SNARL_STYLE_ENGINE_INFO
Dim pc As BControl
Dim sz As String

    Select Case Control.GetName()
    Case "lb>engines"
        mSelectedEngine = ""

        With g_StyleRoster
            If .EngineAt(Val(Value), pEngineInfo) Then

                mSelectedEngine = pEngineInfo.Name

                If mPage.Find("lbl>desc", pc) Then
                    If (pEngineInfo.Obj Is Nothing) Then
                        pc.SetText "Not loaded"

                    Else
                        pc.SetText pEngineInfo.Description

                    End If

                End If

                If mPage.Find("lbl>path", pc) Then _
                    pc.SetText pEngineInfo.Path

                If mPage.Find("fb2>loadunload", pc) Then
                    If (pEngineInfo.Obj Is Nothing) Then
                        pc.SetText "Load"

                    Else
                        pc.SetText "Unload"

                    End If

                    pc.SetEnabled ((pEngineInfo.Flags And &H80000000) = 0)

                End If

            Else
            
            End If

        End With


    Case "ftb>engine"

        ' /* retrieve info for the currently selected engine */

        If Not mPage.Find("lb>engines", pc) Then _
            Exit Sub

        If Not g_StyleRoster.EngineAt(Val(pc.GetValue), pEngineInfo) Then _
            Exit Sub

        Select Case Val(Value)
        Case 1
            ShellExecute 0, "open", pEngineInfo.Path, vbNullString, vbNullString, SW_SHOWNORMAL

'        Case 3
'            If g_StyleRoster.StylesForEngine(mSelectedEngine, sz) Then
'                If MsgBox("Uninstalling this style engine will remove the following styles:" & vbCrLf & vbCrLf & sz & vbCrLf & vbCrLf & "Are you sure you want to do this?", _
'                        vbQuestion Or vbYesNo, App.Title) = vbYes Then
'
'                End If
'
'            End If

        Case 3

Dim szEngine As String

            szEngine = InputBox("Style Engine to install?", App.Title)
            If szEngine = "" Then _
                Exit Sub
    
            If g_GetExtension(szEngine) = "" Then _
                szEngine = szEngine & ".styleengine"
    
            If g_StyleRoster.Load(szEngine, sz, True) Then
                snShowMessage szEngine & " installed ok", "", -1, g_MakePath(App.Path) & "etc\icons\style_engine.png"
                uUpdateList
    
                If g_GetUserFolderPath(sz) Then
                    sz = g_MakePath(sz) & "styles\" & szEngine
                    If Not g_Exists(sz) Then _
                        uCreateBlankFile sz
    
                End If
    
            Else
                MsgBox sz, vbCritical Or vbOKOnly, "Error Installing Style Engine"
    
            End If

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim sz As String

    Select Case Control.GetName()
    Case "fb2>done"
        mPanel.Quit

    Case "fb2>loadunload"
        If (g_StyleRoster Is Nothing) Then _
            Exit Sub

        Select Case Control.GetText
        Case "Load"
            If g_StyleRoster.Load(mSelectedEngine, sz, True) Then
                MsgBox "Style Engine '" & mSelectedEngine & "' loaded", vbInformation Or vbOKOnly, App.Title
                uUpdateList

            Else
                MsgBox "Style Engine '" & mSelectedEngine & "' failed to load" & vbCrLf & vbCrLf & sz, vbInformation Or vbOKOnly, App.Title
            
            End If

        Case "Unload"
            
            If g_StyleRoster.StylesForEngine(mSelectedEngine, sz) Then
                If MsgBox("Uninstalling this style engine will remove the following styles:" & vbCrLf & vbCrLf & _
                          sz & vbCrLf & "Are you sure you want to do this?", _
                          vbQuestion Or vbYesNo, App.Title) = vbYes Then

                    If g_StyleRoster.Unload(mSelectedEngine, sz) Then
                        MsgBox "Style Engine '" & mSelectedEngine & "' unloaded", vbInformation Or vbOKOnly, App.Title
                        uUpdateList

                    Else
                        MsgBox "Error unloading Style Engine", vbExclamation Or vbOKOnly, App.Title

                    End If

                End If
            End If

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 0

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
    mPanel.SetWindow 1                  ' // SPECIAL HACK: tells the BPrefsPanel not to auto-display during Go()

'    mPanel.AddPage new_BPrefsPage("Engines", gfxlibSafeLoadImageO(g_MakePath(App.Path) & "etc\icons\style_engine.png"), Me)

    mPanel.AddPage new_BPrefsPage("", , Me)

    mPanel.SetTitle "Installed Style Engines"
    mPanel.SetWidth 280

    mPanel.Go
    g_SetWindowIconToAppResourceIcon mPanel.hWnd

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

Private Function uUpdateList() As String
Dim pEngineInfo As T_SNARL_STYLE_ENGINE_INFO
Dim pm As CTempMsg
Dim pc As BControl
Dim sz As String
Dim i As Long

    If Not (mPage.Find("lb>engines", pc)) Then _
        Exit Function

    If Not (g_StyleRoster Is Nothing) Then
        If g_StyleRoster.CountEngines > 0 Then

            With g_StyleRoster
                For i = 1 To .CountEngines
                    If .EngineAt(i, pEngineInfo) Then
                        With pEngineInfo
                            sz = sz & IIf(.Name = "", "<unknown>", .Name) & " V" & CStr(.Version) & "." & CStr(.Revision) & " (" & .Date & ")|"

                        End With
                    End If
                Next i
            End With

            pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)

            ' /* set the icons */

            Set pm = New CTempMsg

            With g_StyleRoster

                For i = 1 To .CountEngines
                    .EngineAt i, pEngineInfo
                    If (pEngineInfo.Obj Is Nothing) Then
                        sz = "broken.png"

                    Else
                        sz = "style_engine.png"

                    End If

                    prefskit_SetItem pc, i, "image-file", g_MakePath(App.Path) & "etc\icons\" & sz

                Next i

            End With

        Else
            pc.SetText "<None>"

        End If

    Else
        pc.SetText "<Style Roster not running>"

    End If

    pc.SetValue "1"

End Function

Private Sub uCreateBlankFile(ByVal Path As String)

    On Error Resume Next

Dim n As Integer

    n = FreeFile()
    Open Path For Output As #n
'    Print #n, ""
    Close #n

End Sub
