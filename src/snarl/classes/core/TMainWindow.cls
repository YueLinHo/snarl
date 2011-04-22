VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /* local Win32 api defines */

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

    ' /* these are only used by the deprecated SNARL_GET_VERSION, we still maintain them however */

Private Const APP_VER = 2
Private Const APP_SUB_VER = 4

    ' /* used for decoding incoming standard commands */

Private Type SNARLSTRUCTI
    Cmd As Long       ' // what to do...
    Id As Long                  ' // snarl message id (returned by snShowMessage())
    Timeout As Long             ' // timeout in seconds (0=sticky)
    LngData2 As Long            ' // reserved
    Title(0 To 1023) As Byte
    Text(0 To 1023) As Byte
    Icon(0 To 1023) As Byte

End Type

    ' /* used for decoding incoming extended commands */

Private Type SNARLSTRUCTEXI
    Cmd As Long       ' // what to do...
    Id As Long                  ' // snarl message id (returned by snShowMessage())
    Timeout As Long             ' // timeout in seconds (0=sticky)
    LngData2 As Long            ' // reserved
    Title(0 To 1023) As Byte
    Text(0 To 1023) As Byte
    Icon(0 To 1023) As Byte
    Class(0 To 1023) As Byte
    Extra(0 To 1023) As Byte
    Extra2(0 To 1023) As Byte
    Reserved1 As Long
    Reserved2 As Long

End Type

Private Type SNARLREQI
    Command As S_SNARL_COMMANDS
    Token As Long
    PacketData(0 To 4095) As Byte

End Type

Dim mhWnd As Long
Dim mhWndPath As Long               ' // a static control which contains our app path...

Implements MWndProcSink

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
'Static pSOS As SOSSTRUCT
'
'    ' /* check for incoming Lemon console input */
'
'    If lemonTranslate(uMsg, lParam, pSOS) Then
'        lemon_handle pSOS
'        ReturnValue = -1
'        MWndProcSink_WndProc = True
'        Exit Function
'
'    End If

    Select Case uMsg

'    Case lemonGlobalMsg()
'        If wParam = 1 Then
'            lemonRegister App.Title, hWnd
'            g_Debug "** Snarl " & App.Comments & " " & CStr(App.Major) & "." & CStr(App.Revision) & " **"
'            g_Debug "** " & App.LegalCopyright
'            g_Debug ""
'
'        End If
'
'    Case WM_SOS
'        ' /* required to confirm we support a Lemon console */
'        ReturnValue = WM_SOS
'        MWndProcSink_WndProc = True


    Case WM_SNARL_INIT
        mhWnd = hWnd

        ' /* create the hidden static control which contains our path */

        mhWndPath = CreateWindowEx(0, "Static", vbNullString, WS_CHILD, 0, 0, 1, 1, hWnd, 0, App.hInstance, ByVal 0&)
        If mhWndPath <> 0 Then _
            SetWindowText mhWndPath, g_MakePath(App.Path)


    Case WM_SNARL_QUIT
        DestroyWindow mhWndPath


    Case WM_COPYDATA
        ' /* Snarl commands are routed here */
        ReturnValue = uCopyData(lParam, wParam)
        MWndProcSink_WndProc = True


    Case WM_MANAGE_SNARL
        ' /* Snarl control messages */
        uManageSnarl wParam, lParam
        MWndProcSink_WndProc = True


    Case WM_SNARLTEST
        ' /* simple test message - aims to prove Snarl is responding */

        Select Case wParam
        Case 0
            szShowMessage "Test Message", _
                          "This is the WM_SNARLTEST test message", _
                          20, _
                          g_MakePath(App.Path) & "styles\default\default_icon.png"

            ReturnValue = -1
            MWndProcSink_WndProc = True

        Case 1
            szShowMessage "Test Message", _
                          "This is a test message (lParam was 0x" & g_HexStr(lParam) & ")", _
                          20, _
                          g_MakePath(App.Path) & "styles\default\default_icon.png"

            ReturnValue = -1
            MWndProcSink_WndProc = True

        Case 64
            szShowMessage g_HexStr((Rnd * 65535) Or &HC0000000), _
                          g_HexStr(Rnd * 16777215) & "-" & g_HexStr(Rnd * 16777215), _
                          20, _
                          g_MakePath(App.Path) & "styles\default\default_icon.png"

            ReturnValue = -1
            MWndProcSink_WndProc = True

        Case Else
            ReturnValue = M_NOT_IMPLEMENTED
            MWndProcSink_WndProc = True

        End Select


'    Case WM_SYSCOMMAND
'        If (wParam = SC_SCREENSAVE) And (g_ConfigGet("auto_sticky_on_screensaver") = 1) Then _


    Case MSG_SHOW_PREFS
        Select Case wParam
        Case 1
            ' /* install specified style engine or extension */
'            uInstall uAtomName(lParam)

        Case 2
            ' /* configure specified style engine or extension */
            uConfigure uAtomName(lParam)


        Case Else
            frmAbout.NewDoPrefs

        End Select


    Case MSG_QUIT, WM_CLOSE
        g_Debug "{" & IIf(uMsg = WM_CLOSE, "WM_CLOSE", "MSG_QUIT") & "}"
        PostQuitMessage 0


    Case WM_QUERYENDSESSION
        g_Debug "{WM_QUERYENDSESSION: " & g_HexStr(wParam) & " " & g_HexStr(lParam) & "}"
'        sn41EZNotify gSnarlToken, "", "", "The system is shutting down"


    Case WM_ENDSESSION
        g_Debug "{WM_ENDSESSION: " & g_HexStr(wParam) & " " & g_HexStr(lParam) & "}"
        If wParam <> 0 Then _
            PostQuitMessage 0

    End Select

End Function

Private Function uManageSnarl(ByVal wParam As Long, ByVal lParam As Long) As Long

    ' /* manage Snarl from external applications - undocumented for now */

    g_Debug "uManageSnarl(" & CStr(wParam) & ", " & CStr(lParam) & ")", LEMON_LEVEL_PROC

    Select Case wParam

    Case E_MISC_CMDS
        ' /* miscellaneous stuff */
        Select Case lParam
        Case 1
            ' /* enable sticky option */
            g_ConfigSet "sticky_snarls", IIf(g_ConfigGet("sticky_snarls") = "1", "0", "1")

        Case 2
            ' /* show missed panel */
            If Not (g_NotificationRoster Is Nothing) Then _
                g_NotificationRoster.ShowMissedPanel

        End Select


    Case 1
        ' /* stop Snarl */
        g_SetRunning False

    Case 2
        ' /* start Snarl */
        g_SetRunning True

    Case 3
        ' /* restart Snarl */
        g_SetRunning False
        DoEvents
        Sleep 1500
        DoEvents
        g_SetRunning True

    Case 4
        ' /* show prefs panel */
'        uNewDoPrefs

    Case 5
        ' /* reload extensions */
        If Not (g_ExtnRoster Is Nothing) Then _
            g_ExtnRoster.Reload

        frmAbout.bUpdateExtList

    Case 6
        ' /* unload extensions */
'        If Not (g_ExtnRoster Is Nothing) Then _
'            g_ExtnRoster.Unload
'            g_ExtnRoster.SetEnabled False
'            melonLibClose g_ExtnRoster
'            globalExtEnable False
'            globalExtUnload
        Sleep 500
'        bUpdateExtList
        Sleep 500

    Case 7
        ' /* load extensions */
'        If Not (g_ExtnRoster Is Nothing) Then _
'            g_ExtnRoster.Load
'
'            melonLibOpen g_ExtnRoster
'            g_ExtnRoster.SetEnabled True
'            globalExtGetAll
'            globalExtEnable True

'        bUpdateExtList

    Case 8
        ' /* unload specific extension */
        
'Dim sz As String
'
'            g_Debug "WM_MANAGE_SNARL: UNLOAD_EXTENSION: " & lstrlenA(lParam) & " " & lstrlenW(lParam)
'            g_Debug "WM_MANAGE_SNARL: UNLOAD_EXTENSION [W]:" & g_CopyStrW(lParam)
'
'            sz = String$(lstrlenA(lParam) + 1, 0)
'            lstrcpyA sz, lParam
'            g_Debug "WM_MANAGE_SNARL: UNLOAD_EXTENSION [A]:" & g_TrimStr(sz)
'
'            If Not (g_ExtnRoster Is Nothing) Then
'
'                Sleep 500
'            bUpdateExtList
'                Sleep 500
'
'            End If


    Case 9
        ' /* reload config */
        g_ConfigInit


    Case E_MANAGE_STYLE_ROSTER
        ' /*
        '   manage the style roster - lParam defines what to do, as follows:
        '       1 = open style roster
        '       2 = close style roster
        ' */

        Select Case lParam
        Case 1
            ' /* open */
            If Not (g_StyleRoster Is Nothing) Then
                melonLibOpen g_StyleRoster
'                uUpdateStyleList

            End If

        Case 2
            ' /* close */
            If Not (g_StyleRoster Is Nothing) Then
                melonLibClose g_StyleRoster
'                uUpdateStyleList

            End If

        End Select


    Case E_SET_DND_MODE
        g_Debug "TMainWindow.uManageSnarl(): E_SET_DND_MODE is no longer supported", LEMON_LEVEL_WARNING

'        ' /* lParam specifies the mode - 0 = disabled (decrement lock), 1 = enabled (increment lock) */
'        If lParam = 0 Then
'            g_LockDoNotDisturb False
'
'        ElseIf lParam = 1 Then
'            g_LockDoNotDisturb True
'
'        End If


    Case Else
        g_Debug "uManageSnarl(): unknown command '" & CStr(wParam) & "'", LEMON_LEVEL_WARNING

    End Select

End Function

Private Function uCopyData(ByVal lpStruct As Long, ByVal hwndFrom As Long) As Long
Dim pcds As COPYDATASTRUCT

    ' /* handle WM_COPYDATA */

    CopyMemory pcds, ByVal lpStruct, Len(pcds)
    g_Debug "TMainWindow.uCopyData(): dwData=" & CStr(pcds.dwData) & " hWndFrom=" & g_HexStr(hwndFrom)

'    If (pcds.dwData <> 2) And (pcds.dwData <> &H534E4C02) Then
'        ' /* not from Snarl */
'        g_Debug "TMainWindow.uCopyData(): Unknown COPYDATASTRUCT->dwData value '" & CStr(pcds.dwData) & "' from 0x" & g_HexStr(hwndFrom)
'        Exit Function
'
'    End If

    ' /* possible Snarl request */

Dim pss As SNARLSTRUCTI

    If g_IsRunning Then

        Select Case pcds.dwData

        Case &H534E4C03
            ' /* V42 API request */
            uCopyData = uDoV42Request(pcds.cbData, pcds.lpData, hwndFrom)

        Case &H534E4C02
            ' /* V41 API request */
            uCopyData = uDoV41Request(pcds.lpData, hwndFrom)

        Case 2

            ' /* retrieve the SNARLSTRUCT so we can get the command */

            CopyMemory pss, ByVal pcds.lpData, Len(pss)
            g_Debug "TMainWindow.uCopyData(): cmd=" & g_XCommandStr(pss.Cmd) & " id=" & pss.Id & " lngdata2=" & g_HexStr(pss.LngData2), LEMON_LEVEL_PROC

            If pss.Cmd >= SNARL_EX_SHOW Then
                ' /* extended command */
                uCopyData = uDoCmdEx(pcds.lpData, hwndFrom)

            Else
                ' /* standard command */
                uCopyData = uDoCmd(pss, hwndFrom)

            End If

        Case Else
            g_Debug "TMainWindow.uCopyData(): Unknown COPYDATASTRUCT->dwData value '" & CStr(pcds.dwData) & "' from 0x" & g_HexStr(hwndFrom)
            gSetLastError SNARL_ERROR_BAD_PACKET

        End Select

    Else
        g_Debug "TMainWindow.uCopyData(): notification engine is not running", LEMON_LEVEL_WARNING
        gSetLastError SNARL_ERROR_NOT_RUNNING

        ' /* TO-DO: we might want to log the notification somewhere? */

    End If

End Function

Private Function uDoCmd(ByRef pss As SNARLSTRUCTI, ByVal hwndFrom As Long) As Long
Dim pa As TApp
Dim i As Long
Dim j As Long
Dim sz As String

    Select Case pss.Cmd

    Case SNARL_SHOW
'        If hWndFrom = 0 Then
'            g_Debug "uDoCmd(): hWndFrom not provided - can't determine sender", LEMON_LEVEL_WARNING
'
'        Else
'            If g_AppRoster.FindByWindow(hWndFrom, pa) Then
'
''                j = g_AppRoster.AppAt(i).Alerts.IndexOf("_all")
''                g_Debug "uDoCmd(): app found at #" & i & " is '" & g_Applet(i).Name & "'; '_all' is " & CStr(j)
''                If j > 0 Then _
''                    sz = uGetStyleName(g_Applet(i).Alerts.EntryAt(j).Value)
''
''                g_Debug "uDoCmd(): Alert-specifc style is '" & sz & "'"
'
'            Else
'                g_Debug "uDoCmd(): app for config #" & g_HexStr(hWndFrom) & " not found", LEMON_LEVEL_CRITICAL
'
'            End If
'        End If
'
'        uDoCmd = g_NotificationRoster.Add(0, SNARL_CLASS_ANON, g_UTF8(pss.Title), g_UTF8(pss.Text), pss.Timeout, g_UTF8(pss.Icon), pss.LngData2, pss.Id, "", sz, "")

        If g_AppRoster.FindByToken(gSnarlToken, pa, gSnarlPassword) Then
            ' /* generate the alert using the SNARL_CLASS_ANON class so the user can block them if they want */
            uDoCmd = pa.Show(SNARL_CLASS_ANON, g_UTF8(pss.Title), g_UTF8(pss.Text), pss.Timeout, _
                             g_UTF8(pss.Icon), pss.LngData2, pss.Id)

        Else
            ' /* not found */
            g_Debug "uDoCmd('SNARL_SHOW'): Snarl internal app not found", LEMON_LEVEL_CRITICAL
            uDoCmd = M_ABORTED

        End If

    Case SNARL_HIDE
        uDoCmd = g_NotificationRoster.Hide(pss.Id, "", "", "")

    Case SNARL_UPDATE
        uDoCmd = uOldUpdate(pss.Id, g_UTF8(pss.Title), g_UTF8(pss.Text), g_UTF8(pss.Icon))

    Case SNARL_IS_VISIBLE
        uDoCmd = g_NotificationRoster.IsVisible(pss.Id, "", "", "")

    Case SNARL_GET_VERSION
        uDoCmd = MAKELONG(APP_SUB_VER, APP_VER)

    Case SNARL_REGISTER_CONFIG_WINDOW
        uDoCmd = g_AppRoster.OldAdd(g_UTF8(pss.Title), pss.LngData2, pss.Id, "", "", 0)
'        If uDoCmd = M_OK Then _
            bUpdateAppList

    Case SNARL_REVOKE_CONFIG_WINDOW
        uDoCmd = g_AppRoster.OldRemove(pss.LngData2)            ' // returns BOOL
        If uDoCmd Then
'            bUpdateAppList
            uDoCmd = M_OK

        Else
            uDoCmd = M_NOT_FOUND

        End If

    ' /* V37 */

    Case SNARL_REGISTER_ALERT
        uDoCmd = gfRegisterAlert(g_UTF8(pss.Title), g_UTF8(pss.Text), pss.LngData2)

    Case SNARL_REVOKE_ALERT
        If pss.LngData2 = &HFFFE& Then
            ' /* SNARL_GET_REVISION (undocumented) V38.128 */
            uDoCmd = App.Revision

        Else
            uDoCmd = M_NOT_IMPLEMENTED     ' // for future use

        End If

    Case SNARL_GET_VERSION_EX
        uDoCmd = App.Major

    Case SNARL_REGISTER_CONFIG_WINDOW_2
        uDoCmd = g_AppRoster.OldAdd(g_UTF8(pss.Title), pss.LngData2, pss.Id, g_UTF8(pss.Icon), g_UTF8(pss.Text), 0)
'        If uDoCmd = M_OK Then _
            bUpdateAppList

    Case SNARL_SET_TIMEOUT
'        uDoCmd = globalSetTimeout(pss.Id, pss.LngData2)
        uDoCmd = g_NotificationRoster.SetAttribute(pss.Id, SNARL_ATTRIBUTE_TIMEOUT, CStr(pss.LngData2))

    Case SNARL_SET_CLASS_DEFAULT
        ' /* added as private in V38, will be doc'd in V39 */
        uDoCmd = gfSetAlertDefault(pss.Timeout, g_UTF8(pss.Text), pss.LngData2, g_UTF8(pss.Icon))

    Case SNARL_CHANGE_ATTR
        ' /* added as private in V38, will be doc'd in V39 */
        uDoCmd = g_NotificationRoster.SetAttribute(pss.Id, pss.LngData2, g_UTF8(pss.Text))

    Case SNARL_REGISTER_APP
        ' /* added as private in V38, will be doc'd in V39 */
        If Not (g_AppRoster Is Nothing) Then
            uDoCmd = g_AppRoster.Register(g_UTF8(pss.Title), pss.LngData2, pss.Id, g_UTF8(pss.Icon), g_UTF8(pss.Text), pss.Timeout)
'            bUpdateAppList

        Else
            g_Debug "uDoCmd(): [SNARL_REGISTER_APP] app roster is not available", LEMON_LEVEL_CRITICAL

        End If

    Case SNARL_UNREGISTER_APP
        ' /* added as private in V38, will be doc'd in V39 */
        If Not (g_AppRoster Is Nothing) Then
            uDoCmd = g_AppRoster.OldUnregister(pss.LngData2)
'            bUpdateAppList

        Else
            g_Debug "uDoCmd(): [SNARL_UNREGISTER_APP] app roster is not available", LEMON_LEVEL_CRITICAL

        End If

    Case SNARL_ADD_CLASS
        ' /* added as private in V38, will be doc'd in V39 */
        If Not (g_AppRoster Is Nothing) Then
            uDoCmd = gfAddClass(pss.Timeout, g_UTF8(pss.Text), pss.LngData2, g_UTF8(pss.Title))

        Else
            g_Debug "uDoCmd(): [SNARL_UNREGISTER_APP] app roster is not available", LEMON_LEVEL_CRITICAL

        End If

'    With pss
'        .Cmd = SNARL_ADD_CLASS
'        .Text = uToUTF8(Class)
'        .LngData2 = Flags
'        .Timeout = GetCurrentProcessId()
'
'    End With

        ' /* ------------------------------ V40 (private) ------------------------------ */

    Case SNARL_LOAD_EXTENSION
        ' /* added as private in V38, will be doc'd in V40 */
        If Not (g_ExtnRoster Is Nothing) Then
            uDoCmd = g_ExtnRoster.LoadExtension(g_UTF8(pss.Text))

        Else
            g_Debug "uDoCmd(): [SNARL_LOAD_EXTENSION] extension roster is not available", LEMON_LEVEL_CRITICAL

        End If

    Case SNARL_UNLOAD_EXTENSION
        ' /* added as private in V38, will be doc'd in V40 */
        If Not (g_ExtnRoster Is Nothing) Then
            uDoCmd = g_ExtnRoster.UnloadExtension(g_UTF8(pss.Text))

        Else
            g_Debug "uDoCmd(): [SNARL_UNLOAD_EXTENSION] extension roster is not available", LEMON_LEVEL_CRITICAL

        End If

'    Case SNARL_COUNT_NOTIFICATIONS
'        If Not (g_NotificationRoster Is Nothing) Then
'            uDoCmd = g_NotificationRoster.GetCount(pss.Timeout, (pss.LngData2 = 1), g_UTF8(pss.Text))
'
'        Else
'            uDoCmd = M_ABORTED
'
'        End If

    Case SNARL_PREVIEW_SCHEME
        ' /*
        '  * snarlcmd->title = style name
        '  * snarlcmd->text = scheme name
        '  * snarlcmd->timeout = UNUSED
        '  * snarlcmd->lngdata2 = flags (only 0x0001 [priority] defined)
        '  * snarlcmd->id = percent value (must be 1><100 to have any effect)
        '  */

        g_DoSchemePreview2 g_UTF8(pss.Title), g_UTF8(pss.Text), ((pss.LngData2 And 1) <> 0), pss.Id


    Case SNARL_LOAD_STYLE_ENGINE
        ' /* private - not due for documentation */
        If Not (g_StyleRoster Is Nothing) Then
            uDoCmd = IIf(g_StyleRoster.Load(g_UTF8(pss.Text), "", True), M_OK, M_NOT_FOUND)

        Else
            g_Debug "uDoCmd(): [SNARL_LOAD_STYLE_ENGINE] style roster is not available", LEMON_LEVEL_CRITICAL
            uDoCmd = M_FAILED

        End If

    Case SNARL_UNLOAD_STYLE_ENGINE
        ' /* private - not due for documentation */
        If Not (g_ExtnRoster Is Nothing) Then
            uDoCmd = IIf(g_StyleRoster.Unload(g_UTF8(pss.Text), ""), M_OK, M_NOT_FOUND)

        Else
            g_Debug "uDoCmd(): [SNARL_UNLOAD_STYLE_ENGINE] style roster is not available", LEMON_LEVEL_CRITICAL
            uDoCmd = M_FAILED

        End If







    Case Else
        g_Debug "uDoCmd(): Unknown SNARLSTRUCT->Cmd (" & pss.Cmd & ")", LEMON_LEVEL_WARNING
        uDoCmd = M_NOT_IMPLEMENTED

    End Select

End Function

Private Function uDoCmdEx(ByVal lpData As Long, ByVal hwndFrom As Long) As Long
Dim pssex As SNARLSTRUCTEXI
Dim pa As TApp

    CopyMemory pssex, ByVal lpData, Len(pssex)

    Select Case pssex.Cmd

    Case SNARL_EX_SHOW
        If hwndFrom = 0 Then
            g_Debug "uDoCmdEx(): [SNARL_EX_SHOW] hWndFrom not provided - can't determine sender", LEMON_LEVEL_CRITICAL
            uDoCmdEx = M_BAD_HANDLE
            Exit Function

        End If

        ' /* find the registered application - for SNARL_EX_SHOW the app must be registered */

        If g_AppRoster.FindByWindow(hwndFrom, pa) Then
            ' /* found */
            uDoCmdEx = pa.Show(g_UTF8(pssex.Class), g_UTF8(pssex.Title), g_UTF8(pssex.Text), pssex.Timeout, _
                               g_UTF8(pssex.Icon), pssex.LngData2, pssex.Id, g_UTF8(pssex.Extra))

        Else
            ' /* not found */
            g_Debug "SNARL_EX_SHOW: app (window) " & g_HexStr(hwndFrom) & " not found", LEMON_LEVEL_CRITICAL
            uDoCmdEx = M_NOT_FOUND

        End If

    Case SNARL_SHOW_NOTIFICATION
        ' /* find the registered application - must be registered */
        If g_AppRoster.FindByPid(pssex.Reserved1, pa) Then
            ' /* found */
            uDoCmdEx = pa.Show(g_UTF8(pssex.Class), g_UTF8(pssex.Title), g_UTF8(pssex.Text), pssex.Timeout, _
                               g_UTF8(pssex.Icon), pssex.LngData2, pssex.Id, g_UTF8(pssex.Extra))

        Else
            ' /* not found */
            g_Debug "SNARL_SHOW_NOTIFICATION: app '" & g_UTF8(pssex.Extra2) & "' not found", LEMON_LEVEL_CRITICAL
            uDoCmdEx = M_NOT_FOUND

        End If

    Case Else
        g_Debug "uDoCmdEx(): Unknown SNARLSTRUCT->Cmd (" & pssex.Cmd & ")", LEMON_LEVEL_WARNING
        uDoCmdEx = M_NOT_IMPLEMENTED

    End Select

End Function

Private Function uDoV42Request(ByVal cbRequest As Long, ByVal lpRequest As Long, ByVal SenderPID As Long) As Long
Dim sz As String

    ' /* easy: get the request, do some pre-processing on it and send it over to g_DoAction()
    '    which handles everything for us, including setting/clearing LastError */

    Debug.Print "TMainWindow.uDoV42Request(): data is " & cbRequest & " byte(s)"



    ' /* V42 requests are variable length */

    sz = String(cbRequest, 0)
    CopyMemory ByVal StrPtr(sz), ByVal lpRequest, cbRequest
    sz = g_UTF8(sz)

    Debug.Print "TMainWindow.uDoV42Request(): data is '" & sz & "'"

    ' /* must at least have an action */

    If sz = "" Then
        gSetLastError SNARL_ERROR_BAD_PACKET
        Exit Function

    End If

    ' /* tokenise special character pairs */

    sz = Replace$(sz, "&&", "%26")
    sz = Replace$(sz, "==", "%3d")

    ' /* convert to internal format (note that 'real' equals and ampersands
    '    will still be url-encoded so they're safe */

    sz = Replace$(sz, "=", "::")
    sz = Replace$(sz, "&", "#?")

Dim szAction As String
Dim i As Long

    ' /* find the action/arg marker */

    i = InStr(sz, "?")
    If i Then
        ' /* action and args */
        szAction = g_SafeLeftStr(sz, i - 1)
        sz = g_SafeRightStr(sz, Len(sz) - i)

    Else
        ' /* just an action */
        szAction = sz
        sz = ""

    End If

    ' /* set the packed data and URL-decode the arguments at the same time */

Dim pData As BPackedData
Dim lToken As Long

    Set pData = New BPackedData
    pData.SetTo g_DecodeURL(sz)
    lToken = g_SafeLong(pData.ValueOf("token"))

    ' /* pass to master function */

    uDoV42Request = g_DoAction(szAction, lToken, pData, App.Major)      ' // R2.4.1 include major API version in this

    If uDoV42Request = 0 Then
        uDoV42Request = -g_QuickLastError()

    ElseIf uDoV42Request = -1 Then
        uDoV42Request = 0

    End If

End Function

Private Function uDoV41Request(ByVal lpRequest As Long, ByVal SenderPID As Long) As Long
Dim pReq As SNARLREQI

    ' /* this is much neater under V42: we convert the command into an action
    '    name and send it over to g_DoAction() which handles everything for
    '    us, including setting/clearing LastError */

    CopyMemory pReq, ByVal lpRequest, Len(pReq)

Dim szAction As String

    ' /* set the packed data */

Dim pData As BPackedData

    Set pData = New BPackedData
    pData.SetTo g_UTF8(pReq.PacketData)

    ' /* parse the command */

    Select Case pReq.Command

    Case SNARLX41_REGISTER_APP
        szAction = "reg"

    Case SNARLX41_UNREGISTER_APP
        szAction = "unreg"

    Case SNARLX41_UPDATE_APP
        szAction = "updateapp"

    Case SNARLX41_ADD_CLASS
        szAction = "addclass"

    Case SNARLX41_REMOVE_CLASS
        szAction = "remclass"

    Case SNARLX41_LAST_ERROR
        ' /* old error handling */
        uDoV41Request = Abs(GetProp(ghWndMain, "last_error"))
        Exit Function

    Case SNARLX41_NOTIFY
        szAction = "notify"

    Case SNARLX41_UPDATE_NOTIFICATION
        szAction = "update"

    Case SNARLX41_HIDE_NOTIFICATION
        szAction = "hide"

    Case SNARLX41_IS_NOTIFICATION_VISIBLE
        szAction = "isvisible"

    Case Else
        ' /* no need to process these */
        gSetLastError SNARL_ERROR_UNKNOWN_COMMAND
        Exit Function

    End Select

    uDoV41Request = g_DoAction(szAction, pReq.Token, pData)

End Function

Private Function uOldUpdate(ByVal Token As Long, ByVal Title As String, ByVal Text As String, ByVal IconPath As String) As M_RESULT

    ' /* Only include anything that isn't an empty string.  The known bug here is, of course, that
    '    the user can't clear a particular item, although they can use SetAttribute() instead */

Dim ppd As BPackedData

    Set ppd = New BPackedData
    If Title <> "" Then _
        ppd.Add "title", Title

    If Text <> "" Then _
        ppd.Add "text", Text

    If IconPath <> "" Then _
        ppd.Add "icon", IconPath

    If g_NotificationRoster.Update(Token, ppd) Then
        uOldUpdate = M_OK

    Else
        uOldUpdate = M_FAILED

    End If

End Function

Private Function uAtomName(ByVal Atom As Long) As String
Dim sz As String
Dim hr As Long

    sz = String$(1024, 0)
    hr = GetClipboardFormatName(Atom, sz, Len(sz))
    If hr Then _
        uAtomName = g_SafeLeftStr(sz, hr)

End Function

Private Function uConfigure(ByVal Item As String) As Boolean
Dim i As Long

    Item = Replace$(Item, "%20", " ")
    Debug.Print "TMainWindow.uConfigure(): " & Item

    Select Case LCase$(g_GetExtension(Item))
    Case "extension"
        If (g_ExtnRoster Is Nothing) Then _
            Exit Function

        i = g_ExtnRoster.IndexOf(g_RemoveExtension(Item))
        If i Then
            frmAbout.NewDoPrefs 5
            uConfigure = frmAbout.DoExtensionConfig(i)

        End If

    Case "style"
        If (g_StyleRoster Is Nothing) Then _
            Exit Function

        i = g_StyleRoster.IndexOf(g_RemoveExtension(Item))
        Debug.Print "style: " & i

        If i Then
            frmAbout.NewDoPrefs 4
            uConfigure = frmAbout.DoStyleConfig(i)

        End If

    End Select

End Function

'Private Sub uInstall(ByVal Item As String)
'
'End Sub
