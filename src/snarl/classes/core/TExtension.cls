VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Sub CoFreeLibrary Lib "ole32.dll" (ByVal hInst As Long)

Public Enum E_EXTENSION_STATE
    E_ES_FAILED = -1                '// mError should contain reason
    E_ES_NOT_LOADED = 0
    E_ES_LOADED = 1

End Enum

Dim mClass As String
Dim mPath As String
Dim mState As E_EXTENSION_STATE
Dim mInfo As extension_info

Dim mObj As MWndProcSink
Dim mObj41 As ISnarlExtension
Dim mLastErr As String

Dim mhWndPrefsPanel As Long

Implements mObject
Implements MWndProcSink

Private Property Get MObject_Type() As String

End Property

Public Sub SetTo(ByVal Class As String, ByVal Path As String)

    g_Debug "TExtension.SetTo()", LEMON_LEVEL_PROC_ENTER

    mClass = Class
    mPath = Path

    ' /* uInit() attempts to create the extension object and, if successful, queries information
    '    about it - it does not start or initialize the extension */

    If uInit() Then
        g_Debug "'" & Class & "': ok"

    Else
        g_Debug "'" & Class & "': failed (" & mLastErr & ")"

    End If

    Set mObj = Nothing
    Set mObj41 = Nothing

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Function uInit() As Boolean         ' // ByVal Class As String, ByVal Path As String
Dim iExt As ISnarlExtension
Dim iSink As MWndProcSink
Dim iVer As MVersionInfo

    ' /* creates, but does _not_ initialize, the extension */

    On Error Resume Next

    g_Debug "TExtension.uInit('" & mClass & "')", LEMON_LEVEL_PROC_ENTER

    mLastErr = ""
    mState = E_ES_NOT_LOADED

    err.Clear
    Set iExt = CreateObject(mClass & ".extension")
    If (err.Number = 0) And (Not (iExt Is Nothing)) Then
        ' /* is V41 extension */
        g_Debug "V41 extension"
        iExt.GetInfo mInfo
        Set mObj41 = iExt

    Else
        ' /* must support MVersionInfo */
        g_Debug "not V41 extension, checking MVersionInfo support..."
        err.Clear
        Set iVer = CreateObject(mClass & ".extension")
        If err.Number <> 0 Then
            g_Debug "doesn't support MVersionInfo (" & err.Description & ")", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
            mLastErr = "Not a valid extension"
            mState = E_ES_FAILED
            Exit Function

        End If

        ' /* must support MWndProcSink */
        g_Debug "checking MWndProcSink support..."
        err.Clear
        Set iSink = iVer
        If err.Number <> 0 Then
            g_Debug "doesn't support MWndProcSink (" & err.Description & ")", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
            mLastErr = "Not a valid extension"
            mState = E_ES_FAILED
            Exit Function

        End If

Dim iEnum As MSimpleEnum
Dim sz As String
Dim dw As Long

        If iVer.Version >= 40 Then
            ' /* V40 extensions must support MSimpleEnum */
            g_Debug "V40 extension, checking MSimpleEnum support..."
            err.Clear
            Set iEnum = iVer
            If err.Number <> 0 Then
                g_Debug "doesn't support MSimpleEnum", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
                mLastErr = "Not a valid extension"
                mState = E_ES_FAILED
                Exit Function

            End If

            ' /* V40-specific stuff */

            With mInfo
                iEnum.FindItem "release", sz
                .Release = sz
                .Version = iVer.Version
                .Revision = iVer.Revision
                ' /* get flags */
                dw = 0
                iSink.WndProc HWND_SNARL, SNARL_EXT_GET_FLAGS, 0, 0, 0, dw
                .Flags = dw

            End With

        Else
            ' /* pre-V40 extensions can support MSimpleEnum (error trap will catch it if not) */

            g_Debug "is V39 or earlier"
            Set iEnum = iVer
            With mInfo
                .Release = CStr(iVer.Version) & "." & CStr(iVer.Revision)
                .Flags = SNARL_EXTN_IS_CONFIGURABLE                 ' // (have to assume this)

            End With

        End If

        ' /* okay */

        Set mObj = iSink

        With mInfo
            ' /* general stuff */
            .Date = iVer.Date
            .Description = iVer.Name

            ' /* if the object also supports MSimpleEnum, we can get more info from it */

            If Not (iEnum Is Nothing) Then

                If iEnum.FindItem("path", sz) Then _
                    .Path = sz

                If iEnum.FindItem("name", sz) Then _
                    .Name = sz

                If iEnum.FindItem("copyright", sz) Then _
                    .Copyright = sz

                If iEnum.FindItem("url", sz) Then _
                    .URL = sz

            End If

            If .Path = "" Then _
                .Path = mPath

            .IconPath = g_MakePath(.Path) & "icon.png"

            ' /* fix this manually if the extension doesn't provide it */

            If .Name = "" Then _
                .Name = Class

        End With

    End If

    uInit = True

End Function

Public Function Name() As String

    Name = IIf(mInfo.Name = "", mClass, mInfo.Name)

End Function

Public Function Description() As String

'    MsgBox mClass & " > " & mState

    Select Case mState
    Case E_ES_FAILED
        Description = "Failed: " & mLastErr

    Case E_ES_NOT_LOADED
        Description = "Not loaded" ' IIf(mInfo.Description = "", "Failed: " & mLastErr, mInfo.Description)

    Case E_ES_LOADED
        Description = mInfo.Description

    End Select

End Function

Public Function State() As E_EXTENSION_STATE

    State = mState

End Function

Public Function IsConfigurable() As Boolean

    If mState = E_ES_LOADED Then _
        IsConfigurable = (mInfo.Flags And SNARL_EXTN_IS_CONFIGURABLE)

End Function

'Public Function IsEnabled() As Boolean
'
'    IsEnabled = mEnabled
'
'End Function
'
'Public Function StartedOkay() As Boolean
'
'    StartedOkay = mStarted
'
'End Function

Public Function Icon() As String

    Icon = mInfo.IconPath

End Function

Public Function Path() As String

    Path = mInfo.Path

End Function

Public Function SetEnabled(ByVal Enabled As Boolean) As Boolean

    If Enabled Then
        SetEnabled = uEnable()

    Else
        SetEnabled = uDisable(True)

    End If

End Function

Private Function uEnable() As Boolean

    On Error Resume Next

    If mState = E_ES_LOADED Then
        g_Debug "TExtension.uEnable(): '" & mClass & "' is already loaded", LEMON_LEVEL_WARNING
        Exit Function

    End If

    ' /* create instance */

    If Not uInit() Then '//mClass, mInfo.Path
        g_Debug "TExtension.uEnable(): uInit() failed on '" & mClass & "'", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim dw As Long

    If Not (mObj41 Is Nothing) Then
        ' /* V41 */
        dw = mObj41.Initialize()
        If dw <> 0 Then
            ' /* failed */
            uGetLastError
            g_Debug "TExtension.uEnable(): '" & mInfo.Name & "' failed to initialize (" & mLastErr & ")", LEMON_LEVEL_WARNING
            mState = E_ES_FAILED

        Else
            ' /* success */
            mObj41.Start
            mState = E_ES_LOADED

        End If

    Else
        ' /* V40 or earlier, so send it a SNARL_EXT_INIT */
        dw = M_FAILED
        uCallProc SNARL_EXT_INIT, 0, 0, dw
        If dw = M_OK Then
            ' /* no return value for SNARL_EXT_START */
            uCallProc SNARL_EXT_START, 0, 0, 0
            mState = E_ES_LOADED

        Else
            ' /* extension returned error *//
            uGetLastError
            g_Debug "TExtension.uEnable(): '" & mInfo.Name & "' failed SNARL_EXT_INIT (0x" & g_HexStr(dw) & "): " & mLastErr
            mState = E_ES_FAILED

        End If

    End If

    uEnable = True

End Function

Private Function uDisable(ByVal UnloadNow As Boolean) As Boolean

    ' /* can't disable a failed extension */

    If mState = E_ES_FAILED Then _
        Exit Function

    If Not (mObj41 Is Nothing) Then
        ' /* V41 */
        If mState = E_ES_LOADED Then
            ' /* only stop it if it's loaded */
            mObj41.Stop
            mObj41.TidyUp

        Else
            g_Debug "TExtension.uDisable(): '" & mClass & "' is not loaded", LEMON_LEVEL_WARNING

        End If

        ' /* trash it */
        Set mObj41 = Nothing

    Else
        ' /* pre-V41 */
        If mState = E_ES_LOADED Then
            ' /* no return value for SNARL_EXT_STOP */
            uCallProc SNARL_EXT_STOP, 0, 0, 0

            ' /* no return value for SNARL_EXT_QUIT */
            uCallProc SNARL_EXT_QUIT, 0, 0, 0

        Else
            g_Debug "TExtension.uDisable(): '" & mClass & "' is not loaded", LEMON_LEVEL_WARNING

        End If

        ' /* trash it */
        Set mObj = Nothing

    End If


    If UnloadNow Then
        If g_IsWinXPOrBetter() Then
            CoFreeUnusedLibrariesEx 0, 0

        Else
            ' /* for pre-Windows XP systems, we try to flush the object out using CoFreeLibrary() */
'            hModule = GetModuleHandle(.Name)
'            g_Debug "TExtension.uTrash(): hModule('" & .Name & "')=0x" & g_HexStr(hModule)
'            CoFreeLibrary hModule

        End If
    End If

    mState = E_ES_NOT_LOADED
    uDisable = True

End Function

Public Function DoPrefs(ByVal hWndPrefsPanel As Long) As Boolean

'    If Not mEnabled Then
'        g_Debug "TExtension.DoPrefs(): '" & mInfo.Name & "' is not enabled", LEMON_LEVEL_WARNING
'        Exit Function
'
'    End If
'
'    If (mInfo.Flags And SNARL_EXTN_IS_CONFIGURABLE) = 0 Then
'        g_Debug "TExtension.DoPrefs(): '" & mInfo.Name & "' is not configurable", LEMON_LEVEL_WARNING
'        Exit Function
'
'    End If

    If Not Me.IsConfigurable Then
        g_Debug "TExtension.DoPrefs(): '" & mClass & "' is not configurable", LEMON_LEVEL_WARNING
        Exit Function

    End If

Dim hWndConfig As Long

    If Not (mObj41 Is Nothing) Then
        ' /* V41 */
        hWndConfig = mObj41.GetConfigWindow()

    ElseIf Not (mObj Is Nothing) Then
        ' /* send it a SNARL_EXT_PREFS */

        mObj.WndProc HWND_SNARL, SNARL_EXT_PREFS, 0, hWndPrefsPanel, 0, hWndConfig
        g_Debug "TExtension.DoPrefs(): '" & mInfo.Name & "' SNARL_EXT_PREFS=" & g_HexStr(hWndConfig)

        If mInfo.Version >= 40 Then _
            g_Debug "TExtension.DoPrefs(): V40 extension..."

    Else
        g_Debug "TExtension.DoPrefs(): '" & mInfo.Name & "' is not valid"
        Exit Function

    End If


    If IsWindow(hWndConfig) <> 0 Then
        mhWndPrefsPanel = hWndPrefsPanel
        g_MoveWindow hWndConfig, 0, 0, 0, 0, True, mhWndPrefsPanel
        g_ShowWindow hWndConfig, True, True

        AddSubClass hWndConfig, Me
        SetWindowLong hWndConfig, GWL_HWNDPARENT, mhWndPrefsPanel

        EnableWindow mhWndPrefsPanel, 0
        DoPrefs = True

    Else
        g_Debug "TExtension.DoPrefs(): extension returned invalid window", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Sub CallProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

'    If menabled Then _
        CallProc = uCallProc(uMsg, wParam, lParam, Failed)

End Sub

    ' /* calls the Extension's MWndProcSink() hook with the specified parameters.  Returns True if everything
    '    went okay (Result contains the returned value), False if there was a serious error.  If this function
    '    returns False, LastError() is automatically set */

Private Function uCallProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef Result As Long) As Boolean

    On Error GoTo ex

    mLastErr = ""
    uCallProc = mObj.WndProc(HWND_SNARL, uMsg, wParam, lParam, 0, Result)
    Exit Function

ex: mLastErr = "Critical initialisation error, please contact the Extension vendor"
    uCallProc = False

End Function

Private Function uGetLastError() As String

    On Error Resume Next

    mLastErr = ""

    If Not (mObj41 Is Nothing) Then
        ' /* V41 */
        mObj41.LastError mLastErr

    ElseIf Not (mObj Is Nothing) Then
        ' /* pre-V41 */

Dim pse As MSimpleEnum

        err.Clear
        Set pse = mObj
        If err.Number = 0 Then _
            pse.FindItem "last_error", mLastErr

    End If

End Function

Public Function InfoString() As String

    InfoString = CStr(mInfo.Release) & IIf(mInfo.Date <> "", " (" & mInfo.Date & ")", "") & IIf(mInfo.Copyright <> "", " " & mInfo.Copyright, "")

End Function

Public Function VerString(Optional ByVal IncludeDate As Boolean = True) As String

    VerString = CStr(mInfo.Release) & IIf(IncludeDate, IIf(mInfo.Date <> "", " (" & mInfo.Date & ")", ""), "")

End Function

Public Function LastError() As String

    LastError = mLastErr

End Function

'Public Function Copyright() As String
'
'    Copyright = mInfo.Copyright
'
'End Function
'
'Public Function SupportURL() As String
'
'    SupportURL = mInfo.URL
'
'End Function

'Public Function FriendlyName() As String
'
'    ' /* TO DO: determine if we need to distinguish between name and friendlyname */
'
'    FriendlyName = mInfo.Name
'
''    FriendlyName = mInfo.FriendlyName
'
'End Function

Private Function MWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If uMsg = WM_DESTROY Then
        ' /* config window has gone */
        RemoveSubClass hWnd                     ' // really necessary?
        EnableWindow mhWndPrefsPanel, -1
        g_ShowWindow mhWndPrefsPanel, True, True
        mhWndPrefsPanel = 0

    End If

End Function

'Public Function SysVer() As Long
'
'    SysVer = mInfo.Version
'
'End Function

'Public Function SysRev() As Long
'
'    SysRev = mInfo.Revision
'
'End Function

'Public Function Flags() As SNARL_EXTENSION_FLAGS
'
'    Flags = mInfo.Flags
'
'End Function

Public Sub Pulse()

    If (mObj41 Is Nothing) Or ((mInfo.Flags And SNARL_EXTN_WANTS_PULSE) = 0) Or (mState <> E_ES_LOADED) Then _
        Exit Sub

    mObj41.Pulse

End Sub

Public Function Class() As String

    Class = mClass

End Function

Public Sub ShowDetails()
Dim sz As String

    With mInfo
        ' /* R2.4 */
        sz = "Version: " & IIf(.Version = 0, "?", CStr(.Version)) & "." & IIf(.Revision = 0, "?", CStr(.Revision))

        If .Date <> "" Then _
            sz = sz & " (" & .Date & ")"

        If .Author <> "" Then _
            sz = sz & "\n" & .Author

        If .Copyright <> "" Then _
            sz = sz & "\n" & .Copyright

        If (.SupportEmail <> "") Or (.URL <> "") Then _
            sz = sz & "\n"

        If .SupportEmail <> "" Then _
            sz = sz & "\nSupport: " & .SupportEmail

        If .URL <> "" Then _
            sz = sz & "\nWebsite: " & .URL

    End With

    g_NotificationRoster.Hide gExtDetailsToken, "", "", ""

    gExtDetailsToken = g_PrivateNotify("", mInfo.Name & " " & mInfo.Release, sz, , mInfo.IconPath)

    If gExtDetailsToken <> 0 Then
        If (mInfo.SupportEmail <> "") And (InStr(mInfo.SupportEmail, "@") <> 0) Then _
            g_QuickAddAction gExtDetailsToken, "Contact vendor...", "mailto:" & mInfo.SupportEmail

        If (g_SafeLeftStr(mInfo.URL, 7) = "http://") Or (g_SafeLeftStr(mInfo.URL, 8) = "https://") Then _
            g_QuickAddAction gExtDetailsToken, "Visit website...", mInfo.URL

    End If

End Sub


