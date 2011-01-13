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

'Public Enum E_EXTENSION_STATUS
'    E_NOT_AN_EXTENSION              ' // couldn't create the object or assign it to MVersionInfo/MWndProcSink
'    E_NOT_ENABLED                   ' // creatable but not enabled (in exclude list)
'    E_ENABLED                       ' // enabled, started ok
'    E_FAILED_TO_START               ' // enabled, failed to start
'
'End Enum
    
Dim mInfo As extension_info

Dim mClass As String

Dim mObj As MWndProcSink
Dim mObj2 As ISnarlExtension
Dim mEnabled As Boolean             '// means "not in exclude list"
Dim mStarted As Boolean             '// only applies if mEnabled is True - means extension started ok
Dim mLastErr As String

Dim mhWndPrefsPanel As Long

Implements MObject
Implements MWndProcSink

Private Property Get MObject_Type() As String

End Property

Public Function SetTo(ByVal Class As String, ByVal Path As String) As Boolean

    ' /* first (and only) check to see if the passed class really is an extension */

    g_Debug "TExtension.SetTo()", LEMON_LEVEL_PROC_ENTER

    SetTo = uInit(Class, Path)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function uInit(ByVal Class As String, ByVal Path As String) As Boolean
Dim iExt As ISnarlExtension
Dim iVer As MVersionInfo
Dim hr As Long

    ' /* reminder! the reason for breaking out the object creation into a seperate
    '    function is so that enabling and disabling an extension cause Snarl to
    '    (a) confirm the extension is (still) valid and (b) to retrieve information
    '    about the extension each time.  This function is called by SetTo()
    '    whenever the extension is first discovered and subsequently by uEnable()
    '    whenever it needs to be started up */

    On Error Resume Next

    g_Debug "TExtension.uInit()", LEMON_LEVEL_PROC_ENTER

    g_Debug "clsid is '" & Class & "', testing for V41 API..."

    err.Clear
    Set iExt = CreateObject(Class & ".extension")
    If (err.Number = 0) And (Not (iExt Is Nothing)) Then
        ' /* is V41 extension */
        g_Debug "V41 extension"
        iExt.GetInfo mInfo
        Set mObj2 = iExt

    Else

        ' /* must support MVersionInfo */

        g_Debug "not V41 extension, checking MVersionInfo support..."

        err.Clear
        Set iVer = CreateObject(Class & ".extension")
        If err.Number <> 0 Then
            g_Debug "doesn't support MVersionInfo (" & err.Description & ")", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
            Exit Function

        End If

        ' /* must support MWndProcSink */

Dim iSink As MWndProcSink

        g_Debug "checking MWndProcSink support..."

        err.Clear
        Set iSink = iVer
        If err.Number <> 0 Then
            g_Debug "doesn't support MWndProcSink (" & err.Description & ")", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
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
                .Path = Path

            .IconPath = g_MakePath(.Path) & "icon.png"

            ' /* fix this manually if the extension doesn't provide it */

            If .Name = "" Then _
                .Name = Class

        End With

    End If

    mClass = Class

    g_Debug "success! name='" & mInfo.Name & "' release='" & mInfo.Release & "' (V" & CStr(mInfo.Version) & "." & CStr(mInfo.Revision) & ")", LEMON_LEVEL_INFO Or LEMON_LEVEL_PROC_EXIT
    uInit = True

End Function

Public Function Name() As String

    Name = mInfo.Name

End Function

Public Function Description() As String

    Description = mInfo.Description

End Function

Public Function IsEnabled() As Boolean

    IsEnabled = mEnabled

End Function

Public Function StartedOkay() As Boolean

    StartedOkay = mStarted

End Function

Public Function Icon() As String

    Icon = mInfo.IconPath

End Function

Public Function Path() As String

    Path = mInfo.Path

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    If Enabled Then
        uEnable

    Else
        uDisable True

    End If

End Sub

Private Sub uEnable()

    On Error Resume Next

    If mEnabled Then _
        Exit Sub

    mStarted = False
    mLastErr = ""

    ' /* create instance */

    uInit mClass, mInfo.Path

    ' /* is always enabled, even if it doesn't start... */

    mEnabled = True

Dim dw As Long

    If Not (mObj2 Is Nothing) Then
        ' /* V41 */

        dw = mObj2.Initialize()
        If dw <> 0 Then
            ' /* failed */
            uGetLastError
            g_Debug "TExtension.uEnable(): '" & mInfo.Name & "' failed to initialize (" & mLastErr & ")", LEMON_LEVEL_WARNING

        Else
            ' /* success */
            mObj2.Start
            mStarted = True

        End If

    Else
        ' /* send it a SNARL_EXT_INIT */

        dw = M_FAILED
        uCallProc SNARL_EXT_INIT, 0, 0, dw

        If dw = M_OK Then
            ' /* no return value for SNARL_EXT_START */
            uCallProc SNARL_EXT_START, 0, 0, 0
            mStarted = True

        Else
            ' /* extension returned error *//
            g_Debug "TExtension.uEnable(): '" & mInfo.Name & "' responded with 0x" & g_HexStr(dw) & " to SNARL_EXT_INIT"
            uGetLastError

            g_Debug "TExtension.uEnable(): lasterror is '" & mLastErr & "'"

        End If

    End If

End Sub

Private Sub uDisable(ByVal UnloadNow As Boolean)

'    If Not mEnabled Then _
        Exit Sub

    If Not (mObj2 Is Nothing) Then
        ' /* V41 */

        If mStarted Then
            mObj2.Stop
            mObj2.Tidyup

        End If

        ' /* trash it */
        Set mObj2 = Nothing

    Else
        ' /* pre-V41 */

        If mStarted Then
            ' /* no return value for SNARL_EXT_STOP */
            uCallProc SNARL_EXT_STOP, 0, 0, 0

            ' /* no return value for SNARL_EXT_QUIT */
            uCallProc SNARL_EXT_QUIT, 0, 0, 0

        End If

        ' /* trash it */
        Set mObj = Nothing

    End If

    If UnloadNow Then
        If Not g_IsWinXPOrBetter() Then
        ' /* for pre-Windows XP systems, we try to flush the object out using CoFreeLibrary() */
'            hModule = GetModuleHandle(.Name)
'            g_Debug "TExtension.uTrash(): hModule('" & .Name & "')=0x" & g_HexStr(hModule)
'            CoFreeLibrary hModule

        Else
            ' /* for XP and beyond we use this instead */
            CoFreeUnusedLibrariesEx 0, 0

        End If

    End If

    mStarted = False
    mEnabled = False

End Sub

'Public Function SetEnabled2(ByVal Enabled As Boolean, Optional ByVal DelayUnload As Boolean, Optional ByVal WriteConfig As Boolean = True) As Boolean
'Dim bErr As Boolean
'Dim dw As Long
'
'    ' /* nothing to do? */
'
''    If Enabled = menabled Then _
''        Exit Function
''
''    If Enabled Then
''
''        ' /* create the object */
''
''        If Not bInit(mInfo.Name, mInfo.Path, False) Then
''            g_Debug "TExtension.SetEnabled2(): failed to create instance of '" & mInfo.Name & "'", LEMON_LEVEL_CRITICAL
''            mlasterr = "Unable to create instance, please contact the Extension vendor"
''            Exit Function
''
''        End If
''
''        ' /* send it a SNARL_EXT_INIT */
''
''        If Not uCallProc(SNARL_EXT_INIT, 0, 0, dw) Then
''            ' /* lasterror is already set */
''            g_Debug "TExtension.SetEnabled2(): " & mInfo.Name & ": uCallProc(SNARL_EXT_INIT) failed", LEMON_LEVEL_CRITICAL
''            Exit Function
''
''        End If
''
''        If dw <> M_OK Then
''            ' /* extension returned error *//
''            mlasterr = uGetLastError(mInfo.Object)
''            g_Debug "TExtension.SetEnabled2(): SNARL_EXT_INIT returned error code " & g_HexStr(dw) & " last_error=" & mlasterr
''            Exit Function
''
''        End If
''
''        ' /* call SNARL_EXT_START if the extension isn't already running */
''
''        If Not mstarted Then
''            If Not uCallProc(SNARL_EXT_START, 0, 0, dw) Then
''                ' /* lasterror is already set */
''                g_Debug "TExtension.SetEnabled2(): " & mInfo.Name & ": uCallProc(SNARL_EXT_START) failed", LEMON_LEVEL_CRITICAL
''                Exit Function
''
''            End If
''
''            If dw <> M_OK Then
''                ' /* extension returned error *//
''                mlasterr = uGetLastError(mInfo.Object)
''                g_Debug "TExtension.SetEnabled2(): SNARL_EXT_START returned error code " & g_HexStr(dw) & " last_error=" & mlasterr
''                Exit Function
''
''            End If
''
''            g_Debug "TExtension.SetEnabled2(): SNARL_EXT_START=" & g_HexStr(dw)
''            mstarted = True
''
''        End If
''
''    Else
''
''        ' /* do SNARL_EXT_STOP but only if the extension is running */
''
''        If mstarted Then
''            dw = uCallProc(SNARL_EXT_STOP, 0, 0, bErr)
''            g_Debug "TExtension.SetEnabled2(): SNARL_EXT_STOP=" & g_HexStr(dw)
''            mstarted = False
''
''        End If
''
''        ' /* send the extension a SNARL_EXT_QUIT */
''
''        dw = uCallProc(SNARL_EXT_QUIT, 0, 0, bErr)
''        g_Debug "TExtension.SetEnabled2(): SNARL_EXT_QUIT=" & g_HexStr(dw)
''
''        ' /* trash it */
''        Set mInfo.Object = Nothing
''
''        If Not DelayUnload Then
''            If Not g_IsWinXPOrBetter() Then
''            ' /* for pre-Windows XP systems, we try to flush the object out using CoFreeLibrary() */
''    '            hModule = GetModuleHandle(.Name)
''    '            g_Debug "TExtension.uTrash(): hModule('" & .Name & "')=0x" & g_HexStr(hModule)
''    '            CoFreeLibrary hModule
''
''            Else
''                ' /* for XP and beyond we use this instead */
''                CoFreeUnusedLibrariesEx 0, 0
''
''            End If
''        End If
''    End If
''
''    menabled = Enabled
''
''    ' /* write out the exclude file */
''
''    If WriteConfig Then _
''        g_ExtnRoster.WriteExcludeList
''
''    SetEnabled2 = True
'
'End Function

Public Function DoPrefs(ByVal hWndPrefsPanel As Long) As Boolean

    If Not mEnabled Then
        g_Debug "TExtension.DoPrefs(): '" & mInfo.Name & "' is not enabled", LEMON_LEVEL_WARNING
        Exit Function

    End If

    If (mInfo.Flags And SNARL_EXTN_IS_CONFIGURABLE) = 0 Then
        g_Debug "TExtension.DoPrefs(): '" & mInfo.Name & "' is not configurable", LEMON_LEVEL_WARNING
        Exit Function

    End If
    
Dim hWndConfig As Long

    If Not (mObj2 Is Nothing) Then
        ' /* V41 */
        hWndConfig = mObj2.GetConfigWindow()

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

    If Not (mObj2 Is Nothing) Then
        ' /* V41 */
        mObj2.LastError mLastErr

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

Public Function Copyright() As String

    Copyright = mInfo.Copyright

End Function

Public Function SupportURL() As String

    SupportURL = mInfo.URL

End Function

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

Public Function SysVer() As Long

    SysVer = mInfo.Version

End Function

Public Function SysRev() As Long

    SysRev = mInfo.Revision

End Function

Public Function Flags() As SNARL_EXTENSION_FLAGS

    Flags = mInfo.Flags

End Function

Public Sub Pulse()

    If (mObj2 Is Nothing) Or ((mInfo.Flags And SNARL_EXTN_WANTS_PULSE) = 0) Or (Not mStarted) Then _
        Exit Sub

    mObj2.Pulse

End Sub

Public Function Class() As String

    Class = mClass

End Function

Public Sub ShowDetails()
Dim sz As String

    With mInfo
        ' /* R2.4 */
        If (.Version <> 0) And (.Revision <> 0) Then _
            sz = "Version: " & CStr(.Version) & "." & CStr(.Revision)

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


