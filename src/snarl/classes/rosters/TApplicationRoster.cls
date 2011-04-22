VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TApplicationRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TApplicationRoster.cls
'/
'/  Description:    Manages registered applications
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Dim mApp() As TApp
Dim mCount As Long
Dim mIndex As Long

Dim mRefs As Long
Dim mAppToken As Long       ' // V41
Dim mSnarlAppMsg As Long

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */

        Me.Unregister gSnarlToken, gSnarlPassword
        SendToAll SNARL_BROADCAST_QUIT

        ReDim mApp(0)
        mCount = 0

    End If

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT
Dim pAppInfo As T_SNARL_APP
Dim pSnarlApp As TApp

    g_Debug "TApplicationRoster.Open()", LEMON_LEVEL_PROC

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        mSnarlAppMsg = RegisterWindowMessage("SnarlAppMessage")     ' // V40.25
        mAppToken = &H2C0

        ' /* register ourselves with, er, ourself! */

        g_IsRunning = True
        g_Debug "TApplicationRoster.Open(): registering ourself..."

        ' /* V42: password protect Snarl */

'        gSnarlPassword = uGenerateSnarlPassword()

        With pAppInfo
            .hWnd = frmAbout.hWnd
            .Icon = g_MakePath(App.Path) & "etc\icons\snarl.png"
            .LargeIcon = .Icon
            .Name = App.Title
            .Password = gSnarlPassword
            .Pid = GetCurrentProcessId()
            .Signature = App.ProductName
            .uMsg = &H555

        End With

        gSnarlToken = uAddByStruct(pAppInfo, pSnarlApp)

        If gSnarlToken = 0 Then
            ' /* panic? */

        Else
            g_Debug "TApplicationRoster.Open(): registered as " & g_HexStr(gSnarlToken) & ", adding classes..."

            With pSnarlApp
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_GENERAL & "#?name::Welcome Message and other notifications#?password::" & gSnarlPassword)
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_APP_REG & "#?name::Application registered#?password::" & gSnarlPassword)
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_APP_UNREG & "#?name::Application unregistered#?password::" & gSnarlPassword)

                ' /* V38.97 */
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_JSON & "#?name::Incoming JSON notifications#?password::" & gSnarlPassword)

                ' /* V38.142 */
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_ANON_NET & "#?name::Anonymous remote notifications#?password::" & gSnarlPassword)

                ' /* V39.21 */
                .AddClass g_newBPackedData("id::" & SNARL_CLASS_ANON & "#?name::Anonymous local notifications#?password::" & gSnarlPassword)

                ' /* R2.4 */
'                .AddClass g_newBPackedData("id::" & SNARL_CLASS_LOW_PRIORITY & "#?name::Low priority notifications#?password::" & gSnarlPassword)
'                .AddClass g_newBPackedData("id::" & SNARL_CLASS_SYSTEM & "#?name::Snarl Internal Notifications#?password::" & gSnarlPassword)

            End With
        End If

        g_IsRunning = False

    End If

    g_Debug "TApplicationRoster.Open() done", LEMON_LEVEL_PROC

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mCount

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
Dim i As Long

    If mCount Then
        Name = LCase$(Name)
        For i = 1 To mCount
            If LCase$(mApp(i).Name) = Name Then
                MRoster_FindItem = i
                Exit Function

            End If
        Next i

        g_Debug "TApplicationRoster.FindItem(): app '" & Name & "' not found", LEMON_LEVEL_WARNING

    Else
        g_Debug "TApplicationRoster.FindItem(): no apps registered!", LEMON_LEVEL_CRITICAL

    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.MObject

    If (Index > 0) And (Index <= mCount) Then _
        Set MRoster_ItemAt = mApp(Index)

End Function

Private Function MRoster_NextItem(Item As melon.MObject) As Boolean

    If mIndex <= mCount Then
        Set Item = mApp(mIndex)
        mIndex = mIndex + 1
        MRoster_NextItem = True

    End If

End Function

Private Sub MRoster_Rewind()

    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "application.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

Public Function Add41(ByRef Data As BPackedData, Optional ByVal IsRemote As Boolean) As Long
Dim sz As String
Dim i As Long

    g_Debug "TApplicationRoster.Add41()", LEMON_LEVEL_PROC_ENTER

    ' /* validate args */

    sz = Data.ValueOf("app-sig")
    If sz = "" Then _
        sz = Data.ValueOf("id")

    If (Data.ValueOf("title") = "") Or (sz = "") Then
        g_Debug "required arg missing", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    If InStr(sz, " ") Then
        g_Debug "application signature cannot contain spaces", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* check for duplicate and return existing token */

    If mCount Then
        For i = 1 To mCount
            If mApp(i).Signature = sz Then
                ' /* R2.4 DR7: check same password provided */
                If (mApp(i).Password <> "") And (mApp(i).Password <> Data.ValueOf("password")) Then
                    g_Debug "'" & sz & "' already registered: password mismatch", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
                    gSetLastError SNARL_ERROR_AUTH_FAILURE
                    Add41 = 0
                    Exit Function

                Else
                    g_Debug "'" & sz & "' already registered (token=" & mApp(i).Token & ")", LEMON_LEVEL_PROC_EXIT

                    uUpdate i, Data

                    Add41 = mApp(i).Token
                    Exit Function

                End If
            End If
        Next i
    End If

Dim pInfo As T_SNARL_APP

    ' /* configure the item */

    With pInfo
        .Name = Data.ValueOf("title")
        .Icon = Data.ValueOf("icon")

        ' /* R2.4 Beta 4 - support for Base64-encoded icons */

        If (.Icon = "") And (Data.Exists("icon-base64")) Then _
            .Icon = g_GetBase64Icon(Data.ValueOf("icon-base64"))

        .Signature = sz
        .LargeIcon = .Icon                       ' // no need for two icons: V41 spec recommends 128x128px icon

        If Data.Exists("reply-to") Then
            ' /* V42.22: new name */
            .hWnd = Val(Data.ValueOf("reply-to"))

        Else
            .hWnd = Val(Data.ValueOf("hwnd"))

        End If

        If Data.Exists("reply") Then
            ' /* V42.22: new name */
            .uMsg = Val(Data.ValueOf("reply"))

        Else
            .uMsg = Val(Data.ValueOf("umsg"))

        End If

        .Flags = Val(Data.ValueOf("flags"))
        .Password = Data.ValueOf("password")
        .IsRemote = IsRemote

        g_Debug "name='" & .Name & "' signature='" & .Signature & "' flags=0x" & g_HexStr(.Flags)

    End With

    Add41 = uAddByStruct(pInfo)

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function uAddByStruct(ByRef Info As T_SNARL_APP, Optional ByRef NewAppObject As TApp) As Long
Dim pa As TApp

    ' /* assign token and bump counter on */

    Info.Token = mAppToken
    mAppToken = mAppToken + 4

    Set pa = New TApp

    ' /* unlikely, but bInit() can fail... */

    If Not pa.bInit(Info) Then _
        Exit Function

Dim bShow As Boolean

    ' /* add it to our roster */

    uAlphaAdd pa
    bShow = True

    ' /* don't show if it's us or we're locked */

    If (pa.Name = App.Title) Or (g_IgnoreLock) > 0 Then
        bShow = False

    ElseIf g_ConfigGet("notify_on_first_register") = "1" Then
        ' /* if notify_on_first_register is set, we need to take account of the IsNewApp setting */
        bShow = pa.IsNew()

    End If

Dim lFlags As E_NOTIFICATION_FLAGS
Dim i As Long

    If bShow Then

        ' /* if app has a password or was registered via some IP protocol
        '    add the relevant system flags here */

        If pa.Password <> "" Then _
            lFlags = lFlags Or NF_SECURE

        If pa.IsRemote Then _
            lFlags = lFlags Or NF_REMOTE

        ' /* display the 'app registered' notification */

        i = g_PrivateNotify(SNARL_CLASS_APP_REG, _
                            "Application registered", _
                            pa.Name & " registered successfully with Snarl", _
                            -1, _
                            IIf(pa.Icon = "", g_MakePath(App.Path) & "etc\icons\no_icon.png", pa.Icon), _
                            , , , _
                            lFlags)

        If i Then
            ' /* add the action(s) */
            g_QuickAddAction i, "Manage Notifications...", "!snarl cfg " & pa.Signature
            If pa.HasConfig Then _
                g_QuickAddAction i, pa.Name & " Preferences...", "!snarl manage " & pa.Signature

        End If
    End If

    frmAbout.bUpdateAppList
    Set NewAppObject = pa
    uAddByStruct = pa.Token

End Function

Public Function OldAdd(ByVal Name As String, ByVal hWnd As Long, ByVal uMsg As Long, ByVal Icon As String, ByVal LargeIcon As String, ByVal Pid As Long) As M_RESULT

    g_Debug "TApplicationRoster.OldAdd()", LEMON_LEVEL_PROC_ENTER

Dim pInfo As T_SNARL_APP
Dim pa As TApp

    If Me.FindByWindow(hWnd, pa) <> 0 Then
        g_Debug "config window is already registered", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        OldAdd = M_ALREADY_EXISTS
        Exit Function

    End If

    ' /* configure the item */

    With pInfo
        .Name = Name
        .hWnd = hWnd
        .uMsg = uMsg
        .Icon = Icon
        .LargeIcon = LargeIcon
        .Signature = Name
        ' /* V38.119 - for V39: can unregister an app using it's name and pid */
        .Pid = Pid

    End With

    If uAddByStruct(pInfo) <> 0 Then
        ' /* success (token was returned) */
        OldAdd = M_OK

    Else
        OldAdd = M_FAILED

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

'Private Sub uAdd(ByRef theApp As TApp)          ', ByVal IsNewApp As Boolean
'
'    uAlphaAdd theApp
'
'    ' /* check version - deprecated as of V41 (wasn't ever used?) */
'
''    If (theApp.hWnd <> 0) And (theApp.uMsg <> 0) Then
''        i = SendMessage(theApp.hWnd, theApp.uMsg, SNARL_ASK_APPLET_VER, ByVal 0&)
''        If i <> 0 Then
''            g_Debug "TApplicationRoster.uAdd(): app responded to SNARL_ASK_APPLET_VER with " & CStr(i), LEMON_LEVEL_INFO
''
''        Else
''            g_Debug "TApplicationRoster.uAdd(): app ignored SNARL_ASK_APPLET_VER"
''
''        End If
''
''    End If
'
'Dim bShow As Boolean
'
'    bShow = True
'
'    ' /* don't show if it's us or we're locked */
'
''Debug.Print "(): " & gPrefs.notify_on_first_register & " >> " & g_IgnoreLock
'
'    If (theApp.Name = App.Title) Or (g_IgnoreLock) > 0 Then
'        bShow = False
'
'    ElseIf g_ConfigGet("notify_on_first_register") = "1" Then
'        ' /* if notify_on_first_register is set, we need to take account of the IsNewApp setting */
'        bShow = theApp.IsNew()
'
'    End If
'
'Dim lFlags As SNRL_NOTIFICATION_FLAGS
'Dim i As Long
'
'    If bShow Then
'
'        ' /* if app has a password or was registered via some IP protocol
'        '    add the relevant system flags here */
'
'        If theApp.Password <> "" Then _
'            lFlags = lFlags Or SNRL_NOTIFICATION_SECURE
'
'        If theApp.IsRemote Then _
'            lFlags = lFlags Or SNRL_NOTIFICATION_REMOTE
'
'        ' /* create the notification */
'
'        i = g_PrivateNotify(SNARL_CLASS_APP_REG, _
'                            "Application registered", _
'                            theApp.Name & " registered successfully with Snarl", _
'                            -1, _
'                            IIf(theApp.LargeIcon = "", g_MakePath(App.Path) & "etc\icons\snarl.png", theApp.LargeIcon), , , _
'                            SNARL41_NOTIFICATION_AUTO_DISMISS, lFlags)
'
'        If i Then
'            ' /* add the action(s) */
'            sn42AddAction i, theApp.Name & " Settings...", "!snarl cfg " & theApp.Signature
'            If theApp.HasConfig Then _
'                sn42AddAction i, "Configure " & theApp.Name & "...", "!snarl manage " & theApp.Signature
'
''            szMenu = "Notification Settings...#?!snarl cfg " & theApp.Signature
''            If theApp.HasConfig Then _
'                szMenu = szMenu & "||Configure " & theApp.Name & "...#?!snarl manage " & theApp.Signature
'
''            snChangeAttribute i, SNARL_ATTRIBUTE_MENU, szMenu
'
'        End If
'
'    End If
'
'    frmAbout.bUpdateAppList
'
'End Sub

Private Sub uAlphaAdd(ByRef theApp As TApp)
Static i As Long
Static j As Long

    ' /* add it alpha-sorted */

    If mCount Then
        For i = 1 To mCount
            If LCase$(theApp.Name) < LCase$(mApp(i).Name) Then
                mCount = mCount + 1
                ReDim Preserve mApp(mCount)
                For j = mCount To (i + 1) Step -1
                    Set mApp(j) = mApp(j - 1)

                Next j

                Set mApp(i) = theApp
                Exit Sub

            End If
        Next i
    End If
    
    ' /* drop through here if no other apps or can be added to end of list */

    mCount = mCount + 1
    ReDim Preserve mApp(mCount)
    Set mApp(mCount) = theApp

End Sub

Public Function Register(ByVal Name As String, ByVal hWnd As Long, ByVal uMsg As Long, ByVal Icon As String, ByVal LargeIcon As String, ByVal Pid As Long) As M_RESULT

    ' /* for snRegisterApp() - to be introduced in V39 */
    g_Debug "TApplicationRoster.Register()", LEMON_LEVEL_PROC_ENTER
    g_Debug "Name='" & Name & "' pid=" & CStr(Pid) & " hWnd=0x" & g_HexStr(hWnd) & " ReplyMsg=0x" & g_HexStr(uMsg)

    If Pid = 0 Then
        Register = M_INVALID_ARGS
        g_Debug "pid can't be null", LEMON_LEVEL_CRITICAL

    ElseIf uFindByPid(Pid) <> 0 Then
        Register = M_ALREADY_EXISTS
        g_Debug "app with pid " & CStr(Pid) & " already registered", LEMON_LEVEL_CRITICAL

    Else
        ' /* carry on from here */
        Register = OldAdd(Name, hWnd, uMsg, Icon, LargeIcon, Pid)

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function OldRemove(ByVal hWnd As Long) As Boolean

    If mCount = 0 Then _
        Exit Function

Dim i As Long

    For i = 1 To mCount
        If mApp(i).hWnd = hWnd Then
            OldRemove = (Remove(i) = M_OK)
            Exit Function

        End If

    Next i

    g_Debug "TApplicationRoster.OldRemove(): app (window) " & g_HexStr(hWnd) & " not found", LEMON_LEVEL_WARNING

End Function

Public Function FindByWindow(ByVal hWnd As Long, ByRef App As TApp) As Boolean

    If (hWnd = 0) Or (mCount = 0) Then _
        Exit Function

Dim i As Long

    For i = 1 To mCount
        If mApp(i).hWnd = hWnd Then
            Set App = mApp(i)
            FindByWindow = True
            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindByWindow(): window " & g_HexStr(hWnd) & " not found", LEMON_LEVEL_CRITICAL

End Function

Public Function FindByToken(ByVal Token As Long, ByRef App As TApp, ByVal Password As String) As Boolean

    If (Token = 0) Or (mCount = 0) Then _
        Exit Function

Static i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            ' /* V42: validate password, if the app has one */
            If (mApp(i).Password <> "") And (mApp(i).Password <> Password) Then
                g_Debug "TApplicationRoster.FindByToken(): incorrect password", LEMON_LEVEL_CRITICAL
                gSetLastError SNARL_ERROR_AUTH_FAILURE

            Else
                gSetLastError SNARL_SUCCESS
                Set App = mApp(i)
                FindByToken = True

            End If

            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindByToken(): " & g_HexStr(Token) & " not found", LEMON_LEVEL_CRITICAL
    gSetLastError SNARL_ERROR_NOT_REGISTERED

End Function

Public Function FindBySignature(ByVal Signature As String, ByRef App As TApp, ByVal Password As String) As Boolean

    If (Signature = "") Or (mCount = 0) Then _
        Exit Function

Static i As Long

    For i = 1 To mCount
        If mApp(i).Signature = Signature Then
            ' /* V42: validate password, if the app has one */
            If (mApp(i).Password <> "") And (mApp(i).Password <> Password) Then
                g_Debug "TApplicationRoster.FindBySignature(): incorrect password", LEMON_LEVEL_CRITICAL
                gSetLastError SNARL_ERROR_AUTH_FAILURE

            Else
                gSetLastError SNARL_SUCCESS
                Set App = mApp(i)
                FindBySignature = True

            End If

            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindBySignature(): " & Signature & " not found", LEMON_LEVEL_CRITICAL
    gSetLastError SNARL_ERROR_NOT_REGISTERED

End Function

Public Function Find(ByVal Name As String, ByRef App As TApp) As Boolean
Dim i As Long

    i = MRoster_FindItem(Name)
    If i > 0 Then
        Set App = mApp(i)
        Find = True

    End If

End Function

Public Function AppAt(ByVal Index As Long) As TApp

    Set AppAt = MRoster_ItemAt(Index)

End Function

Public Sub SendToAll(ByVal EventId As SNARL_GLOBAL_EVENTS)
Static i As Long

    If mCount = 0 Then _
        Exit Sub

    g_Debug "TApplicationRoster.SendToAll(): sending event " & CStr(EventId) & " to all apps..."

    For i = mCount To 1 Step -1
        If (IsWindow(mApp(i).hWnd) <> 0) And (mApp(i).uMsg <> 0) Then _
            PostMessage mApp(i).hWnd, mApp(i).uMsg, EventId, ByVal 0&

    Next i

End Sub

Public Function CountApps() As Long

    CountApps = mCount

End Function

'Public Sub WriteConfig()
'Dim i As Long
'
'    If mCount Then
'        For i = 1 To mCount
'            mApp(i).UpdateConfig
'
'        Next i
'    End If
'
'End Sub

Public Sub ResetAlerts()
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            mApp(i).ResetAlerts

        Next i
    End If

End Sub

Public Function OldUnregister(ByVal Pid As Long) As M_RESULT

    If Pid = 0 Then
        g_Debug "TApplicationRoster.OldUnregister(): invalid arg", LEMON_LEVEL_CRITICAL
        OldUnregister = M_INVALID_ARGS
        Exit Function

    End If

    If mCount = 0 Then
        g_Debug "TApplicationRoster.OldUnregister(): no applications registered", LEMON_LEVEL_CRITICAL
        OldUnregister = M_FAILED
        Exit Function

    End If

Dim sz As String
Dim i As Long

    i = uFindByPid(Pid)
    If i Then
        sz = mApp(i).Name
        OldUnregister = Remove(i)
        g_Debug "TApplicationRoster.OldUnregister(): app " & CStr(Pid) & " (" & sz & ") unregistered okay"

    Else

        g_Debug "TApplicationRoster.OldUnregister(): app #" & CStr(Pid) & " not found", LEMON_LEVEL_CRITICAL
        OldUnregister = M_NOT_FOUND

    End If

End Function

Public Function Remove(ByVal Index As Long) As M_RESULT

    If (Index < 1) Or (Index > mCount) Then
        g_Debug "TApplicationRoster.Remove(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Remove = M_INVALID_ARGS
        Exit Function

    End If

Dim szIcon As String
Dim sz As String
Dim i As Long

    sz = mApp(Index).Name
    szIcon = IIf(mApp(Index).Icon = "", g_MakePath(App.Path) & "etc\icons\no_icon.png", mApp(Index).Icon)

    If Index < mCount Then
        For i = Index To (mCount - 1)
            Set mApp(i) = mApp(i + 1)

        Next i
    End If

    mCount = mCount - 1
    ReDim Preserve mApp(mCount)

    ' /* don't show if it's us unregistering ;-) */

    If (sz <> App.Title) And (g_IgnoreLock < 1) And (g_ConfigGet("notify_on_first_register") = "0") Then _
        g_PrivateNotify SNARL_CLASS_APP_UNREG, "Application unregistered", _
                     sz & " unregistered successfully with Snarl", , _
                     szIcon

    frmAbout.bUpdateAppList

    Remove = M_OK

End Function

Private Function uFindByPid(ByVal Pid As Long) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If (mApp(i).Pid = Pid) Then
                uFindByPid = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function FindByPid(ByVal Pid As Long, ByRef App As TApp) As Boolean
Dim i As Long

    i = uFindByPid(Pid)
    If i Then
        Set App = mApp(i)
        FindByPid = True

    End If

End Function

Public Function IndexOfPidAndName(ByVal Pid As Long, ByVal Name As String) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If (mApp(i).Pid = Pid) And (mApp(i).Name = Name) Then
                IndexOfPidAndName = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function IndexOf(ByVal Title As String) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If mApp(i).Name = Title Then
                IndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function IndexOfSig(ByVal Signature As String) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
'            MsgBox mApp(i).Signature & " / " & Signature
            If mApp(i).Signature = Signature Then
                IndexOfSig = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function UnregisterBySig(ByVal Signature As String, ByVal Password As String) As Long
Static i As Long

    i = IndexOfSig(Signature)
    If i = 0 Then
        g_Debug "TApplicationRoster.UnregisterBySig(): '" & Signature & "' not found", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_NOT_REGISTERED

    Else
        UnregisterBySig = uUnregister(i, Password)

    End If

End Function

Public Function Unregister(ByVal Token As Long, ByVal Password As String) As Long
Static i As Long

    i = uIndexOfToken(Token)
    If i = 0 Then
        g_Debug "TApplicationRoster.Unregister(): '" & CStr(Token) & "' not found", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_NOT_REGISTERED

    Else
        Unregister = uUnregister(i, Password)

    End If

End Function

Private Function uUnregister(ByVal Index As Long, ByVal Password As String) As Long

    ' /* WARNING: index is _NOT_ bounds-checked */

    ' /* V42: password must match */
    If mApp(Index).Password <> "" Then
        If mApp(Index).Password <> Password Then
            gSetLastError SNARL_ERROR_AUTH_FAILURE
            Exit Function

        End If
    End If

    Remove Index
    uUnregister = -1

End Function

Private Function uIndexOfToken(ByVal Token As Long) As Long

    If (mCount < 1) Or (Token = 0) Then _
        Exit Function

Static i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            uIndexOfToken = i
            Exit Function

        End If
    Next i

End Function

Public Function Update(ByVal Token As Long, ByRef Data As BPackedData) As Long

    If (Data Is Nothing) Then
        g_Debug "TApplicationRoster.Update(): arg missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

Static i As Long

    i = uIndexOfToken(Token)
    If i = 0 Then
        g_Debug "TApplicationRoster.Update(): '" & CStr(Token) & "' not found", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_NOT_REGISTERED
        Exit Function

    End If

    ' /* V42: password must match */

    If mApp(i).Password <> "" Then
        If mApp(i).Password <> Data.ValueOf("password") Then
            g_Debug "TApplicationRoster.Update(): '" & CStr(Token) & "' password mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_AUTH_FAILURE
            Exit Function

        End If
    End If

'        If Data.Exists("title") Then _
        mApp(i)Name = Data.ValueOf("title")

    If Data.Exists("icon") Then _
        mApp(i).SetIcon Data.ValueOf("icon")

    frmAbout.bUpdateAppList
    Update = -1

End Function

'Friend Sub bAddClass(ByRef Class As TAlert)
'
'    mClasses = mClasses + 1
'    ReDim Preserve mClass(mClasses)
'    Set mClass(mClasses) = Class
'
'End Sub
'
'Public Function FindClass(ByVal Token As Long, ByRef Class As TAlert) As Boolean
'
'    If mClasses = 0 Then _
'        Exit Function
'
'Dim i As Long
'
'    For i = 1 To mClasses
'        If mClass(i).Token = Token Then
'            Set Class = mClass(i)
'            FindClass = True
'            Exit Function
'
'        End If
'    Next i
'
'End Function

Public Function SnarlAppsMenu() As OMMenu
Dim pm As OMMenu
Dim i As Long

    Set SnarlAppsMenu = New OMMenu

    With SnarlAppsMenu
        If mCount Then
            For i = 1 To mCount
                If mApp(i).IsSnarlApp Then
'                    Set pm = New OMMenu
'                    pm.AddItem pm.CreateItem("cfg" & CStr(i), "Settings...", , (mApp(i).Flags And E_APP_HAS_PREFS))
'                    pm.AddItem pm.CreateItem("abt" & CStr(i), "About...", , (mApp(i).Flags And E_APP_HAS_ABOUT))
                    .AddItem .CreateItem("!" & CStr(i), mApp(i).Name, , (mApp(i).Flags And SNARLAPP_HAS_PREFS))

                End If
            Next i
        End If

        If .CountItems = 0 Then _
            .AddItem .CreateItem("", "None", , False)

    End With

End Function

Public Sub SnarlAppDo(ByVal Index As Long, ByVal Cmd As Long)

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

'    Debug.Print "SNARLAPPDO: " & Cmd & " > " & mApp(Index).hWnd

    PostMessage mApp(Index).hWnd, mSnarlAppMsg, Cmd, ByVal 0&

End Sub

Private Function uGenerateSnarlPassword() As String
Dim i As Integer
Dim c As Integer

    Randomize Timer
    c = (Rnd * 24) + 8      ' // variable length between 8 and 32 chars

    For i = 1 To c
        Randomize Timer
        uGenerateSnarlPassword = uGenerateSnarlPassword & g_HexStr(Rnd * 255, 2)

    Next i

End Function

'Public Function FindBySignature(ByVal Signature As String, ByRef App As TApp) As Boolean
'Dim i As Long
'
'    i = IndexOfSig(Signature)
'    If i Then
'        Set App = mApp(i)
'        FindBySignature = True
'
'    End If
'
'End Function

Private Sub uUpdate(ByVal Index As Long, ByRef Data As BPackedData)

    If (Index < 1) Or (Index > mCount) Or (Data Is Nothing) Then _
        Exit Sub

    ' /* can only change certain things */

Dim rm As Long
Dim rt As Long

    With mApp(Index)

        If Data.Exists("icon") Then _
            .SetIcon Data.ValueOf("icon")

        If Data.Exists("reply-to") Then
            rt = Val(Data.ValueOf("reply-to"))

        Else
            rt = Val(Data.ValueOf("hwnd"))

        End If

        If Data.Exists("reply") Then
            ' /* V42.22: new name */
            rm = Val(Data.ValueOf("reply"))

        Else
            rm = Val(Data.ValueOf("umsg"))

        End If

        .SetReply rt, rm

    End With

End Sub

