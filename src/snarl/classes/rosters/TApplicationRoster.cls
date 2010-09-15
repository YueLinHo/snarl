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

        SendAll SNARL_QUIT

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

    g_Debug "TApplicationRoster.Open()", LEMON_LEVEL_PROC

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        mSnarlAppMsg = RegisterWindowMessage("SnarlAppMessage")     ' // V40.25
        mAppToken = &H2C0

        ' /* register ourselves with, er, ourself! */
    
        g_IsRunning = True
'        snRegisterConfig2 frmAbout.hWnd, App.Title, &H555, g_MakePath(App.Path) & "etc\icons\snarl.png"

        g_Debug "TApplicationRoster.Open(): registering ourself..."
        snRegisterApp App.Title, g_MakePath(App.Path) & "etc\icons\snarl.png", g_MakePath(App.Path) & "etc\icons\snarl.png", frmAbout.hWnd, &H555

        g_Debug "TApplicationRoster.Open(): registering classes..."
        snAddClass SNARL_CLASS_WELCOME
        snAddClass SNARL_CLASS_APP_REG
        snAddClass SNARL_CLASS_APP_UNREG
        ' /* V38.97 */
        snAddClass SNARL_CLASS_JSON
        ' /* V38.142 */
        snAddClass SNARL_CLASS_ANON_NET
        ' /* V39.21 */
        snAddClass SNARL_CLASS_ANON
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

Public Function Add41(ByRef Data As BPackedData) As Long
Dim i As Long

    g_Debug "TApplicationRoster.Add41()", LEMON_LEVEL_PROC_ENTER

    ' /* validate args */

    If (Data.ValueOf("title") = "") Or (Data.ValueOf("id") = "") Then
        g_Debug "required arg missing", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL41_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* check for duplicate and return existing token */

Dim sz As String

    If mCount Then
        sz = Data.ValueOf("id")
        For i = 1 To mCount
            If mApp(i).Signature = sz Then
                g_Debug "'" & sz & "' already registered (token=" & mApp(i).Token & ")", LEMON_LEVEL_PROC_EXIT
                Add41 = mApp(i).Token
                Exit Function

            End If
        Next i
    End If

Dim pInfo As T_SNARL_APP

    ' /* configure the item */

    With pInfo
        .Name = Data.ValueOf("title")
        .Icon = Data.ValueOf("icon")
        .Signature = Data.ValueOf("id")
        .LargeIcon = .Icon                       ' // no need for two icons: V41 spec recommends 128x128px icon

        .hWnd = Val(Data.ValueOf("hwnd"))
        .uMsg = Val(Data.ValueOf("umsg"))

        .Flags = Val(Data.ValueOf("flags"))

        g_Debug "name='" & .Name & "' signature='" & .Signature & "' flags=0x" & g_HexStr(.Flags)

    End With

Dim bIsNewApp As Boolean
Dim pa As TApp

    Set pa = New TApp
    If pa.bInit(pInfo, bIsNewApp, mAppToken) Then
        uAdd pa, bIsNewApp
        Add41 = pa.Token

    Else
        Add41 = 0

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function OldAdd(ByVal Name As String, ByVal hWnd As Long, ByVal uMsg As Long, ByVal Icon As String, ByVal LargeIcon As String, ByVal pid As Long) As M_RESULT

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
        .pid = pid

    End With

Dim bIsNewApp As Boolean

    Set pa = New TApp
    If pa.bInit(pInfo, bIsNewApp, mAppToken) Then
        uAdd pa, bIsNewApp
        OldAdd = M_OK

    Else
        OldAdd = M_FAILED

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub uAdd(ByRef theApp As TApp, ByVal IsNewApp As Boolean)

    uAlphaAdd theApp

    ' /* check version - deprecated as of V41 (wasn't ever used?) */

'    If (theApp.hWnd <> 0) And (theApp.uMsg <> 0) Then
'        i = SendMessage(theApp.hWnd, theApp.uMsg, SNARL_ASK_APPLET_VER, ByVal 0&)
'        If i <> 0 Then
'            g_Debug "TApplicationRoster.uAdd(): app responded to SNARL_ASK_APPLET_VER with " & CStr(i), LEMON_LEVEL_INFO
'
'        Else
'            g_Debug "TApplicationRoster.uAdd(): app ignored SNARL_ASK_APPLET_VER"
'
'        End If
'
'    End If

Dim bShow As Boolean

    bShow = True

    ' /* don't show if it's us or we're locked */

'Debug.Print "(): " & gPrefs.notify_on_first_register & " >> " & g_IgnoreLock

    If (theApp.Name = App.Title) Or (g_IgnoreLock) > 0 Then
        bShow = False

    ElseIf g_ConfigGet("notify_on_first_register") = "1" Then
        ' /* if notify_on_first_register is set, we need to take account of the IsNewApp setting */
        bShow = IsNewApp

    End If

Debug.Print ">> " & bShow

    If bShow Then _
        snShowNotification SNARL_CLASS_APP_REG, "Application Registered", _
                            theApp.Name & " registered successfully with Snarl", 10, _
                            IIf(theApp.LargeIcon = "", g_MakePath(App.Path) & "etc\icons\snarl.png", theApp.LargeIcon)


    frmAbout.bUpdateAppList

End Sub

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

Public Function Register(ByVal Name As String, ByVal hWnd As Long, ByVal uMsg As Long, ByVal Icon As String, ByVal LargeIcon As String, ByVal pid As Long) As M_RESULT

    ' /* for snRegisterApp() - to be introduced in V39 */
    g_Debug "TApplicationRoster.Register()", LEMON_LEVEL_PROC_ENTER
    g_Debug "Name='" & Name & "' pid=" & CStr(pid) & " hWnd=0x" & g_HexStr(hWnd) & " ReplyMsg=0x" & g_HexStr(uMsg)

    If pid = 0 Then
        Register = M_INVALID_ARGS
        g_Debug "pid can't be null", LEMON_LEVEL_CRITICAL

    ElseIf uFindByPid(pid) <> 0 Then
        Register = M_ALREADY_EXISTS
        g_Debug "app with pid " & CStr(pid) & " already registered", LEMON_LEVEL_CRITICAL

    Else
        ' /* carry on from here */
        Register = OldAdd(Name, hWnd, uMsg, Icon, LargeIcon, pid)

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

Public Function FindByToken(ByVal Token As Long, ByRef App As TApp) As Boolean

    If (Token = 0) Or (mCount = 0) Then _
        Exit Function

Dim i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            Set App = mApp(i)
            FindByToken = True
            Exit Function

        End If
    Next i

    g_Debug "TApplicationRoster.FindByToken(): " & g_HexStr(Token) & " not found", LEMON_LEVEL_CRITICAL

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

Public Sub SendAll(ByVal Id As Long)
Dim i As Long

    If mCount Then
        For i = mCount To 1 Step -1
            If (IsWindow(mApp(i).hWnd) <> 0) And (mApp(i).uMsg <> 0) Then _
                SendMessage mApp(i).hWnd, mApp(i).uMsg, Id, ByVal 0&

        Next i
    End If

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

Public Function OldUnregister(ByVal pid As Long) As M_RESULT

    If pid = 0 Then
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

    i = uFindByPid(pid)
    If i Then
        sz = mApp(i).Name
        OldUnregister = Remove(i)
        g_Debug "TApplicationRoster.OldUnregister(): app " & CStr(pid) & " (" & sz & ") unregistered okay"

    Else

        g_Debug "TApplicationRoster.OldUnregister(): app #" & CStr(pid) & " not found", LEMON_LEVEL_CRITICAL
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
    szIcon = IIf(mApp(Index).LargeIcon = "", g_MakePath(App.Path) & "etc\icons\snarl.png", mApp(Index).LargeIcon)

    If Index < mCount Then
        For i = Index To (mCount - 1)
            Set mApp(i) = mApp(i + 1)

        Next i
    End If

    mCount = mCount - 1
    ReDim Preserve mApp(mCount)

    ' /* don't show if it's us unregistering ;-) */

    If (sz <> App.Title) And (g_IgnoreLock < 1) And (g_ConfigGet("notify_on_first_register") = "0") Then _
        snShowNotification SNARL_CLASS_APP_UNREG, "Application Unregistered", _
                        sz & " unregistered successfully with Snarl", 10, _
                        szIcon

    frmAbout.bUpdateAppList

    Remove = M_OK

End Function

Private Function uFindByPid(ByVal pid As Long) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If (mApp(i).pid = pid) Then
                uFindByPid = i
                Exit Function

            End If
        Next i
    End If

End Function

Public Function FindByPid(ByVal pid As Long, ByRef App As TApp) As Boolean
Dim i As Long

    i = uFindByPid(pid)
    If i Then
        Set App = mApp(i)
        FindByPid = True

    End If

End Function

Public Function IndexOfPidAndName(ByVal pid As Long, ByVal Name As String) As Long
Dim i As Long

    If mCount Then
        For i = 1 To mCount
            If (mApp(i).pid = pid) And (mApp(i).Name = Name) Then
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

Public Function Unregister(ByVal Token As Long) As Long
Static i As Long

    i = IndexOfToken(Token)
    If i = 0 Then
        g_Debug "TApplicationRoster.Unregister(): '" & CStr(Token) & "' not found", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_NOT_REGISTERED

    Else
        Remove i
        Unregister = -1

    End If

End Function

Private Function IndexOfToken(ByVal Token As Long) As Long

    If (mCount < 1) Or (Token = 0) Then _
        Exit Function

Static i As Long

    For i = 1 To mCount
        If mApp(i).Token = Token Then
            IndexOfToken = i
            Exit Function

        End If
    Next i

End Function

Public Function Update(ByVal Token As Long, ByRef Data As BPackedData) As Long
Static i As Long

    i = IndexOfToken(Token)
    If i = 0 Then
        g_Debug "TApplicationRoster.Update(): '" & CStr(Token) & "' not found", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_NOT_REGISTERED

    Else

'        If Data.Exists("title") Then _
            mApp(i)Name = Data.ValueOf("title")

        If Data.Exists("icon") Then _
            mApp(i).SetIcon Data.ValueOf("icon")

        frmAbout.bUpdateAppList
        Update = -1

    End If

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
                    .AddItem .CreateItem("!" & CStr(i), mApp(i).Name, , (mApp(i).Flags And SNARL41_APP_HAS_PREFS))

                End If
            Next i
        End If

        If .CountItems = 0 Then _
            .AddItem .CreateItem("", "None", , False)

    End With

End Function

Public Sub SnarlAppDo(ByVal Index As Long, ByVal Cmd As SNARL_APP_COMMANDS)

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

'    Debug.Print "SNARLAPPDO: " & Cmd & " > " & mApp(Index).hWnd

    PostMessage mApp(Index).hWnd, mSnarlAppMsg, Cmd, ByVal 0&

End Sub
