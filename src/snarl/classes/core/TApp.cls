VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TApp.cls
'/
'/  Description:    Represents a single registered application
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Dim mInfo As T_SNARL_APP        ' // app info
Dim mAlert() As TAlert          ' // class(es) registered against this app
Dim mAlerts As Long

Dim mSettings As ConfigFile     ' // persistent settings
Dim mAppConfig As ConfigSection ' // persistent app settings
Dim mIsUnknown As Boolean       ' // new app (or, one without persistent settings anyway)

Implements MObject

Private Sub Class_Initialize()

    mInfo.Name = "Snarl"

End Sub

Private Property Get MObject_Type() As String
End Property

Friend Function bInit(ByRef Info As T_SNARL_APP) As Boolean
Dim sz As String
Dim i As Long
Dim f As Boolean

    g_Debug "TApp.bInit()", LEMON_LEVEL_PROC_ENTER

    ' /* store registration time/date */

    Info.Timestamp = Now()

    LSet mInfo = Info
    mIsUnknown = False

    ' /* do we have a config file for this app? */

    Set mSettings = New ConfigFile
    With mSettings
        ' /* R2.31: we know that SnarlConfigPath is valid */
        .File = gPrefs.SnarlConfigPath & "etc\" & g_MakeFilename(Info.Signature) & ".class"
        .Load

        ' /* V40.25: must have a "_cfg" section */

        i = .FindSection("_cfg")
        If i = 0 Then
            Set mAppConfig = .AddSectionObj("_cfg")
            uWriteSettings

        Else
            Set mAppConfig = .SectionAt(i)

        End If

        ' /* must have an "_all" section */

        mIsUnknown = Not (.SectionExists("_all"))

        If mIsUnknown Then
            .AddSection "_all"
            uWriteSettings

        End If

    End With

    ' /* all apps get this class */

    AddAlert "_all", ""
    bInit = True

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function AddAlert(ByVal Class As String, ByVal Description As String, Optional ByRef ClassObject As TAlert, Optional ByVal Flags As E_CLASS_FLAGS) As M_RESULT

    g_Debug "TApp.AddAlert()", LEMON_LEVEL_PROC_ENTER

    ' /* check the class isn't already registered */

    If Me.FindAlert(Class, Nothing) Then
        g_Debug "class '" & Class & "' already exists", LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        gSetLastError SNARL_ERROR_CLASS_ALREADY_EXISTS
        AddAlert = M_ALREADY_EXISTS
        Exit Function

    End If

Dim bNeedUpdate As Boolean
Dim pClassSettings As ConfigSection
Dim i As Long

    ' /* load up persistent settings if any exist */

    g_Debug "checking persistent settings..."

    With mSettings
        i = .FindSection(Class)
        If i = 0 Then
            g_Debug "no config for '" & Class & "', creating section now..."
            Set pClassSettings = .AddSectionObj(Class)
            uWriteSettings

        Else
            g_Debug "got settings for '" & Class & "'"
            Set pClassSettings = .SectionAt(i)

        End If

    End With

    ' /* add and intialize it */

    g_Debug "creating class object..."
    Set ClassObject = New TAlert
    ClassObject.bInit Class, Description, Me, pClassSettings, Flags

    mAlerts = mAlerts + 1
    ReDim Preserve mAlert(mAlerts)
    Set mAlert(mAlerts) = ClassObject

    AddAlert = M_OK

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function IndexOf(ByVal Class As String) As Long

    If (mAlerts = 0) Or (Class = "") Then _
        Exit Function

Dim i As Long

    For i = 1 To mAlerts
        If mAlert(i).Name = Class Then
            IndexOf = i
            Exit Function

        End If
    Next i

End Function

Public Function FindAlert(ByVal Class As String, ByRef Alert As TAlert) As Boolean
Dim i As Long

    i = IndexOf(Class)
    If i Then
        Set Alert = mAlert(i)
        Alert.Index = i
        FindAlert = True

    End If

End Function

Public Function CountAlerts() As Long

    CountAlerts = mAlerts

End Function

Public Function AlertAt(ByVal Index As Long) As TAlert

    If (Index > 0) And (Index <= mAlerts) Then
        Set AlertAt = mAlert(Index)
        AlertAt.Index = Index

    End If

End Function

Public Sub WriteConfig()

    uWriteSettings

End Sub

Public Sub ResetAlerts()
Dim i As Long

    If mAlerts Then
        For i = 1 To mAlerts
            mAlert(i).Reset

        Next i

    End If

End Sub

Public Function Name() As String

    Name = mInfo.Name

End Function

Public Function hWnd() As Long

    hWnd = mInfo.hWnd

End Function

Public Function uMsg() As Long

    uMsg = mInfo.uMsg

End Function

Public Function Pid() As Long

    Pid = mInfo.Pid

End Function

Public Function Icon() As String

    Icon = mInfo.Icon

End Function

'Public Function LargeIcon() As String
'
'    LargeIcon = mInfo.LargeIcon
'
'End Function

Public Function UseSameSettingsForAllClasses() As Boolean

    If Not (mAppConfig Is Nothing) Then _
        UseSameSettingsForAllClasses = Val(mAppConfig.GetValueWithDefault("use_same_settings_for_all_classes", "0"))

End Function

Public Sub SetUseSameSettingsForAllClasses(ByVal Enabled As Boolean)

    If Not (mAppConfig Is Nothing) Then
        mAppConfig.Update "use_same_settings_for_all_classes", IIf(Enabled, "1", "0")
        If Not (mSettings Is Nothing) Then _
            uWriteSettings

    End If

End Sub

Public Function Show(ByVal Class As String, Optional ByVal Title As String, Optional ByVal Text As String, Optional ByVal Timeout As Long, Optional ByVal Icon As String, Optional ByVal hWndReply As Long, Optional ByVal uMsg As Long, Optional ByVal SoundFile As String, Optional ByVal Style As String, Optional ByVal Ack As String) As M_RESULT
Dim pc As TAlert

    ' /* V39.86: wasn't taking account of the "Use Same Settings..." option */

    If Me.UseSameSettingsForAllClasses Then
        Set pc = mAlert(1)

    ElseIf Not Me.FindAlert(Class, pc) Then
        ' /* is the class registered? */
        g_Debug "TApp.Show(): class '" & Class & "' not registered, using default...", LEMON_LEVEL_WARNING
        If mAlerts > 0 Then _
            Set pc = mAlert(1)              ' // default/catch-all...

    End If

    ' /* here, we must have a class... */

    If (pc Is Nothing) Then
        g_Debug "TApp.Show(): failed - no class", LEMON_LEVEL_CRITICAL
        Show = M_FAILED
        Exit Function

    End If

Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        ' /* alert class won't know about these two */

        .hWndReply = hWndReply
        If .hWndReply = 0 Then _
            .hWndReply = mInfo.hWnd

        .uReplyMsg = uMsg
        If .uReplyMsg = 0 Then _
            .uReplyMsg = mInfo.uMsg

        .Pid = mInfo.Pid
        .Title = Title
        .Text = Text
        .Timeout = Timeout
        .IconPath = Icon
        .SndFile = SoundFile
        .StyleName = style_GetStyleName(Style)
        .SchemeName = style_GetSchemeName(Style)
        .DefaultAck = Ack

    End With

    ' /* pass it to the class... */

    Show = pc.ShowNotification(pInfo, Nothing)
    If Show = 0 Then _
        Show = M_FAILED

'    Show = pc.ShowNotification(mInfo.pid, Title, Text, Timeout, Icon, hWndReply, uMsg, SoundFile, Style, Ack)

End Function

Public Function Show41(ByVal ClassId As String, ByRef pData As BPackedData, Optional ByRef ReplySocket As CSocket, Optional ByVal IntFlags As E_NOTIFICATION_FLAGS, Optional ByVal RemoteHost As String) As Long

    ' /* this is where we transition from a BPackedData into a T_NOTIFICATION_INFO struct */

    If (pData Is Nothing) Then _
        Exit Function

Dim pClass As TAlert

    ' /* V39.86: wasn't taking account of the "Use Same Settings..." option */

    If Me.UseSameSettingsForAllClasses Then
        ' /* use the All/Other class definitions */
        Set pClass = mAlert(1)

    ElseIf Not Me.FindAlert(ClassId, pClass) Then
        ' /* is the class registered? */
        g_Debug "TApp.Show(): class '" & ClassId & "' not registered, using default...", LEMON_LEVEL_WARNING
        If mAlerts > 0 Then _
            Set pClass = mAlert(1)              ' // default/catch-all...

        ' /* to-do: block notifications with invalid class names, especially those
        '    sourcing from remotely registered apps */

    End If

    ' /* here, we must have a class... */

    If (pClass Is Nothing) Then
        g_Debug "TApp.Show(): couldn't find a suitable class", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_SYSTEM
        Exit Function

    End If

    ' /* convert the packed data into a T_NOTIFICATION_INFO struct */

Dim pInfo As T_NOTIFICATION_INFO
Dim i As Long

    With pInfo

        ' /* R2.4 DR7: "callback" is now preferred over "ack" */

        If pData.Exists("callback") Then
            .DefaultAck = pData.ValueOf("callback")

        Else
            .DefaultAck = pData.ValueOf("ack")

        End If

        ' /* R2.4 DR7: new "uid" arg */

        If pData.Exists("uid") Then _
            .CustomUID = pData.ValueOf("uid")

        ' /* R2.4 DR7: new "value" arg */

        .Value = pData.ValueOf("value")

        ' /* R2.4 DR7: internal-only args (can only be set at g_DoNotify() level) */

        .IntFlags = IntFlags

        ' /* R2.4 DR7: static copy of sender */

        .RemoteHost = RemoteHost

        ' /* flags: deprecated */

        If pData.Exists("flags") Then
            i = Val("&H" & pData.ValueOf("flags"))
            .Flags = (i And &HFFFF&)                ' // only keep user flags

        End If

        ' /* R2.4 DR7: "reply-to" is now preferred over "hwnd" */

        .hWndReply = Val(pData.ValueOf("reply-to"))
        If .hWndReply = 0 Then _
            .hWndReply = Val(pData.ValueOf("hwnd"))

        If .hWndReply = 0 Then _
            .hWndReply = mInfo.hWnd                 ' // if still no reply window, use app-defined value (which may still be 0)

        .IconPath = pData.ValueOf("icon")

        .Priority = Val(pData.ValueOf("priority"))
        Set .Socket = ReplySocket

        .Text = Replace$(pData.ValueOf("text"), "\n", vbCrLf)

        .Timeout = -1                               ' // pre-set with default
        If pData.Exists("timeout") Then _
            .Timeout = Val(pData.ValueOf("timeout"))

        .Title = Replace$(pData.ValueOf("title"), "\n", vbCrLf)

        ' /* R2.4 DR7: "reply" is now preferred over "umsg" */

        .uReplyMsg = Val(pData.ValueOf("reply"))
        If .uReplyMsg = 0 Then _
            .uReplyMsg = Val(pData.ValueOf("umsg"))

        If .uReplyMsg = 0 Then _
            .uReplyMsg = mInfo.uMsg                 ' // if still no reply message, use app-defined value (which may still be 0)



        ' /* R2.4.2: caller can now request a particular style */

        If pData.Exists("style") Then
            .StyleName = style_GetStyleName(pData.ValueOf("style"))
            .SchemeName = style_GetSchemeName(pData.ValueOf("style"))

        End If


        ' /* V42 styles need the unabridged content */

        .OriginalContent = Replace$(pData.AsString(), "\n", vbCrLf)

    End With

    ' /* pass it to the class... */

    Show41 = pClass.ShowNotification(pInfo, pData)

End Function


'Friend Function ShowNew(ByVal Class As String, ByRef Info As T_NOTIFICATION_INFO) As Long
'Dim pc As TAlert
'
'    ' /* V39.86: wasn't taking account of the "Use Same Settings..." option */
'
'    If Me.UseSameSettingsForAllClasses Then
'        Set pc = mAlert(1)
'
'    ElseIf Not Me.FindAlert(Class, pc) Then
'        ' /* is the class registered? */
'        g_Debug "TApp.Show(): class '" & Class & "' not registered, using default...", LEMON_LEVEL_WARNING
'        If mAlerts > 0 Then _
'            Set pc = mAlert(1)              ' // default/catch-all...
'
'    End If
'
'    ' /* here, we must have a class... */
'
'    If (pc Is Nothing) Then
'        g_Debug "TApp.Show(): couldn't find a suitable class", LEMON_LEVEL_CRITICAL
'        gSetLastError SNARL_ERROR_FAILED
'        Exit Function
'
'    End If
'
'    ' /* alert class won't know about these two */
'
'    If Info.hWndReply = 0 Then _
'        Info.hWndReply = mInfo.hWnd
'
'    If Info.uReplyMsg = 0 Then _
'        Info.uReplyMsg = mInfo.uMsg
'
'    ' /* secure */
'
''    If mInfo.Password <> "" Then _
'        Info.Flags = Info.Flags Or SNARL42_NOTIFICATION_SECURE
'
'    ' /* pass it to the class... */
'
'    ShowNew = pc.ShowNotification(Info)
'
'End Function

Public Function Token() As Long

    Token = mInfo.Token

End Function

Public Sub SetIcon(ByVal Path As String)

    mInfo.Icon = Path
    mInfo.LargeIcon = Path

End Sub

Public Function AddClass(ByRef Data As BPackedData, Optional ByRef ClassObject As TAlert) As Long

'//ByVal Name As String, ByVal Flags As E_ALERT_CLASS_FLAGS, Optional ByRef ClassObject As TAlert

    ' /* validate some key stuff */

    If (Data Is Nothing) Then
        g_Debug "TApp.AddClass(): config error", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_FAILED
        Exit Function

    End If

    If Not Data.Exists("id") Then
        g_Debug "TApp.AddClass(): required arg 'id' missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    If Not Data.Exists("name") Then
        g_Debug "TApp.AddClass(): required arg 'name' missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* enabled/disabled (backwards compatability with pre-release editions) */

Dim lFlags As E_CLASS_FLAGS

    If Data.Exists("enabled") Then
        If Data.ValueOf("enabled") = "0" Then _
            lFlags = lFlags Or E_CLASS_DISABLED

    End If

    ' /* Use AddAlert() to actually add the class to our app - LastError will be set on failure */

    If AddAlert(Data.ValueOf("id"), Data.ValueOf("name"), ClassObject, lFlags) <> M_OK Then _
        Exit Function

    ' /* do some V41 post-initialization */

    With ClassObject
        If Data.Exists("callback") Then
            ' /* R2.4 DR7: name change */
            .DefaultAck = Data.ValueOf("callback")

        Else
            .DefaultAck = Data.ValueOf("ack")

        End If

'        Debug.Print .Description & " >> " & Data.Exists("icon")

        .DefaultIcon = Data.ValueOf("icon")

        ' /* R2.4 Beta 4: check for MIME encoded icon */

        If (.DefaultIcon = "") And (Data.Exists("icon-base64")) Then _
            .DefaultIcon = g_GetBase64Icon(Data.ValueOf("icon-base64"))

        .DefaultSound = Data.ValueOf("sound")
        .DefaultText = Data.ValueOf("text")
        .DefaultTimeout = Val(Data.ValueOf("duration"))
        .DefaultTitle = Data.ValueOf("title")

    End With

    frmAbout.bUpdateClassList mInfo.Token
    AddClass = -1

End Function

Public Function RemClass(ByRef Data As BPackedData, Optional ByVal RemoveAll As Boolean = False) As Long
Static i As Long
Static j As Long

    ' /* validate some key stuff */

    If (Data Is Nothing) Then
        g_Debug "TApp.RemClass(): config error", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_FAILED
        Exit Function

    End If

    ' /* special case: if the 'all' tag is present and it's set to "1" then
    '    remove all classes */

    If (Data.ValueOf("all") = "1") Or (RemoveAll) Then
        g_Debug "TApp.RemClass(): removing all classes..."
        mAlerts = 0
        ReDim mAlert(0)
        AddAlert "_all", ""
        RemClass = -1
        frmAbout.bUpdateClassList mInfo.Token
        gSetLastError SNARL_SUCCESS

    ElseIf Not Data.Exists("id") Then
        g_Debug "TApp.RemClass(): required arg 'id' missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING

    ElseIf Data.ValueOf("id") = "_all" Then
        g_Debug "TApp.RemClass(): cannot remove '_all' class", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_FAILED

    Else
        i = uIndexOf(Data.ValueOf("id"))
        If i = 0 Then
            g_Debug "TApp.RemClass(): class '" & Data.ValueOf("id") & "' not found", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_CLASS_NOT_FOUND

        Else

            If i < mAlerts Then
                For j = i To (mAlerts - 1)
                    Set mAlert(j) = mAlert(j + 1)

                Next j

            End If

            mAlerts = mAlerts - 1
            ReDim Preserve mAlert(mAlerts)
            frmAbout.bUpdateClassList mInfo.Token
            RemClass = -1
            gSetLastError SNARL_SUCCESS

        End If
    End If

End Function

Public Function Signature() As String

    Signature = mInfo.Signature

End Function

Public Function Flags() As SNARLAPP_FLAGS

    Flags = mInfo.Flags

End Function

Public Function IsSnarlApp() As Boolean

    If IsWindow(mInfo.hWnd) <> 0 Then _
        IsSnarlApp = (GetProp(mInfo.hWnd, "snarl_app") = 1) Or ((mInfo.Flags And SNARLAPP_IS_WINDOWLESS) <> 0)

End Function

Public Function HasConfig() As Boolean

    If Not IsSnarlApp() Then _
        Exit Function

    If IsWindow(mInfo.hWnd) <> 0 Then _
        HasConfig = ((mInfo.Flags And SNARLAPP_HAS_PREFS) <> 0)

End Function

Public Sub DoAbout()

    If (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARLAPP_HAS_ABOUT)) Then
        ' /* ask the app to display its own details */
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_DO_ABOUT, ByVal 0&

    Else
        ' /* just show what we can */
        g_PrivateNotify "", _
                        mInfo.Name, _
                        "Registered: " & CStr(mInfo.Timestamp) & _
                        IIf(mInfo.Password <> "", "\nPassword protected", "") & _
                        IIf(mInfo.IsRemote, "\nSource is remote", ""), _
                        , _
                        mInfo.Icon

    End If

End Sub

Public Sub DoSettings()

    If (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARLAPP_HAS_PREFS)) Then _
        PostMessage mInfo.hWnd, snAppMsg(), SNARLAPP_DO_PREFS, ByVal 0&

End Sub

Private Function uIndexOf(ByVal Id As String) As Long

    If mAlerts = 0 Then _
        Exit Function

Dim i As Long

    For i = 1 To mAlerts
        If mAlert(i).Name = Id Then
            uIndexOf = i
            Exit Function

        End If
    Next i

End Function

Private Sub uWriteSettings()

    If (mSettings Is Nothing) Then _
        Exit Sub

    If gSysAdmin.TreatSettingsAsReadOnly Then
        g_Debug "TApp.uWriteSettings(): blocked by admin setting"

    Else
        mSettings.Save

    End If

End Sub

Public Function Password() As String

    Password = mInfo.Password

End Function

Public Function IsRemote() As Boolean

    IsRemote = mInfo.IsRemote

End Function

Public Function IsNew() As Boolean

    IsNew = mIsUnknown

End Function

'Public Sub SetLastError(ByVal err As Long)
'
'    mLastError = err
'
'End Sub
'
'Public Function LastError() As Long
'
'    LastError = mLastError
'
'End Function

Public Sub SetReply(ByVal ReplyTo As Long, ByVal ReplyMessage As Long)

    If ReplyTo <> 0 Then _
        mInfo.hWnd = ReplyTo

    If ReplyMessage <> 0 Then _
        mInfo.uMsg = ReplyMessage

End Sub
