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

Implements MObject

Private Sub Class_Initialize()

    mInfo.Name = "Snarl"

End Sub

Private Property Get MObject_Type() As String
End Property

Friend Function bInit(ByRef Info As T_SNARL_APP, ByRef IsNew As Boolean, ByRef AppToken As Long) As Boolean
Dim sz As String
Dim i As Long
Dim f As Boolean

    g_Debug "TApp.bInit()", LEMON_LEVEL_PROC_ENTER

    LSet mInfo = Info
    IsNew = False

    mInfo.Token = AppToken
    AppToken = AppToken + 4

    ' /* do we have a config file for this app? */

    Set mSettings = New ConfigFile
    With mSettings
        If g_GetUserFolderPath(sz) Then
            ' /* TO DO: we should probably check the resulting path for validity here */
            .File = g_MakePath(sz) & "etc\" & g_MakeFilename(Info.Signature) & ".class"
            .Load

        Else
            g_Debug "couldn't find Snarl user folder - settings will not be saved", LEMON_LEVEL_WARNING

        End If

        ' /* V40.25: must have a "_cfg" section */

        i = .FindSection("_cfg")
        If i = 0 Then
            Set mAppConfig = .AddSectionObj("_cfg")
            .Save

        Else
            Set mAppConfig = .SectionAt(i)

        End If

        ' /* must have an "_all" section */

        IsNew = Not (.SectionExists("_all"))

        If IsNew Then
            .AddSection "_all"
            .Save

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
        gSetLastError SNARL41_ERROR_CLASS_ALREADY_EXISTS
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
            .Save

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

Public Function FindAlert(ByVal Class As String, ByRef Alert As TAlert) As Boolean
Dim i As Long

    If mAlerts Then
        For i = 1 To mAlerts
            If mAlert(i).Name = Class Then
                Set Alert = mAlert(i)
                Alert.Index = i
                FindAlert = True
                Exit Function

            End If
        Next i

    Else
        g_Debug "TApp.FindAlert(): '" & Name & "' has no alerts"

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

'    Debug.Print mSettings.File

    If Not (mSettings Is Nothing) Then _
        mSettings.Save

'Dim i As Long
'
'    If Not (mConfig Is Nothing) Then _
'        mConfig.Update "_use_same_settings_for_all_classes", IIf(mInfo.UseSameSettingsForAllClasses, "1", "0")
'
'    If mAlerts Then
'        For i = 1 To mAlerts
'            mAlert(i).UpdateConfig
'
'        Next i
'
'    End If

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

Public Function pid() As Long

    pid = mInfo.pid

End Function

Public Function Icon() As String

    Icon = mInfo.Icon

End Function

Public Function LargeIcon() As String

    LargeIcon = mInfo.LargeIcon

End Function

Public Function UseSameSettingsForAllClasses() As Boolean

    If Not (mAppConfig Is Nothing) Then _
        UseSameSettingsForAllClasses = Val(mAppConfig.GetValueWithDefault("use_same_settings_for_all_classes", "0"))

End Function

Public Sub SetUseSameSettingsForAllClasses(ByVal Enabled As Boolean)

    If Not (mAppConfig Is Nothing) Then
        mAppConfig.Update "use_same_settings_for_all_classes", IIf(Enabled, "1", "0")
        If Not (mSettings Is Nothing) Then _
            mSettings.Save

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

        .pid = mInfo.pid
        .Title = Title
        .Text = Text
        .Timeout = Timeout
        .IconPath = Icon
        .SndFile = SoundFile
        .StyleToUse = Style
        .DefaultAck = Ack

    End With

    ' /* pass it to the class... */

    Show = pc.ShowNotification(pInfo)
    If Show = 0 Then _
        Show = M_FAILED

'    Show = pc.ShowNotification(mInfo.pid, Title, Text, Timeout, Icon, hWndReply, uMsg, SoundFile, Style, Ack)

End Function

Friend Function ShowNew(ByVal Class As String, ByRef Info As T_NOTIFICATION_INFO) As Long
Dim pc As TAlert

    ' /* V39.86: wasn't taking about of the "Use Same Settings..." option */

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
        g_Debug "TApp.Show(): couldn't find a suitable class", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_FAILED
        Exit Function

    End If

    ' /* alert class won't know about these two */

    If Info.hWndReply = 0 Then _
        Info.hWndReply = mInfo.hWnd

    If Info.uReplyMsg = 0 Then _
        Info.uReplyMsg = mInfo.uMsg

    ' /* pass it to the class... */

    ShowNew = pc.ShowNotification(Info)

End Function

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
        gSetLastError SNARL41_ERROR_FAILED
        Exit Function

    End If

    If Not Data.Exists("id") Then
        g_Debug "TApp.AddClass(): required arg 'id' missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_ARG_MISSING
        Exit Function

    End If

    If Not Data.Exists("name") Then
        g_Debug "TApp.AddClass(): required arg 'name' missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_ARG_MISSING
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
        .DefaultAck = Data.ValueOf("ack")
        .DefaultIcon = Data.ValueOf("icon")
        .DefaultSound = Data.ValueOf("sound")
        .DefaultText = Data.ValueOf("text")
        .DefaultTimeout = Val(Data.ValueOf("duration"))
        .DefaultTitle = Data.ValueOf("title")

    End With

    frmAbout.bUpdateClassList mInfo.Token
    AddClass = -1

End Function

Public Function RemClass(ByRef Data As BPackedData) As Long

    ' /* validate some key stuff */

    If (Data Is Nothing) Then
        g_Debug "TApp.RemClass(): config error", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_FAILED
        Exit Function

    End If

    If Data.ValueOf("all") = "1" Then
        g_Debug "TApp.RemClass(): removing all classes..."
        mAlerts = 0
        ReDim mAlert(0)
        AddAlert "_all", ""
        RemClass = -1
        frmAbout.bUpdateClassList mInfo.Token

    ElseIf Not Data.Exists("id") Then
        g_Debug "TApp.RemClass(): required arg 'id' missing", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_ARG_MISSING


    ElseIf Data.ValueOf("id") = "_all" Then
        g_Debug "TApp.RemClass(): cannot remove '_all' class", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_FAILED
        Exit Function

    End If


Dim i As Long

    i = uIndexOf(Data.ValueOf("id"))
    If i = 0 Then
        g_Debug "TApp.RemClass(): class '" & Data.ValueOf("id") & "' not found", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_CLASS_NOT_FOUND
        Exit Function

    End If

Dim j As Long

    If i < mAlerts Then
        For j = i To (mAlerts - 1)
            Set mAlert(j) = mAlert(j + 1)

        Next j

    End If

    mAlerts = mAlerts - 1
    ReDim Preserve mAlert(mAlerts)
    frmAbout.bUpdateClassList mInfo.Token
    RemClass = -1

End Function

Public Function Signature() As String

    Signature = mInfo.Signature

End Function

Public Function Flags() As SNARL41_APP_FLAGS

    Flags = mInfo.Flags

End Function

Public Function IsSnarlApp() As Boolean

    If IsWindow(mInfo.hWnd) <> 0 Then _
        IsSnarlApp = (GetProp(mInfo.hWnd, "snarl_app") = 1) Or ((mInfo.Flags And SNARL41_APP_IS_WINDOWLESS) <> 0)

End Function

Public Sub DoAbout()

    If (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARL41_APP_HAS_ABOUT)) Then _
        PostMessage mInfo.hWnd, sn41AppMsg(), SNARL41_APP_ABOUT, ByVal 0&

End Sub

Public Sub DoSettings()

    If (IsWindow(mInfo.hWnd) <> 0) And ((mInfo.Flags And SNARL41_APP_HAS_PREFS)) Then _
        PostMessage mInfo.hWnd, sn41AppMsg(), SNARL41_APP_PREFS, ByVal 0&

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
