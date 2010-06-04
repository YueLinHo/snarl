VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Index As Long                ' // set by AlertAt() or FindAlert()

    ' /* Class defaults - set by snSetAlertClassDefault()  (V38) */

Public DefaultTitle As String
Public DefaultText As String
Public DefaultTimeout As Long
Public DefaultSound As String
Public DefaultIcon As String
Public DefaultAck As String

Public Token As Long                '// V41
Dim mLastToken As Long              '// V41: returned in certain cases (e.g. during duplicate suppression)

    ' /* V39 - used to suppress duplicates */
Dim mPrevTitle As String
Dim mPrevText As String
Dim mPrevTick As Long

Public Enum E_CLASS_DUPLICATES
    E_CLASS_DUPLICATES_ALLOW
    E_CLASS_DUPLICATES_BLOCK
    E_CLASS_DUPLICATES_DELAY

End Enum

Private Type T_ALERT
    Class As String                         '// unique identifier - once assigned, can't be changed
    Desc As String                          '// displayed name - added in V38, doc'd in V39
    App As TApp                             '// app who this class belongs to
    Config As ConfigSection                 '// V41: persistent settings
    Defaults As BPackedData

    Forward As BPackedData                  '// each style/scheme to use

End Type

Dim mContent As T_ALERT

Implements MObject

Private Sub Class_Initialize()

    With mContent
        .Class = "_internal"
        .Desc = ""
        Set .App = New TApp

    End With

End Sub

Private Property Get MObject_Type() As String
End Property

Friend Sub bInit(ByVal Id As String, ByVal Name As String, ByRef Owner As TApp, ByRef Settings As ConfigSection)

    ' /* initialize */

    g_Debug "TAlert.bInit()", LEMON_LEVEL_PROC_ENTER

    With mContent
        .Class = Id
        .Desc = Name
        Set .App = Owner
        Set .Config = Settings

        ' /* defaults */
        ' /* added in R2.1 (moved here in V41): if this global setting is enabled, new classes are auto-disabled when created */

        g_Debug "setting defaults..."

        Set .Defaults = New BPackedData
        .Defaults.SetTo "enabled::" & IIf(g_Prefs.ignore_new_classes, "0", "1") & "#?" & _
                        "show_on_screen::1#?" & _
                        "position::" & CStr(E_START_DEFAULT_POS) & "#?" & _
                        "duration::" & CStr(E_DURATION_APP_DECIDES) & "#?" & _
                        "custom_timeout::10#?" & _
                        "custom_priority::0#?" & _
                        "priority::0#?" & _
                        "use_style::0#?" & _
                        "style::#?" & _
                        "use_custom_icon::0#?" & _
                        "custom_icon::#?" & _
                        "duplicates::" & CStr(E_CLASS_DUPLICATES_ALLOW) & "#?" & _
                        "use_custom_sound::0#?" & _
                        "custom_sound::"

        Set .Forward = New BPackedData

        ' /* read forwarding info */

        If Not (Settings Is Nothing) Then _
            .Forward.SetTo Settings.GetValueWithDefault("forward", "")

    End With

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function IsEnabled() As Boolean

    If Not (mContent.Config Is Nothing) Then _
        IsEnabled = (Me.GetValue("enabled") = "1")

End Function

Public Function Name() As String

    Name = mContent.Class

End Function

Public Function CustomStyle() As String

    If Not (mContent.Config Is Nothing) Then _
        CustomStyle = Me.GetValue("style")

End Function

Public Function StyleName() As String

    StyleName = style_GetStyleName(CustomStyle)

End Function

Public Function SchemeName() As String

    SchemeName = style_GetSchemeName(CustomStyle)

End Function

Public Sub SetStyleAndScheme(ByVal Style As String, ByVal Scheme As String)

    If (mContent.Config Is Nothing) Then _
        Exit Sub

Dim sz As String

    sz = LCase$(Style & "/" & Scheme)               ' // convert to style/scheme pair

    If sz = LCase$(Me.CustomStyle) Then _
        Exit Sub                                    ' // ignore if no change

'    If ((Me.GetValue("style") = "") And (sz = g_ConfigGet("default_style"))) Or (sz = LCase$(Me.GetValue("style"))) Then _
        Exit Sub

    Me.SetValue "style", sz
    Debug.Print "*** scheme changed to '" & sz & "' for class " & mContent.Class & " ***"

End Sub

'Public Function Sound() As E_CLASS_SOUND
'
'    Sound = E_CLASS_SOUND_APP_DECIDES
'
'    If (mContent.Config Is Nothing) Then _
'        Exit Function
'
'Dim sz As String
'
'    If mContent.Config.Find("sound", sz) Then _
'        Sound = Val(sz)
'
'End Function
'
'Public Function CustomSound() As String
'
'    If (mContent.Config Is Nothing) Then _
'        Exit Function
'
'    CustomSound = mContent.Config.GetValueWithDefault("custom_sound", "")
'
'End Function

Public Sub Reset()

    If (mContent.Config Is Nothing) Or (mContent.Defaults Is Nothing) Then _
        Exit Sub

Dim s1 As String
Dim s2 As String

    With mContent
        .Defaults.Rewind
        Do While .Defaults.GetNextItem(s1, s2)
            .Config.Update s1, s2

        Loop

        Set .Forward = New BPackedData

    End With

End Sub

Public Function Description() As String

    If mContent.Class = "_all" Then
        Description = IIf(mContent.App.CountAlerts > 1, "Other Notifications", "All Notifications")

    Else
        Description = IIf(mContent.Desc = "", mContent.Class, mContent.Desc)

    End If

End Function

Public Function IsClickThru() As Boolean

    If Not (mContent.Config Is Nothing) Then _
        IsClickThru = (Me.GetValue("ignore_input") = "1")

End Function

Public Function GetValue(ByVal Name As String) As String

    ' /* pre-set to default */

    If Not (mContent.Defaults Is Nothing) Then _
        GetValue = mContent.Defaults.ValueOf(Name)

Dim sz As String

    If Not (mContent.Config Is Nothing) Then
        If mContent.Config.Find(Name, sz) Then _
            GetValue = sz

    End If

End Function

Public Sub SetValue(ByVal Name As String, ByVal Value As String)

    If Not (mContent.Config Is Nothing) Then
        mContent.Config.Update Name, Value
        mContent.App.WriteConfig

    End If

End Sub

Public Function App() As TApp

    If (mContent.App Is Nothing) Then _
        MsgBox "ALERT!"

    Set App = mContent.App

End Function

Public Sub TestNotification()
Dim pInfo As T_NOTIFICATION_INFO

    With pInfo
        .Title = Me.Description
        .Text = "Test Notification"
        .Timeout = 7
        .IconPath = Me.App.Icon

    End With

    Me.ShowNotification pInfo

End Sub

Friend Function ShowNotification(ByRef Info As T_NOTIFICATION_INFO) As Long

    ' /* IMPORTANT! This should *only* be called by TApp.Show() */

    gLastErr = SNARL41_ERROR_SYSTEM
    If (g_NotificationRoster Is Nothing) Then _
        Exit Function

    ' /* do nothing if class is disabled */

    If Me.GetValue("enabled") = "0" Then
        g_Debug "TAlert.ShowNotification(): class '" & mContent.Class & "' has been disabled by the user"
        gLastErr = SNARL41_ERROR_CLASS_BLOCKED
        Exit Function

    End If

    ' /* check for duplicate content */

Dim lDup As E_CLASS_DUPLICATES

    If (Info.Title = mPrevTitle) And (Info.Text = mPrevText) Then
        ' /* duplicate content - what should we do? */

        lDup = Me.GetValue("duplicates")

        If lDup = E_CLASS_DUPLICATES_BLOCK Then
            g_Debug "TAlert.ShowNotification(): blocked duplicate", LEMON_LEVEL_INFO
            ShowNotification = mLastToken
            Exit Function

        ElseIf (lDup = E_CLASS_DUPLICATES_DELAY) And ((GetTickCount() - mPrevTick) <= g_Prefs.suppress_delay) Then
            g_Debug "TAlert.ShowNotification(): blocked duplicate within delay timeout", LEMON_LEVEL_INFO
            ShowNotification = mLastToken
            Exit Function

        End If

    Else
        mPrevTitle = Info.Title
        mPrevText = Info.Text

    End If

    mPrevTick = GetTickCount()


    ' /* figure out timeout */

Dim lDur As E_NOTIFICATION_DURATION

    lDur = Val(Me.GetValue("duration"))

    Select Case lDur
    Case E_DURATION_APP_DECIDES
        ' /* V39: app can modify duration via class default */
        If Me.DefaultTimeout > 0 Then _
            Info.Timeout = Me.DefaultTimeout

    Case E_DURATION_CUSTOM
        Info.Timeout = Val(Me.GetValue("custom_timeout"))

    Case E_DURATION_DEFAULT
        Info.Timeout = g_Prefs.default_duration

    Case Else
        g_Debug "TAlert.ShowNotification(): invalid duration setting '" & CStr(lDur) & "'", LEMON_LEVEL_WARNING
        Info.Timeout = g_Prefs.default_duration

    End Select

    ' /* V38 - if any item isn't provided we use the class default */

    If Info.Title = "" Then _
        Info.Title = Me.DefaultTitle

    If Info.Text = "" Then _
        Info.Text = Me.DefaultText

    ' /* V40.25: custom icon */

    If Me.GetValue("use_custom_icon") = "1" Then
        Info.IconPath = Me.GetValue("custom_icon")

    ElseIf Info.IconPath = "" Then
        Info.IconPath = Me.DefaultIcon

    End If

    If Info.SndFile = "" Then _
        Info.SndFile = Me.DefaultSound

    If Info.DefaultAck = "" Then _
        Info.DefaultAck = Me.DefaultAck

    ' /* V40.25: custom priority */

    If Me.GetValue("custom_priority") = "1" Then _
        Info.Priority = Val(Me.GetValue("priority"))

    ' /* sound */

'    If Me.Sound = E_CLASS_SOUND_CUSTOM Then _
'        SoundFile = Me.CustomSound

    Info.Position = Val(Me.GetValue("position"))

    If Me.GetValue("use_style") = "1" Then _
        Info.StyleToUse = Me.GetValue("style")

    ' /* send it to the notification roster */

    ShowNotification = g_NotificationRoster.Add(Me, Info)

    mLastToken = Info.Token
    Debug.Print "TAlert.ShowNotification(): token was " & CStr(mLastToken)

End Function

Public Sub AddForward(ByVal StyleAndScheme As String)

    If (mContent.Forward Is Nothing) Then _
        Exit Sub

    With mContent
        .Forward.Add LCase$(Replace$(StyleAndScheme, ": ", "/")), ""    ' // add the forwarder
        .Config.Update "forward", mContent.Forward.AsString()           ' // translate the packed data into a string
        .App.WriteConfig                                                ' // write the config

    End With

End Sub

Public Sub RemForward(ByVal StyleAndScheme As String)

    If (mContent.Forward Is Nothing) Then _
        Exit Sub

    StyleAndScheme = LCase$(Replace$(StyleAndScheme, ": ", "/"))

    ' /* BPackedData doesn't have a Remove method (oops) so we rebuild without the removed one */

Dim ppd As BPackedData
Dim sz As String

    Set ppd = New BPackedData

    With mContent.Forward
        .Rewind
        Do While .GetNextItem(sz, "")
            If sz <> StyleAndScheme Then _
                ppd.Add sz, ""

        Loop

    End With

    Set mContent.Forward = ppd

End Sub

Public Function HasForward(ByVal StyleAndScheme As String) As Boolean

    If Not (mContent.Forward Is Nothing) Then _
        HasForward = mContent.Forward.Exists(LCase$(Replace$(StyleAndScheme, ": ", "/")))

End Function

