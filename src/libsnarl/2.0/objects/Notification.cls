VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Notification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum NOTIFICATION_PRIORITY
    LOW_PRIORITY = -1
    NORMAL_PRIORITY = 0
    HIGH_PRIORITY = 1

End Enum

Private Type T_INFO
    Class As String
    Title As String
    Text As String
    Icon As String
    Callback As String
    Duration As Long
    Priority As NOTIFICATION_PRIORITY
    UID As String
    Actions As Actions
    CallbackScript As String
    CallbackScriptType As String
    ReplaceUID As String
    MergeUID As String

End Type

Dim mInfo As T_INFO

Private Sub Class_Initialize()

    mInfo.Duration = -1

End Sub

Friend Function bAsRequest(ByVal AppSignature As String, ByVal Password As String) As String

    bAsRequest = "notify?app-sig=" & AppSignature & "&password=" & Password & _
                 "&class=" & mInfo.Class & "&uid=" & mInfo.UID & _
                 "&title=" & mInfo.Title & "&text=" & mInfo.Text & _
                 "&priority=" & CStr(mInfo.Priority) & "&timeout=" & CStr(mInfo.Duration)

    If mInfo.Icon <> "" Then _
        bAsRequest = bAsRequest & "&icon=" & mInfo.Icon

    If mInfo.Callback <> "" Then _
        bAsRequest = bAsRequest & "&callback=" & mInfo.Callback

    If mInfo.ReplaceUID <> "" Then
        bAsRequest = bAsRequest & "&replace-uid=" & mInfo.ReplaceUID

    ElseIf mInfo.MergeUID <> "" Then
        bAsRequest = bAsRequest & "&merge-uid=" & mInfo.MergeUID

    End If

    If (mInfo.CallbackScript <> "") And (mInfo.CallbackScriptType <> "") Then _
        bAsRequest = bAsRequest & "&callback-script=" & mInfo.CallbackScript & "&callback-script-type=" & mInfo.CallbackScriptType

    ' /* add the actions */

    If Not (mInfo.Actions Is Nothing) Then _
        bAsRequest = bAsRequest & mInfo.Actions.bAsWin32Request

End Function

Public Property Get Class() As String

    Class = mInfo.Class

End Property

Public Property Let Class(ByVal vNewValue As String)

    mInfo.Class = vNewValue

End Property

Public Property Get UID() As String

    UID = mInfo.UID

End Property

Public Property Let UID(ByVal vNewValue As String)

    mInfo.UID = vNewValue

End Property

Public Property Get Title() As String

    Title = mInfo.Title

End Property

Public Property Let Title(ByVal vNewValue As String)

    mInfo.Title = vNewValue

End Property

Public Property Get Text() As String

    Text = mInfo.Text

End Property

Public Property Let Text(ByVal vNewValue As String)

    mInfo.Text = vNewValue

End Property

Public Property Get Icon() As String

    Icon = mInfo.Icon

End Property

Public Property Let Icon(ByVal vNewValue As String)

    mInfo.Icon = vNewValue

End Property

Public Property Get DefaultCallback() As String

    DefaultCallback = mInfo.Callback

End Property

Public Property Let DefaultCallback(ByVal vNewValue As String)

    mInfo.Callback = vNewValue

End Property

Public Property Get Priority() As NOTIFICATION_PRIORITY

    Priority = mInfo.Priority

End Property

Public Property Let Priority(ByVal vNewValue As NOTIFICATION_PRIORITY)

    mInfo.Priority = vNewValue

End Property

Public Property Get Duration() As Long

    Duration = mInfo.Duration

End Property

Public Property Let Duration(ByVal vNewValue As Long)

    mInfo.Duration = vNewValue

End Property

Public Property Get Actions() As Actions

    Set Actions = mInfo.Actions

End Property

Public Property Let Actions(ByVal vNewValue As Actions)

    Set mInfo.Actions = vNewValue

End Property

Public Property Get CallbackScript() As String

    DefaultCallback = mInfo.CallbackScript

End Property

Public Property Let CallbackScript(ByVal vNewValue As String)

    mInfo.CallbackScript = vNewValue

End Property

Public Property Get CallbackScriptType() As String

    CallbackScriptType = mInfo.CallbackScriptType

End Property

Public Property Let CallbackScriptType(ByVal vNewValue As String)

    mInfo.CallbackScriptType = vNewValue

End Property

Public Property Get ReplaceUID() As String

    ReplaceUID = mInfo.ReplaceUID

End Property

Public Property Let ReplaceUID(ByVal vNewValue As String)

    mInfo.ReplaceUID = vNewValue

End Property

Public Property Get MergeUID() As String

    MergeUID = mInfo.MergeUID

End Property

Public Property Let MergeUID(ByVal vNewValue As String)

    mInfo.MergeUID = vNewValue

End Property
