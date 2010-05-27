VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mInfo As notification_info
Dim mUID As Long

Dim WithEvents theForwarder As SNBConnection
Attribute theForwarder.VB_VarHelpID = -1

Private Sub theForwarder_ConnectionStateChange(ByVal State As snb11.SNB_RETURN_CODE)
Dim sz() As String
'Dim ip() As String

    Debug.Print State

    If State = SNB_SUCCESS Then
        sz = Split(mInfo.Title, "#?")
'        ip = Split(get_ip_address_table(), " ")
        theForwarder.ShowNotification "", sz(1), mInfo.Text & "\n\n" & sz(0) & " on " & g_GetComputerName() & " (" & g_GetUserName() & ")", -1

    End If

    g_removeforwarder mUID

End Sub

Public Sub Init(ByVal RemoteComputer As String, ByRef NotificationInfo As notification_info, ByVal UID As Long)

    LSet mInfo = NotificationInfo
    mUID = UID
    Set theForwarder = New SNBConnection
    theForwarder.Connect RemoteComputer

End Sub

Public Function UID() As Long

    UID = mUID

End Function
