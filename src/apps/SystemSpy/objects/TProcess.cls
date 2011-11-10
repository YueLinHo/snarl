VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mInfo As GMPROCESSINFO
Dim mDesc As String

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mInfo.ProcessId)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mInfo.ProcessName

End Function

Friend Sub bInit(ByRef Info As GMPROCESSINFO)

    LSet mInfo = Info
    mDesc = g_GetExeDescription(Info.FullPath)
'    If mDesc = "" Then _
'        mDesc = "<no description>"

'    Debug.Print "!" & mDesc & "!"

End Sub

Public Function PID() As Long

    PID = mInfo.ProcessId

End Function

Public Function Name() As String

    Name = mInfo.ProcessName

End Function

Public Function Description() As String

    Description = mDesc

End Function

Public Function Path() As String

    Path = mInfo.FullPath

End Function
