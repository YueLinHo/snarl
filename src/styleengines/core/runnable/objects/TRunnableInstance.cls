VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunnableInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunnableInstance.cls
    '/
    '/  Description:    Represents a single on-screen instance of a runnable style
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const SEE_MASK_NOCLOSEPROCESS = &H40

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long

End Type

Private Declare Function ShellExecuteEx Lib "shell32" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim mTimeout As Long
Dim mContent As MImage
Dim mStyle As TRunnableStyle

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

    Set IStyleInstance_GetContent = mContent

End Function

Private Function IStyleInstance_Pulse() As Boolean
End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)
End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)
Dim szPath As String
Dim lPid As Long

    On Error GoTo er

    DeleteFile mStyle.Path & "content.png"

    ' /* launch the style provider passing the V42 args as a single param */

Debug.Print "running '" & mStyle.Executable & "'..."
'MsgBox Chr$(34) & mStyle.Executable & Chr$(34) & " " & Chr$(34) & NotificationInfo.Text & "#?scheme::" & NotificationInfo.Scheme & Chr$(34)
Debug.Print "--"
Debug.Print NotificationInfo.Text
Debug.Print "--"

Dim pei As SHELLEXECUTEINFO

    With pei
        .cbSize = LenB(pei)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .lpFile = mStyle.Executable
        .lpParameters = Chr$(34) & NotificationInfo.Text & "#?scheme::" & NotificationInfo.Scheme & Chr$(34)
        .lpDirectory = mStyle.Path

    End With

    If ShellExecuteEx(pei) = 0 Then
        Debug.Print "FAILED: " & g_ApiError()

    Else
        ' /* wait for it to finish */
        WaitForSingleObject pei.hProcess, mTimeout
        CloseHandle pei.hProcess

    End If

    ' /* style should have created a PNG for us... */

Dim pb As mfxBitmap

Debug.Print "reading image..."

    If mStyle.IsRedirect Then

    Else
        Set pb = load_image_obj(mStyle.Path & "content.png")
        If is_valid_image(pb) Then _
            Set mContent = pb

    End If

    ' /* zap it now we have it */

Debug.Print "deleting..."
    DeleteFile mStyle.Path & "content.png"
    Exit Sub

er:
    Debug.Print "xx: " & Err.Description
    Resume Next

End Sub

Public Sub SetStyleX(ByRef Style As TRunnableStyle, Optional ByVal Timeout As Long = 500)

    Set mStyle = Style
    mTimeout = Timeout

End Sub
