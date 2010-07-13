Attribute VB_Name = "mMain"
Option Explicit

    ' /*
    '
    '
    ' */

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const WM_CLOSE = &H10
Public Const WM_TEST = &H400 + 1
Public Const WM_NOTIFICATION = &H400 + 2
'Public Const WM_RELOAD = &H400 + 3
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type T_CONFIG
    Timeout As Long
    UseFeedIcon As Boolean
    FeedRefresh As Long

End Type

Public g_Config As T_CONFIG

Public gDebugMode As Boolean
Public gToken As Long

Public gPanel As BPrefsPanel

Public Const CLASS_NAME = "w>snarss"

Public Sub Main()
Dim hWndExisting As Long

    hWndExisting = FindWindow(CLASS_NAME, CLASS_NAME)

    If InStr(Command$, "-quit") Then
        ' /* quit any existing instance (but don't run this one) */
        If IsWindow(hWndExisting) <> 0 Then _
            SendMessage hWndExisting, WM_CLOSE, 0, ByVal 0&

        Exit Sub

'    ElseIf InStr(Command$, "-reload") Then
'        ' /* if an existing instance is running, tell it to reload tasks */
'        If IsWindow(hWndExisting) <> 0 Then
'            SendMessage hWndExisting, WM_RELOAD, 0, ByVal 0&
'
'        End If
    End If

    If hWndExisting <> 0 Then _
        Exit Sub

    If Not uGotMiscResource() Then
        MsgBox "misc.resource missing or damaged" & vbCrLf & vbCrLf & "This can happen if melon is uninstalled - try reinstalling melon in the first instance", vbCritical Or vbOKOnly, App.Title
        Exit Sub

    End If

    gDebugMode = (InStr(Command$, "-debug") <> 0)

    If gDebugMode Then _
        frmMain.Show

Dim hWnd As Long

    EZRegisterClass CLASS_NAME
    hWnd = EZAddWindow(CLASS_NAME, New TWindow, CLASS_NAME)

    frmMain.List1.AddItem "window: " & g_HexStr(hWnd)
    frmMain.Tag = CStr(hWnd)

    With New BMsgLooper
        .Run

    End With

    If Not (gPanel Is Nothing) Then _
        gPanel.Quit

    EZRemoveWindow hWnd
    EZUnregisterClass CLASS_NAME
    Unload frmMain

End Sub

Private Function uGotMiscResource() As Boolean

    On Error Resume Next

Dim i As Long

    Err.Clear
    i = processor_count()
    uGotMiscResource = (Err.Number = 0)

End Function


