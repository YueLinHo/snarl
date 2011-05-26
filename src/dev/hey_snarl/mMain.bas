Attribute VB_Name = "mMain"
Option Explicit

' Constants that will be used in the API functions
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

' Declare the needed API functions
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal bsName As String, ByVal buff As String, ByVal ch As Long) As Long

'Private Const WM_USER = &H400
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Public Sub Main()
'Dim sz() As String
Dim hr As Long

    If Command$ = "" Then
        uHint SNARL_ERROR_BAD_PACKET

    Else
        uEnd (snDoRequest(g_RemoveQuotes(Command$)))

'        sz() = Split(Command$, " ")
'        uWrite UBound(sz)
'
'        If UBound(sz) <> 0 Then
'            ' /* not the right number of args */
'            uHint SNARL_ERROR_ARG_MISSING
'
'        Else
'            uEnd (snDoRequest(sz(0)))
'
'        End If

    End If

End Sub

Private Sub uHint(ByVal ExitCode As SNARL_STATUS_CODE)

    uWrite "HeySnarl " & CStr(App.Major) & "." & CStr(App.Minor) & " Build " & CStr(App.Revision) & " " & App.LegalCopyright
    uWrite "Usage: heysnarl <request>"
    uWrite "<request> should be enclosed in quotes if it includes spaces"
    ExitProcess ExitCode

End Sub

'======================
' Send output to STDOUT
'======================
'
Private Sub uWrite(ByVal s As String)
Dim llResult As Long

    s = s & vbCrLf
    WriteFile GetStdHandle(STD_OUTPUT_HANDLE), s, Len(s), llResult, ByVal 0&

End Sub

Private Sub uEnd(ByVal ExitCode As SNARL_STATUS_CODE)

    If ExitCode > 0 Then
        uWrite "Ok"

    Else
        uWrite "Failed: " & CStr(ExitCode)

    End If

    ExitProcess ExitCode

End Sub

''============================
'' Get the CGI data from STDIN
''============================
'' Data is collected as a single string. We will read it 1024 bytes at a time.
''
'Sub GetCGIpostData()
'
'   ' Read the standard input handle
'   llStdIn = GetStdHandle(STD_INPUT_HANDLE)
'   ' Get POSTed CGI data from STDIN
'   Do
'      lsBuff = String(1024, 0)    ' Create a buffer big enough to hold the 1024 bytes
'      llBytesRead = 1024          ' Tell it we want at least 1024 bytes
'      If ReadFile(llStdIn, ByVal lsBuff, 1024, llBytesRead, ByVal 0&) Then
'         ' Read the data
'         ' Add the data to our string
'         postData = postData & Left(lsBuff, llBytesRead)
'         If llBytesRead < 1024 Then Exit Do
'      Else
'         Exit Do
'      End If
'   Loop
'
'End Sub


