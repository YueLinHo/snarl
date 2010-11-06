VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mId As Long
Dim WithEvents myConnection As CSocket
Attribute myConnection.VB_VarHelpID = -1
Dim mIncludeHTML As Boolean

Public Sub Accept(ByVal Id As Long, ByVal IncludeHTML As Boolean)

    mIncludeHTML = IncludeHTML
    Set myConnection = New CSocket
    myConnection.Accept Id

End Sub

Private Sub Class_Terminate()

    myConnection.CloseSocket

End Sub

Private Sub myConnection_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String
Dim sReply As String

    myConnection.GetData sz

    If mIncludeHTML Then _
        sReply = "HTTP/1.0 200 OK" & vbCrLf & vbCrLf & "<html><body>"

    sReply = sReply & "SNP/2.0/" & uDecode(sz)

    If mIncludeHTML Then _
        sReply = sReply & "</body></html>" & vbCrLf & vbCrLf

    myConnection.SendData sReply
    myConnection.CloseSocket

End Sub

Private Function uDecode(ByVal Data As String) As String
Dim sz As String
Dim s() As String

    s = Split(Data, vbCrLf)
    If UBound(s) < 0 Then
        ' /* bad request */
        uDecode = "107/Bad Packet"
        Exit Function

    End If

    sz = s(0)

'    snShowMessage "snp", "'" & sz & "' == " & Asc(Right$(sz, 1))

    If (g_SafeLeftStr(sz, "5") <> "GET /") Or (InStr(sz, "HTTP") = 0) Then
        ' /* bad request */
        uDecode = "107/Bad Packet"
        Exit Function

    End If
    
Dim i As Integer
    
    ' /* HTTP/SNP syntax is slightly different:
    '
    '   <command>[?<arg>=<value>[&<arg>=<value>]]
    '
    '   e.g. register?signature=application/vnd-fullphat.test
    '
    ' */

    sz = g_SafeRightStr(sz, Len(sz) - 5)
    i = InStr(sz, "HTTP")
    sz = g_SafeLeftStr(sz, i - 2)

Dim pData As BPackedData
Dim szCmd As String

    ' /* get the command */

    i = InStr(sz, "?")
    If i = 0 Then
        ' /* no data, just a command */
        szCmd = sz
        sz = ""

    Else
        szCmd = g_SafeLeftStr(sz, i - 1)            ' // command
        sz = uCleanURL(g_SafeRightStr(sz, Len(sz) - i))

    End If

    Set pData = New BPackedData
    pData.SetTo sz, "&", "="

Dim hr As Long

    Select Case LCase$(szCmd)
    Case "register"
        hr = sn41RegisterApp(pData.ValueOf("id"), pData.ValueOf("title"), pData.ValueOf("icon"))
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/ERR"

        Else
            uDecode = "0/OK/" & CStr(hr)

        End If

    Case "unregister"
        hr = sn41UnregisterApp(pData.ValueOf("token"))
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK"

        End If

    Case "update_app"
        hr = sn41UpdateApp(Val(pData.ValueOf("token")), pData.ValueOf("title"), pData.ValueOf("icon"))
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK"

        End If

'    SNARL41_SET_CALLBACK

    Case "add_class"
        hr = sn41AddClass(Val(pData.ValueOf("token")), pData.ValueOf("id"), pData.ValueOf("name"))
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK"

        End If

    Case "remove_class"
        hr = sn41RemoveClass(Val(pData.ValueOf("token")), pData.ValueOf("id"))
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK"

        End If

    Case "notify"
        sz = pData.AsString
        sz = Replace$(sz, "=", "::")
        sz = Replace$(sz, "&", "#?")

        hr = sn41Notify(Val(pData.ValueOf("token")), pData.ValueOf("id"), sz)
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK/" & CStr(hr)

        End If

'    SNARL41_UPDATE_NOTIFICATION
'    SNARL41_HIDE_NOTIFICATION
'    SNARL41_IS_NOTIFICATION_VISIBLE
'
'    SNARL41_LAST_ERROR

    Case "last_error"
        uDecode = "0/OK/" & CStr(sn41GetLastError)

    Case "version"
        hr = sn41GetVersion()
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK/" & CStr(hr)

        End If

    Case "test"
        hr = sn41EZNotify(0, "", "SNP/HTTP", "Test Message", -1, g_MakePath(App.Path) & "icon.png")
        If hr = 0 Then
            uDecode = CStr(sn41GetLastError) & "/" & uErr(sn41GetLastError)

        Else
            uDecode = "0/OK/" & CStr(hr)

        End If

    Case Else
        uDecode = CStr(SNARL41_ERROR_UNKNOWN_COMMAND) & "/Unknown Command '" & szCmd & "'"

    End Select

End Function

Private Function uCleanURL(ByVal URL As String) As String
Dim i As Integer

    uCleanURL = URL

    For i = &H20 To &H7F
        uCleanURL = Replace$(uCleanURL, "%" & g_HexStr(i, 2), Chr$(i))

    Next i

End Function

Private Function uErr(ByVal Error As Long) As String

    Select Case Error

    Case SNARL41_SUCCESS
        uErr = "Success"

    Case SNARL41_ERROR_FAILED
        uErr = "Unspecified failure"

    Case SNARL41_ERROR_UNKNOWN_COMMAND
        uErr = "Unknown command"

    Case SNARL41_ERROR_TIMED_OUT
        uErr = "Timed out"

    Case SNARL41_ERROR_ARG_MISSING
        uErr = "Required argument missing"

    Case SNARL41_ERROR_SYSTEM
        uErr = "System error"

    Case SNARL41_ERROR_NOT_RUNNING
        uErr = "Not running"

    Case SNARL41_ERROR_NOT_REGISTERED
        uErr = "App is not registered"

    Case SNARL41_ERROR_ALREADY_REGISTERED
        uErr = "App is already registered"

    Case SNARL41_ERROR_CLASS_ALREADY_EXISTS
        uErr = "Class already exists"

    Case SNARL41_ERROR_CLASS_BLOCKED
        uErr = "Class blocked"

    Case SNARL41_ERROR_CLASS_NOT_FOUND
        uErr = "Class not found"

    Case SNARL41_ERROR_NOTIFICATION_NOT_FOUND
        uErr = "Notification not found"

    Case SNARL41_ERROR_FLOODING
        uErr = "Flooding detect"

    Case Else
        uErr = "Undefined error"

    End Select

End Function
