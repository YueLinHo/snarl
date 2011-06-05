VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" (ByVal pszURL As String, ByVal pszEscaped As String, pcchEscaped As Long, ByVal dwFlags As Long) As Long

''Treat entire URL param as one URL segment
'Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
'Private Const URL_ESCAPE_PERCENT         As Long = &H1000
'Private Const URL_UNESCAPE_INPLACE       As Long = &H100000
'
''escape #'s in paths
'Private Const URL_INTERNAL_PATH          As Long = &H800000
'Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
'Private Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
'Private Const URL_DONT_SIMPLIFY          As Long = &H8000000

Dim mRequest As CHTTPRequest

Implements IStyleInstance

Private Sub IStyleInstance_AdjustPosition(x As Long, y As Long, Alpha As Integer, Done As Boolean)
End Sub

Private Function IStyleInstance_GetContent() As melon.MImage

End Function

Private Function IStyleInstance_Pulse() As Boolean

End Function

Private Sub IStyleInstance_Show(ByVal Visible As Boolean)

End Sub

Private Sub IStyleInstance_UpdateContent(NotificationInfo As libSnarlStyles.notification_info)

    If gConfig.UserKey = "" Then
        g_Debug "StyleInstance.UpdateContent(): UserKey not set", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    If (gConfig.OnlyShowPriorityNotifications) And ((NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY) = 0) Then
        g_Debug "StyleInstance.UpdateContent(): showing non-priority notifications disabled"
        Exit Sub

    End If

Dim sz As String

    sz = "https://prowl.weks.net/publicapi/add?" & _
         "apikey=" & gConfig.UserKey & _
         "&application=" & uGetAppName(NotificationInfo.Title) & _
         "&event=" & g_URLEncode(uGetTitle(NotificationInfo.Title)) & _
         "&description=" & g_URLEncode(uTidy(NotificationInfo.Text)) & _
         "&priority=" & IIf(NotificationInfo.Flags And S_NOTIFICATION_IS_PRIORITY, "1", "0")

'Dim szEscaped As String
'Dim cb As Long
'
'    szEscaped = Space$(4096)
'    cb = Len(szEscaped)
'
'    If UrlEscape(sz, szEscaped, cb, URL_DONT_SIMPLIFY) <> 0 Then
'        g_Debug "StyleInstance.UpdateContent(): url escaping failed", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    Else
'        szEscaped = g_SafeLeftStr(szEscaped, cb)
'
'    End If

'    MsgBox sz

    Set mRequest = New CHTTPRequest

    If gConfig.UseProxyServer Then _
        mRequest.SetProxyInfo gConfig.ProxyServer, gConfig.ProxyServerPort, , gConfig.ProxyUsername, gConfig.ProxyPassword

    mRequest.GetContent sz, gConfig.Timeout
    
'    mServer.Open "GET", " HTTP/1.0", -1

End Sub

Private Function uGetAppName(ByVal Title As String) As String
Dim i As Integer

    uGetAppName = "Snarl"

    i = InStr(Title, "#?")
    If i < 1 Then _
        Exit Function

    uGetAppName = g_SafeLeftStr(Title, i - 1)

End Function

Private Function uGetTitle(ByVal Title As String) As String
Dim i As Integer

    i = InStr(Title, "#?")
    If i < 1 Then _
        Exit Function

    uGetTitle = g_SafeRightStr(Title, Len(Title) - (i + 1))

End Function

'Private Function uReplacePercent(ByVal sz As String) As String
'Dim i As Integer
'Dim j As Integer
'
'    i = InStr(sz, "%")
'    Do While i > 0
'        j = Val("&H" & mId$(sz, i + 1, 2))
'        uReplacePercent = uReplacePercent & g_SafeLeftStr(sz, i - 1) & Chr$(j)
'        sz = g_SafeRightStr(sz, Len(sz) - (i + 2))
'        i = InStr(sz, "%")
'
'    Loop
'
'    uReplacePercent = uReplacePercent & sz
'
'End Function

Private Function uTidy(ByVal Text As String) As String
'Dim sz As String
'Dim a As String
'Dim i As Integer

'    For i = 1 To Len(Text)
'        a = mId$(Text, i, 1)
'        sz = sz & a & " " & g_HexStr(Asc(a), 2) & " "
'
'    Next i
'
'    MsgBox sz


    If gConfig.ReplaceCRLFs Then
        Text = Replace$(Text, Chr$(13), "")
        Text = Replace$(Text, Chr$(10), " ")

    End If

    uTidy = Text

End Function
