VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theFeed As BRSSFeed
Attribute theFeed.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Dim mURL As String
Dim mLastAttempt As Date
Dim mSuccess As Boolean
Dim mEnabled As Boolean
Dim mFirstRun As Boolean
Dim mRefreshing As Boolean

Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Implements MObject

Private Property Get MObject_Type() As String
End Property

Public Function Init(ByVal URL As String, Optional ByVal Enabled As Boolean = True) As Boolean

    mURL = URL
    mEnabled = Enabled
'    mFirstRun = True

    Set theFeed = New BRSSFeed
    Init = Refresh()

    ' /* on success, start a 1 minute timer with a +/- 2 second offset */

    If Init Then _
        Set theTimer = new_BTimer(60000 + ((Rnd * 4000) - 2000))

End Function

Public Sub Quit()

    g_Debug "TFeed.Quit()", LEMON_LEVEL_PROC

    Set theTimer = Nothing

    If Not (theFeed Is Nothing) Then _
        Set theFeed = Nothing

End Sub

'Public Function WasSuccessful() As Boolean
'
'    WasSuccessful = mSuccess
'
'End Function

Private Function uFixStr(ByVal str As String) As String
Dim i As Integer
Dim c As Integer
Dim sz As String
Dim x As Integer
Dim a As String

    ' /* replace consecutive whitespace characters (TAB, CR, LF) with a single space */

    c = Len(str)
    If c Then
        For i = 1 To c
'            g_Debug Mid$(str, i, 1) & " (" & Asc(Mid$(str, i, 1)) & ")"

            Select Case Asc(Mid$(str, i, 1))
            Case 9, 10, 13
                ' /* whitespace character so increment count */
                x = x + 1

            Case Else
                ' /* non-whitespace character so add a single space */
                If x > 0 Then
                    ' /* but _only_ if the previous character added _wasn't_ a space! */
                    If a <> " " Then _
                        sz = sz & " "

                    x = 0

                End If

                a = Mid$(str, i, 1)
                sz = sz & a

            End Select

        Next i

        uFixStr = sz

    End If

End Function

Public Function IsValid() As Boolean

    If Not (theFeed Is Nothing) Then _
        IsValid = theFeed.IsValid


End Function

Private Sub theFeed_Completed(ByVal WasSuccessful As Boolean)
Dim pList As BControl

    On Error Resume Next

    mLastAttempt = Now()
    mSuccess = WasSuccessful
    mRefreshing = False

'    Debug.Print "imgurl: " & theFeed.Channel.ImageUrl
'    Debug.Print "imglink: " & theFeed.Channel.ImageLink

    ' /* if the prefs panel is open, tell the feed list to refresh */

    If Not (gPanel Is Nothing) Then
        If gPanel.Find("feed_list", pList) Then _
            pList.Notify "update", Nothing

    End If

    If Not theFeed.IsValid Then
        frmMain.Add "FAIL: " & g_SafeRightStr(mURL, 80, True)
        Exit Sub

    ElseIf theFeed.CountItems < 1 Then
        frmMain.Add "EMPTY: " & g_SafeRightStr(mURL, 80, True)
        Exit Sub

    End If

Static szLastEntry As String
Static szEntry As String

    With theFeed.ItemAt(1)
        szEntry = .Description '//  .Title & "+" & .Description
        If szEntry <> szLastEntry Then
            ' /* top story has changed */
            frmMain.Add "UPDATED: " & g_SafeRightStr(mURL, 80, True)
            szLastEntry = szEntry
            If Not mFirstRun Then
                ShowHeadline

            Else
                mFirstRun = False

            End If
        End If
    End With

End Sub

Private Sub theTimer_Pulse()

    Refresh

End Sub

Public Function URL() As String

    URL = mURL

End Function

Private Function uEntriesStr() As String

    If (theFeed Is Nothing) Then _
        Exit Function

    Select Case theFeed.CountItems
    Case 0
        uEntriesStr = "No entries"

    Case 1
        uEntriesStr = "1 entry"

    Case Else
        uEntriesStr = CStr(theFeed.CountItems) & " entries"

    End Select

End Function

Public Sub SetEnabled(ByVal Enabled As Boolean)

    mEnabled = Enabled

    If Not (theTimer Is Nothing) Then _
        theTimer.SetEnabled Enabled

End Sub

Public Function IsEnabled() As Boolean

    IsEnabled = mEnabled

End Function

Public Sub ShowHeadline()
Static i As Long

    If (theFeed Is Nothing) Then _
        Exit Sub

    If theFeed.CountItems < 1 Then _
        Exit Sub

    With theFeed.ItemAt(1)
        i = sn41EZNotify(gToken, "", _
                         .Title, _
                         IIf(.Description <> "", .Description, "<no entry>") & vbCrLf & vbCrLf & theFeed.Channel.Title, _
                         , _
                         theFeed.Channel.ImageUrl, , _
                         .Link)

    End With

End Sub

Public Sub ShowSummary()
Static i As Long
Static sz As String

    If (theFeed Is Nothing) Then _
        Exit Sub

    If theFeed.CountItems < 1 Then _
        Exit Sub

    With theFeed
        For i = 1 To MIN(8, .CountItems)
            sz = sz & "• " & .ItemAt(i).Title & IIf(i < .CountItems, vbCrLf, "")

        Next i

        i = sn41EZNotify(gToken, "", _
                         .Channel.Title, _
                         sz, _
                         , _
                         .Channel.ImageUrl, , _
                         .Channel.Link)

    End With

End Sub

Public Function Refresh() As Boolean

    If mRefreshing Then _
        Exit Function

    Refresh = (theFeed.GetContent(mURL) = B_OK)
    mRefreshing = True

End Function

'Public Function IsRefreshing() As Boolean
'
'    IsRefreshing = mRefreshing
'
'End Function

Public Sub FeedInfo()
Static sz As String
Static i As Long

    On Error Resume Next

    If (theFeed Is Nothing) Then _
        Exit Sub

    If Not theFeed.IsValid Then
        i = sn41EZNotify(gToken, "", _
                         "Could not retrieve feed", _
                         "There may be a problem with your Internet connection, or the server providing the feed may be off-line - try accessing the feed from a browser", , _
                         g_MakePath(App.Path) & "failed.png")

    ElseIf (theFeed.Channel Is Nothing) Then
        i = sn41EZNotify(gToken, "", _
                         "Invalid RSS Feed", _
                         "Data was returned but it does not appear to be a valid RSS feed", , _
                         g_MakePath(App.Path) & "failed.png")

    Else
        With theFeed

            Debug.Print .RawXML

            sz = IIf(.Channel.Description <> "", .Channel.Description & vbCrLf, "") & _
                 IIf(.Channel.Copyright <> "", .Channel.Copyright & vbCrLf, "") & _
                 IIf(.FeedVersion <> "", "Version " & .FeedVersion, "")

            If sz <> "" Then _
                i = sn41EZNotify(gToken, "", _
                                 .Channel.Title, _
                                 sz, _
                                 , _
                                 .Channel.ImageUrl, , _
                                 .Channel.Link)

        End With
    End If

End Sub

Public Function TitleOrURL() As String

    If Not (theFeed Is Nothing) Then
        If theFeed.IsValid Then _
            TitleOrURL = Replace$(theFeed.Channel.Title, "|", ":")

    End If

    If TitleOrURL = "" Then _
        TitleOrURL = mURL

End Function

Public Function Status() As String

    If Not (theFeed Is Nothing) Then
        If theFeed.IsValid Then
            If mRefreshing Then
                Status = "Refreshing..."

            Else
                Status = uEntriesStr() & ".  Last refreshed: " & mLastAttempt

            End If

        Else
            Status = "Failed.  Last attempt: " & mLastAttempt

        End If

    Else
        Status = "Error."

    End If

End Function
