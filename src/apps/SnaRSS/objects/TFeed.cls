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

Public Function LastAttempt() As Date

    LastAttempt = mLastAttempt

End Function

Public Function WasSuccessful() As Boolean

    WasSuccessful = mSuccess

End Function

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

    mLastAttempt = Now()
    mSuccess = WasSuccessful
    mRefreshing = False

    ' /* if the prefs panel is open, tell the feed list to refresh */

    If Not (gPanel Is Nothing) Then
        If gPanel.Find("feed_list", pList) Then _
            pList.Notify "update", Nothing

    End If

    If (Not WasSuccessful) Or (theFeed.CountItems < 1) Then _
        Exit Sub

Static szLastEntry As String
Static szEntry As String

    With theFeed.ItemAt(1)
        szEntry = .Description '//  .Title & "+" & .Description
        If szEntry <> szLastEntry Then
            ' /* top story has changed */
            frmMain.Add "changed: " & theFeed.Channel.Title
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

Public Function EntriesStr() As String

    If (theFeed Is Nothing) Then _
        Exit Function

    Select Case theFeed.CountItems
    Case 0
        EntriesStr = "No entries"

    Case 1
        EntriesStr = "1 entry"

    Case Else
        EntriesStr = CStr(theFeed.CountItems) & " entries"

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
            sz = sz & .ItemAt(i).Title & IIf(i < .CountItems, vbCrLf, "")

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

    Refresh = (theFeed.GetContent(URL) = B_OK)
    mRefreshing = True

End Function

Public Function IsRefreshing() As Boolean

    IsRefreshing = mRefreshing

End Function

Public Sub FeedInfo()
Static sz As String
Static i As Long

    On Error Resume Next

    If (theFeed Is Nothing) Then _
        Exit Sub

    With theFeed

        Debug.Print .RawXML

        sz = IIf(.Channel.Description <> "", .Channel.Description & vbCrLf, "") & _
             IIf(.Channel.Copyright <> "", .Channel.Copyright & vbCrLf, "") & _
             IIf(.FeedVersion <> "", "Version " & .FeedVersion, "")

'        sz = "Version: " & .FeedVersion & vbCrLf & "Valid: " & .IsValid & vbCrLf & "Category: " & .Channel.Category & vbCrLf & _
             "Copyright: " & .Channel.Copyright & vbCrLf & "" & .Channel.WebMaster

        If sz <> "" Then _
            i = sn41EZNotify(gToken, "", _
                             .Channel.Title, _
                             sz, _
                             , _
                             .Channel.ImageUrl, , _
                             .Channel.Link)

    End With

End Sub


