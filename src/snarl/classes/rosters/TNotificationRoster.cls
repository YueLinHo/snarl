VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotificationRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TNotificationRoster.cls
    '/
    '/  Description:    Manages on-screen notifications
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Type T_NOTIFICATION
    Window As CSnarlWindow
    Column As Integer           ' // used to help determine placement of other notifications
    StyleName As String         ' // R2.31: expanded
    SchemeName As String        ' // R2.31: expanded
    Info As T_NOTIFICATION_INFO ' // R2.4 DR7: full originating information

End Type

Dim mItem() As T_NOTIFICATION
Dim mItems As Long
Dim mToken As Long
Dim mRefs As Long

Dim mMissed() As T_NOTIFICATION_INFO
Dim mMissedCount As Long                        ' // total in the list
Dim mMissedPanel As TMissedNotificationsPanel
Dim mActualMissed As Long                       ' // number that appeared since DnD enabled

Dim mDelayed() As TDelayedNotification
Dim mDelayedCount As Long

Private Type T_QUEUED
    Detail As T_NOTIFICATION_INFO
    Class As TAlert

End Type

Dim mQueue() As T_QUEUED
Dim mQueueCount As Long

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    g_Debug "TNotificationRoster.Close()", LEMON_LEVEL_PROC

    mRefs = mRefs - 1
    If mRefs > 0 Then
        ' /* still open */
        Exit Function

    End If

    ' /* run tidyup code */

Dim i As Long

    If mItems Then
        g_Debug "TNotificationRoster.Close(): count=" & CStr(mItems)
        For i = mItems To 1 Step -1
            If Not (mItem(i).Window Is Nothing) Then _
                mItem(i).Window.Remove

        Next i
    End If

    ReDim mItem(0)
    mItems = 0

    ' /* remove temporary icons */

Dim sz As String

    sz = String$(MAX_PATH + 1, 0)
    GetTempPath MAX_PATH, sz
    sz = g_TrimStr(sz)
    If sz <> "" Then
        sz = g_MakePath(sz)
        i = 1
        Do While g_Exists(sz & "snarl-icon" & CStr(i))
            delete_file sz & "snarl-icon" & CStr(i)
            i = i + 1

        Loop

    End If

    g_Debug "TNotificationRoster.Close(): done"

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        mToken = &H278

    End If

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

'    MRoster_CountItems = mCount

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
'Dim i As Long
'
'    Name = LCase$(Name)
'
'    If mCount Then
'        For i = 1 To mCount
'            If LCase$(mExt(i).Name) = Name Then
'                MRoster_FindItem = i
'                Exit Function
'
'            End If
'        Next i
'    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.MObject

'    If (Index > 0) And (Index <= mCount) Then _
'        Set MRoster_ItemAt = mExt(Index)

End Function

Private Function MRoster_NextItem(Item As melon.MObject) As Boolean

'    If mIndex <= mCount Then
'        Set Item = mExt(mIndex)
'        mIndex = mIndex + 1
'        MRoster_NextItem = True
'
'    End If

End Function

Private Sub MRoster_Rewind()

'    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "4-Nov-2008"

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "notification.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

Friend Function Add(ByRef Info As T_NOTIFICATION_INFO, ByRef Args As BPackedData) As Long

    ' /* To-do: consider having the BPackedData arrive here instead of a T_NOTIFICATION_INFO struct */

    gSetLastError SNARL_ERROR_FAILED
    If g_IsQuitting Then _
        Exit Function               ' // shutting down - can't do anything

    ' /* log only? */

    If g_ConfigGet("log_only") = "1" Then
        g_WriteToLog Info.Title, Info.Text
        Add = -1                    ' // V40.7 - changed from M_OK to -1 (indicates success but no token available)
        gSetLastError SNARL_SUCCESS
        Exit Function               ' // TO-DO: determine if we should return 0 and set last_error instead

    End If

    ' /* for anything else, the class must be valid */

    If (Info.ClassObj Is Nothing) Then
        g_Debug "TNotificationRoster.Add(): Class is null for '" & Info.Title & "/" & Info.Text & "'", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* R2.31: log last time a notification was generated */

    gLastNotification = Now()

    ' /* R2.31: 'duplicates' is now a simple 'stop spam' toggle - what we do with the potential
    '    spam is for future releases; for now we just discard it */

Dim lQuantum As Long

    lQuantum = g_SafeLong(g_ConfigGet("flood_quantum"))
    If lQuantum < 100 Then _
        lQuantum = 750

    If (Info.ClassObj.GetValue("block_spam") = "1") And ((GetTickCount() - Info.ClassObj.LastNotificationGeneratedTick) < lQuantum) Then
        Debug.Print "!!! SPAM !!! " & CStr(lQuantum)

        ' /* R2.31 - for now, block and return a failure code.  For the future, it would be lovely to
        '    allow more functionality (e.g. update notification, append to notification, send to missed log, etc. */

        gSetLastError SNARL_ERROR_FLOODING
        Exit Function

    End If

    ' /* only record the tick of the last _generated_ notification */

    Info.ClassObj.LastNotificationGeneratedTick = GetTickCount()

    gSetLastError SNARL_SUCCESS

    ' /* R2.3: set datestamp */

    Info.DateStamp = Now()

    ' /* R2.4: generate the "OriginalContent" content */

    uFixupContent Info

    ' /* R2.4 DR7 - can include actions in request */

    Set Info.Actions = new_BTagList()

Dim sn As String
Dim sv As String
Dim act() As String

    If Not (Args Is Nothing) Then
        With Args
            .Rewind
            Do While .GetNextItem(sn, sv)
                If sn = "action" Then
                    act = Split(sv, ",")
                    If UBound(act) = 1 Then _
                        Info.Actions.Add new_BTagItem(act(0), act(1))

                End If
            Loop

        End With
    End If

    ' /* R2.3: send to any selected non-window styles */

Dim pCopyInfo As T_NOTIFICATION_INFO
Dim pt As BTagItem
Dim fc As Long

    If Not (Info.ClassObj.ForwardList Is Nothing) Then
        With Info.ClassObj.ForwardList
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                fc = Val(pt.Value)

                Debug.Print "TNotificationRoster.Add(): is_dnd=" & gDNDMode & " is_away=" & g_IsAway(SP_AWAY_MASK) & "cond=" & g_HexStr(fc)

                If (gDNDMode) And ((fc And 1) <> 0) Then
                    fc = -1

                ElseIf (g_IsAway(SP_AWAY_MASK)) And ((fc And 2) <> 0) Then
                    fc = -1

                ElseIf (fc And 4) <> 0 Then
                    fc = -1

                Else
                    fc = 0

                End If

                If fc <> 0 Then
                    Debug.Print "TNotificationRoster.Add(): forwarding to " & pt.Name

                    LSet pCopyInfo = Info
                    pCopyInfo.StyleName = style_GetStyleName(pt.Name)
                    pCopyInfo.SchemeName = style_GetSchemeName(pt.Name)
                    uCopyTo pCopyInfo, Info.ClassObj

                End If

            Loop

        End With
    End If

    ' /* 40.7 - icon can be an http:// or https:// url - if it is we initiate an
    '    async download of the icon and put the notification onto a stack while
    '    we wait for the download to complete */

Dim pDelayed As TDelayedNotification

    If (LCase$(g_SafeLeftStr(Info.IconPath, 7)) = "http://") Or (LCase$(g_SafeLeftStr(Info.IconPath, 8)) = "https://") Then
        g_Debug "TNotificationRoster.Add(): remote icon specified, starting async download..."

        ' /* to-do: allow user to suppress this by specifying a maximum download wait
        '    of zero seconds - in which case use the class default icon */

        Set pDelayed = New TDelayedNotification
        If pDelayed.Init(Info.IconPath) Then

            ' /* allocate a token now */

            mToken = mToken + 1
            Info.Token = mToken
            pDelayed.SetTo Info

            mDelayedCount = mDelayedCount + 1
            ReDim Preserve mDelayed(mDelayedCount)
            Set mDelayed(mDelayedCount) = pDelayed      '// add to the queue

            Add = mToken                                '// return the token that was allocated
            Exit Function

        Else
            g_Debug "TNotificationRoster.Add2(): remote icon specified but couldn't start icon download", LEMON_LEVEL_WARNING

        End If

    End If


    ' /* R2.4: low-priority notifications cannot be sticky */

    If (Info.Priority < 0) And (Info.Timeout = 0) Then
        Info.Timeout = Val(g_ConfigGet("default_duration"))

    ElseIf Info.Timeout = -1 Then
        ' /* V41: if timeout is -1 then use global default timeout */
        Info.Timeout = Val(g_ConfigGet("default_duration"))

    End If


    ' /*
    '   R2.4 DR8: presence management, as follows:
    '
    '   Make Sticky takes priority over Do Not Disturb Mode, which in turn take priority over Away Mode
    '
    '   End result is:
    '                       Low         Normal      High
    '                       ----------------------------------
    '   Sticky              Sticky      Sticky      Sticky
    '   Busy (DND=true)     Discarded   <user>      Displayed
    '   Away                Discarded   <user>      Missed
    '
    '   Items marked <user> are user configurable.
    '
    ' */

    Debug.Print "++ current presence: away=" & g_HexStr(gAwayFlags) & " dnd=" & gDNDMode

    ' /* sticky snarls take precendence, then DND mode setting, then Away mode setting */

Dim nPMode As SP_PRESENCE_ACTIONS

    If g_ConfigGet("sticky_snarls") = "1" Then

        ' /* R2.4 DR8: *all* notifications are sticky */

        Info.Timeout = 0

    ElseIf gDNDMode Then

        ' /* Do Not Disturb (Busy) mode */

        If Info.Priority < 0 Then
            ' /* low priority = ignored */
            Exit Function

        ElseIf Info.Priority = 0 Then
            ' /* determine what to do with normal priority notifications */
            nPMode = Val(g_ConfigGet("busy_mode"))

        End If

    ElseIf g_IsAway() Then
        
        ' /* Away mode */

        If Info.Priority < 0 Then
            ' /* low priority = ignored */
            Exit Function

        ElseIf Info.Priority = 0 Then
            ' /* determine what to do with normal priority notifications */
            nPMode = Val(g_ConfigGet("away_mode"))

        Else
            ' /* high priority = logged as missed irrespective of class setting */
            uAddToMissedList Info
            Add = Info.Token

        End If
    End If

    Debug.Print "mode action = " & nPMode

    If nPMode <> 0 Then

        ' /* we have something specific to do with the notification */

        Select Case nPMode

        Case SP_LOG_AS_MISSED
            ' /* put in missed list (unless class setting says otherwise) and return the token */
            If Info.ClassObj.GetValue("log_during_dnd") = "1" Then
                uAddToMissedList Info
                Add = Info.Token

            End If

            Exit Function

        Case SP_MAKE_STICKY
            Info.Timeout = 0

        Case SP_DO_NOTHING
            ' /* bin it */
            Exit Function

        Case SP_DISPLAY
            ' /* allow it through */

        End Select

    End If

'    ' /* while DnD is in effect, normal priority notifications are added
'    '    to the missed list (unless the user has blocked them); high
'    '    priority notifications are still displayed */
'
'    If (gDNDMode) And (Info.Priority < 1) Then
'
'        ' /* R2.31: only add if "Log During DND" is enabled for the notification class */
'
'        gSetLastError SNARL_ERROR_DO_NOT_DISTURB
'
'        If (Info.Priority = 0) And (Info.ClassObj.GetValue("log_during_dnd") = "1") Then
'            uAddToMissedList Info
'            Add = Info.Token
'            Exit Function
'
'        Else
'            ' /* otherwise, it's discarded */
'            Add = 0
'            Exit Function
'
'        End If
'
'    End If
'
'
'    ElseIf (g_IsAway()) And (Info.Priority > 0) Then
'        ' /* R2.4 DR8: if away, force timeout to 0 for priority notifications */
'        Info.Timeout = 0
'
'    ElseIf Info.Timeout = -1 Then
'        ' /* V41: if timeout is -1 then use global default timeout */
'        Info.Timeout = Val(g_ConfigGet("default_duration"))
'
'    ElseIf (Info.Priority < 0) And (Info.Timeout = 0) Then
'        ' /* R2.4: low-priority notifications cannot be sticky */
'        Info.Timeout = Val(g_ConfigGet("default_duration"))
'
'    End If


    ' /* == if we're here we need to create an on-screen instance of the notification == */


    If Info.Priority < 0 Then

        ' /* R2.4: for low priority notifications only one may be visible
        '    at any one time, so zap any existing */

        If Me.IsVisible(gCurrentLowPriorityId, "", "", "") Then _
            Me.Hide gCurrentLowPriorityId, "", "", ""

        


'            ' /* we need to queue it */
'
'            mQueueCount = mQueueCount + 1
'            ReDim Preserve mQueue(mQueueCount)
'
'            ' /* if the token hasn't been pre-allocated set it here so the sender
'            '    can still manipulate the content */
'
'            If Info.Token = 0 Then
'                mToken = mToken + 1
'                Info.Token = mToken
'
'            End If
'
'            ' /* queue it */
'
'            With mQueue(mQueueCount)
'                Set .Class = Class
'                LSet .Detail = Info
'
'            End With
'
'            ' /* update the count badge on the existing notification */
'
'            uUpdateQueueCount
'
'            Add = Info.Token
'            Exit Function

    End If

    ' /* R2.4: merging */

Dim i As Long

    If (Info.IntFlags And S_NOTIFICATION_MERGE) Then
        ' /* is there an existing notification we can merge with? */
        i = uFindMergeCandidate(Info.ClassObj, Info.Title)
        If i Then
            ' /* update content */
            mItem(i).Window.MergeWith Info
            gSetLastError SNARL_WAS_MERGED                      ' // set LastError to indicate the merge
            Add = mItem(i).Window.Id
            Exit Function

        End If
    End If

    ' /* R2.4: overflow trimming */

    Info.Title = g_TrimLines(Info.Title)
    Info.Text = g_TrimLines(Info.Text)

    ' /* no style name?  that's okay, use the default */

    If Info.StyleName = "" Then
        Info.StyleName = style_GetStyleName(g_ConfigGet("default_style"))
        Info.SchemeName = style_GetSchemeName(g_ConfigGet("default_style"))

    End If

Dim psw As CSnarlWindow
Dim ar As BRect
Dim p As Long
Dim hr As M_RESULT

    ' /* get the appropriate instance for the notification before we do anything else
    '    we must pass a valid style instance to CSnarlWindow, even if the style
    '    roster is invalid */

Dim pStyle As TStyle

    ' /* get the style from the roster here - this is guaranteed to return a usable TStyle */

    uSafeGetStyle Info.StyleName & "/" & Info.SchemeName, pStyle

    ' /* get the path to the style - we need this for style-provided icons and sounds */

'    Info.IconPath = uTranslateIconPath(Info.IconPath, pStyle.Path)

    ' /* if S_STYLE_WANT_APP_NAME is set, prefix title with the sending app name */

    If (pStyle.Flags And S_STYLE_WANT_APP_NAME) Then
        g_Debug "TNotificationRoster.Add(): style wants application name"
        Info.Title = Info.ClassObj.App.Name & "#?" & Info.Title

    End If

Dim psi As IStyleInstance

    Set psi = pStyle.CreateInstance()
    If (psi Is Nothing) Then _
        Set psi = New TInternalStyle                ' // fail-safe

    ' /* if S_STYLE_IS_WINDOWLESS is set, just pass it to the handler */
    ' /* TO-DO: do we still need to record it? what about our return value? */

    If (pStyle.Flags And S_STYLE_IS_WINDOWLESS) Then
        g_Debug "TNotificationRoster.Add(): style is handling notification display"
        psi.UpdateContent pStyle.CreateNotificationInfo(Info.Title, Info.Text, Info.IconPath, Info.SchemeName, IIf(Info.Priority > 0, S_NOTIFICATION_IS_PRIORITY, 0))
        Add = -1                ' // TEMPORARY while we figure out what best to return here
        Exit Function

    End If

    ' /* is it a single instance style?  if so, we need to check to see if we already have a
    '    a notification of this style on-screen */

    If (pStyle.Flags And S_STYLE_SINGLE_INSTANCE) Then
        p = uFindbyStyleName(Info.StyleName)
        If p Then
            ' /* notification found, so decide what to do next... */
            If (pStyle.Flags And S_STYLE_DONT_QUEUE) Then
                ' /* don't queue notifications - so remove the existing one and create new */
                mItem(p).Window.Hide

            Else
                ' /* TO DO: add to queue */
                Exit Function

            End If
        End If
    End If

    ' /* if the token hasn't been pre-allocated set it here */

    If Info.Token = 0 Then
        mToken = mToken + 1
        Info.Token = mToken

    End If

Dim dwStyleFlags As S_STYLE_FLAGS

    dwStyleFlags = S_STYLE_MULTI_INSTANCE

    If Not (pStyle Is Nothing) Then _
        dwStyleFlags = pStyle.Flags




'    Info.StyleToUse = style_GetSchemeName(Info.StyleToUse)



    Set psw = New CSnarlWindow
    hr = psw.Create(Info, psi, dwStyleFlags, pStyle.Path)

'MsgBox Info.StyleToUse
    
    If hr <> M_OK Then
        ' /* TO-DO: Create() should probably return a more useful set of errors rather than just M_OK */
        g_Debug "TNotificationRoster.Add(): couldn't create window", LEMON_LEVEL_CRITICAL
        Add = 0
        Exit Function

    End If

    ' /* figure out its location on screen */

    If Not uGetNotificationPosition(psw, Info.Position, ar, i) Then
        g_Debug "TNotificationRoster.Add(): couldn't display notification", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_COULD_NOT_DISPLAY
        
        ' /* V42: couldn't find suitable position - mostly likely no room left on-screen */

        If Info.Priority >= 0 Then
            ' /* normal and high priority notifications are added to missed list */

            uAddToMissedList Info
            Add = Info.Token

        Else
            Add = 0

        End If

        Exit Function

    End If

    g_Debug "TNotificationRoster.Add(): placing at " & ar.Left & "," & ar.Top & "-" & ar.Right & "," & ar.Bottom

    ' /* add it to our list */

    mItems = mItems + 1
    ReDim Preserve mItem(mItems)
    With mItem(mItems)
        Set .Window = psw
        .Column = i
'        .hWndReply = Info.hWndReply
        ' /* for V40 (V39.13) */
'        .pid = Info.pid

'        If Not (Class Is Nothing) Then
''            Set .ClassObj = Info.ClassObj
'            .AppToken = Info.ClassObj.App.Token
'
'        End If

'        MsgBox "added: " & Info.StyleToUse

        ' /* V39.23 */
'        .Style = style_GetStyleName(Info.StyleToUse)

        ' /* R2.31 */
        .StyleName = Info.StyleName
        .SchemeName = Info.SchemeName

        ' /* R2.4 DR7 */
        LSet .Info = Info

        .Window.Window.MoveTo ar.Left, ar.Top
        .Window.Show

        Add = Info.Token

    End With

    ' /* R2.4: if this was a low priority notification, remember the token */

    If Info.Priority < 0 Then _
        gCurrentLowPriorityId = Add

    ' /* R2.4 DR8: last thing to do - play the sound */

Dim szSound As String

    If g_ConfigGet("enable_sounds") = "1" Then
        szSound = trim(Info.SndFile)
        If g_ConfigGet("use_style_sounds") = "1" Then
            If (szSound = "") Or ((szSound <> "") And (g_ConfigGet("prefer_style_sounds") = "1")) Then _
                szSound = ""        ' // get_style_sound()

        End If

        ' /* no specific or style-provided sound? */

        If szSound = "" Then _
            szSound = IIf(Info.Priority > 0, g_ConfigGet("default_priority_sound"), g_ConfigGet("default_notification_sound"))

        ' /* play what we have - if we have something */

        g_QuickPlaySound szSound

    End If

'    If Info.SndFile <> "" Then
'        If g_SafeLeftStr(Info.SndFile, 1) = "+" Then
'            ' /* system sound */
'            Info.SndFile = g_SafeRightStr(Info.SndFile, Len(Info.SndFile) - 1)
'            g_Debug "TNotificationRoster.Add2(): playing system sound '" & Info.SndFile & "'...", LEMON_LEVEL_INFO
'            If Not g_PlaySound(Info.SndFile, SND_ALIAS Or SND_NOSTOP Or SND_NOWAIT Or SND_ASYNC) Then _
'                g_Debug "TNotificationRoster.Add2(): PlaySound() failed", LEMON_LEVEL_CRITICAL
'
'        Else
'            ' /* file */
'            g_Debug "TNotificationRoster.Add2(): playing sound '" & Info.SndFile & "'...", LEMON_LEVEL_INFO
'            If Not g_PlaySound(Info.SndFile, SND_FILENAME Or SND_NOSTOP Or SND_NOWAIT Or SND_ASYNC) Then _
'                g_Debug "TNotificationRoster.Add2(): PlaySound() failed", LEMON_LEVEL_CRITICAL
'
'        End If
'    End If

End Function

Public Function Remove(ByVal Id As Long) As Boolean
Dim i As Long

    i = uFind(Id, "", "", "")
    If i Then _
        Remove = uRemove(i)

End Function

Private Function uRemove(ByVal Index As Long) As Boolean

    If (Index < 1) Or (Index > mItems) Then
        g_Debug "TNotificationRoster.uRemove(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim i As Long

'    Debug.Print "-- removing index " & Index & " --"

    Set mItem(Index).Window = Nothing
    If Index < mItems Then
        For i = Index To (mItems - 1)
            LSet mItem(i) = mItem(i + 1)

        Next i

    End If

    mItems = mItems - 1
    ReDim Preserve mItem(mItems)
    uRemove = True

'    If mItems Then
'        For i = 1 To mItems
'            With mItem(i).Window.Frame
'                Debug.Print "** " & .Left & "," & .Top & "-" & .Right & "," & .Bottom
'
'            End With
'        Next i
'    End If

End Function

Private Function uFind(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Long

    ' /* searches visible (on-screen) notifications and returns *index*
    '    of specified notification or zero if not found/error */

    gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND
    If (mItems = 0) Then _
        Exit Function

Static i As Long

    If Token <> 0 Then
        ' /* search on token */
        For i = 1 To mItems
            If mItem(i).Window.Id = Token Then
                uFind = i
                Exit For

            End If
        Next i

    ElseIf UID <> "" Then
        ' /* search on UID/App-Sig pair */
        For i = 1 To mItems
            If (mItem(i).Info.ClassObj.App.Signature = AppSig) And (mItem(i).Info.CustomUID = UID) Then
                uFind = i
                Exit For

            End If
        Next i
    End If

    ' /* found? */

    If uFind = 0 Then _
        Exit Function

    ' /* check password */

    If mItem(i).Info.ClassObj.App.Password = Password Then
        gSetLastError SNARL_SUCCESS

    Else
        gSetLastError SNARL_ERROR_AUTH_FAILURE

    End If

End Function

Private Function uFindInMissedList(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Long

    ' /* searches missed notifications and returns *index* of
    '    specified notification or zero if not found/error */

    gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND
    If (mMissedCount = 0) Then _
        Exit Function

Static i As Long

    If Token <> 0 Then
        ' /* search on token */
        For i = 1 To mMissedCount
            If mMissed(i).Token = Token Then
                uFindInMissedList = i
                Exit For

            End If
        Next i

    ElseIf UID <> "" Then
        ' /* search on UID/App-Sig pair */
        For i = 1 To mItems
            If (mMissed(i).ClassObj.App.Signature = AppSig) And (mMissed(i).CustomUID = UID) Then
                uFindInMissedList = i
                Exit For

            End If
        Next i
    End If

    ' /* found? */

    If uFindInMissedList = 0 Then _
        Exit Function

    ' /* check password */

    If mMissed(i).ClassObj.App.Password = Password Then
        gSetLastError SNARL_SUCCESS

    Else
        gSetLastError SNARL_ERROR_AUTH_FAILURE

    End If

End Function

Public Function Hide(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Boolean
Static i As Long

    i = uFind(Token, UID, AppSig, Password)
    If i Then
        mItem(i).Window.Hide
        gSetLastError SNARL_SUCCESS
        Hide = True

    Else

        i = uFindInMissedList(Token, UID, AppSig, Password)
        If i Then
            g_Debug "TNotificationRoster.Hide(): " & g_HexStr(Token) & " is in missed list"
            uRemoveFromMissedList i

        Else
            g_Debug "TNotificationRoster.Hide(): " & g_HexStr(Token) & " not found", LEMON_LEVEL_WARNING
            gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND

        End If

    End If

End Function

Public Function IsVisible(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Boolean

    ' /* uFind sets lasterror for us */
    IsVisible = (uFind(Token, UID, AppSig, Password) <> 0)

End Function

Public Function WasMissed(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Long

    ' /* uFindInMissedList sets lasterror for us */
    WasMissed = CLng(uFindInMissedList(Token, UID, AppSig, Password))

End Function

'Public Function Update(ByVal Token As Long, ByVal Title As String, ByVal Text As String, ByVal IconPath As String, Optional ByVal Content As String) As M_RESULT
'Dim i As Long
'
'    i = uFind(Token)
'    If i = 0 Then
'        g_Debug "TNotificationRoster.Update(): #" & g_HexStr(Token) & " not found", LEMON_LEVEL_WARNING
'        gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND
'        Update = M_NOT_FOUND
'        Exit Function
'
'    End If
'
'Dim pInfo As T_NOTIFICATION_INFO
'
'    ' /* don't do any special pre-processing here - this is handled by CSnarlWindow->uUupdateStyleContent() */
'
'    With pInfo
'
'        ' /* R2.4: process text and title */
'
'        Title = Replace$(Title, "\+", mItem(i).Window.Title)        ' // specific to Update() only
'        Text = Replace$(Text, "\+", mItem(i).Window.Text)           ' // sepcific to Update() only
'
'        .Title = g_TrimLines(Title)
'        .Text = g_TrimLines(Text)
'
'        ' /* R2.3: get the correct icon taking the current style into consideration */
'        .IconPath = IconPath '// uTranslateIconPath(IconPath, mItem(i).Window.StylePath)
'        .SchemeName = Chr$(255)         ' // means "use existing"
'
'        ' /* R2.4: store original content */
'        .OriginalContent = Content
'
'    End With
'
'    gSetLastError SNARL_SUCCESS
'    Update = mItem(i).Window.Update2(pInfo)
'
'End Function

Private Function uIsFreeSpace(ByRef aRect As BRect) As Boolean
Static i As Long

    ' /* return True if passed rect doesn't overlap an existing notification */

    uIsFreeSpace = True

    If mItems = 0 Then _
        Exit Function

'    g_Debug "uIsFreeSpace(): testing " & aRect.Left & "," & aRect.Top & "-" & aRect.Right & "," & aRect.Bottom

    For i = 1 To mItems
        If mItem(i).Window.Frame.Intersects(aRect) Then
'            With mItem(i).Window.Frame
'                g_Debug "blocked by item " & CStr(i) & " at " & .Left & "," & .Top & ";" & .Right & "," & .Bottom
'
'            End With

            uIsFreeSpace = False
            Exit Function

        End If
    Next i

'    With aRect
'        g_Debug "TNotificationRoster.uIsFreeSpace(): " & .Left & "," & .Top & "-" & .Right & "," & .Bottom & " is okay"
'
'    End With

End Function

Private Function uFindExtent(ByVal Column As Integer, ByVal RightToLeft As Boolean) As Long

    If mItems = 0 Then
        g_Debug "TNotificationRoster.uFindExtent(): no notifications visible"
        Exit Function

    End If

Dim x As Long
Dim i As Long

    x = IIf(RightToLeft, 8192, 0)

    For i = 1 To mItems
        With mItem(i)
            If .Column = Column Then
                ' /* in the correct column */
                If RightToLeft Then
                    If .Window.Frame.Left < x Then _
                        x = .Window.Frame.Left

                Else
'                    Debug.Print "!! " & .Frame.Left & " / " & .Frame.Right

                    If .Window.Frame.Right > x Then _
                        x = .Window.Frame.Right

                End If

            End If
        End With

    Next i

    uFindExtent = x
    g_Debug "TNotificationRoster.uFindExtent(): returning " & x

End Function

Public Sub CloseMultiple(ByVal AppToken As Long)
Dim n As Integer
Dim i As Long

    If AppToken = 0 Then
        ' /* close them all */
        g_Debug "TNotificationRoster.CloseMultiple(): no hWndReply - closing all..."

        If mItems Then
            ' /* fade all notifications out together */
            For n = 255 To 0 Step -(Val(g_ConfigGet("step_size")) * 8)
                For i = 1 To mItems
                    mItem(i).Window.bSetAlpha n

                Next i
            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mItems
                mItem(i).Window.bSetAlpha 0

            Next i
        
            ' /* now zap them */

            For i = mItems To 1 Step -1
                If Not (mItem(i).Window Is Nothing) Then _
                    mItem(i).Window.Remove

            Next i
        End If
    
        ReDim mItem(0)
        mItems = 0

    Else
        ' /* only close snarls with same creating app */

        g_Debug "CloseMultiple(): closing all Snarls from " & g_HexStr(AppToken) & "..."

        If mItems Then
            ' /* fade all related notifications out together */
            For n = 255 To 0 Step -(Val(g_ConfigGet("step_size")) * 8)
                For i = 1 To mItems
                    With mItem(i)
                        If .Info.ClassObj.App.Token = AppToken Then _
                            .Window.bSetAlpha n

                    End With
                Next i
            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mItems
                With mItem(i)
                    If .Info.ClassObj.App.Token = AppToken Then _
                        .Window.bSetAlpha 0

                End With
            Next i
            
            For i = mItems To 1 Step -1
                If mItem(i).Info.ClassObj.App.Token = AppToken Then
                    If Not (mItem(i).Window Is Nothing) Then _
                        mItem(i).Window.Remove

                    uRemove i

                End If
            Next i
        End If
    End If

End Sub

Public Function SetAttribute(ByVal Id As Long, ByVal Attr As Long, ByVal Value As String) As M_RESULT
Dim pw As CSnarlWindow
Dim i As Long

    i = uFind(Id, "", "", "")
    If i = 0 Then
        g_Debug "TNotificationRoster.SetAttribute(): notification 0x" & g_HexStr(Id) & " not found", LEMON_LEVEL_CRITICAL
        SetAttribute = M_NOT_FOUND
        Exit Function

    End If

    SetAttribute = M_FAILED
    Set pw = mItem(i).Window
    If (pw Is Nothing) Then _
        Exit Function

    ' /* found the notification, now change the value */

    SetAttribute = M_OK

    Select Case Attr

    Case SNARL_ATTRIBUTE_TITLE
       SetAttribute = Me.Update(Id, g_newBPackedData("title::" & Value))

    Case SNARL_ATTRIBUTE_TEXT
       SetAttribute = Me.Update(Id, g_newBPackedData("text::" & Value))

    Case SNARL_ATTRIBUTE_ICON
       SetAttribute = Me.Update(Id, g_newBPackedData("icon::" & Value))

    Case SNARL_ATTRIBUTE_TIMEOUT
       SetAttribute = pw.SetTimeout(Val(Value))

    Case SNARL_ATTRIBUTE_SOUND
        g_Debug "TNotificationRoster.SetAttribute(): SNARL_ATTRIBUTE_SOUND cannot be changed", LEMON_LEVEL_WARNING
        SetAttribute = M_NOT_IMPLEMENTED

    Case SNARL_ATTRIBUTE_ACK
        SetAttribute = pw.SetAck(Value)

    Case SNARL_ATTRIBUTE_MENU
        SetAttribute = pw.SetMenu(Value)

    Case Else
        g_Debug "TNotificationRoster.SetAttribute(): unknown Attribute '" & Attr & "'", LEMON_LEVEL_CRITICAL
        SetAttribute = M_INVALID_ARGS

    End Select

End Function

'Private Function uRectIntersect(ByRef aRect1 As BRect, ByRef aRect2 As BRect) As Boolean
'Dim rc1 As RECT
'Dim rc2 As RECT
'Dim rc3 As RECT
'
'    With aRect1
'        rc1.Left = .Left
'        rc1.Top = .Top
'        rc1.Right = .Right
'        rc1.Bottom = .Bottom
'
'    End With
'
'    With aRect2
'        rc2.Left = .Left
'        rc2.Top = .Top
'        rc2.Right = .Right
'        rc2.Bottom = .Bottom
'
'    End With
'
'    uRectIntersect = CBool(IntersectRect(rc3, rc1, rc2))
'
'End Function

Public Sub ShowMissedPanel()

    If (mMissedPanel Is Nothing) Then _
        Set mMissedPanel = New TMissedNotificationsPanel

    ResetMissedCount
    mMissedPanel.Go

End Sub

Public Function ActualMissedCount() As Long

    ActualMissedCount = mActualMissed

End Function

Public Sub ResetMissedCount()

    mActualMissed = 0
    frmAbout.bMissedNotificationsChanged

End Sub

Public Function MissedCount() As Long

    MissedCount = mMissedCount

End Function

'Friend Function MissedAt(ByVal Index As Long, ByRef Info As G_NOTIFICATION_CONTENT) As Boolean
'
'    If (Index > 0) And (Index <= mMissedCount) Then
'        LSet Info = mMissed(Index)
'        MissedAt = True
'
'    End If
'
'End Function

Friend Function MissedAt(ByVal Index As Long, ByRef Info As T_NOTIFICATION_INFO) As Boolean

    If (Index > 0) And (Index <= mMissedCount) Then
        LSet Info = mMissed(Index)
        MissedAt = True

    End If

End Function

Public Sub ClearMissed()

    ' /* clear the array */

    mMissedCount = 0
    ReDim mMissed(0)

    ' /* reset the missed count */

    ResetMissedCount

End Sub

Public Function HitTest(ByVal x As Long, ByVal y As Long) As Long

    If mItems = 0 Then _
        Exit Function

Dim px As BPoint
Dim i As Long

    Set px = new_BPoint(x, y)
    For i = mItems To 1 Step -1
'    Debug.Print "hittest: " & mItem(i).Style & " " & mItem(i).Frame.Left & " " & mItem(i).Frame.Right & " " & mItem(i).Frame.Top & " " & mItem(i).Frame.Bottom
        If mItem(i).Window.Frame.Contains(px) Then
            HitTest = i
            Exit Function

        End If
    Next i

End Function

Public Function NotificationAt(ByVal Index As Long) As CSnarlWindow

    If (Index > 0) And (Index <= mItems) Then _
        Set NotificationAt = mItem(Index).Window

End Function

Private Function uGetClassFlagsStr(ByRef Class As TAlert, ByVal ForcePriority As Boolean) As String
'Dim dw As E_ALERT_CLASS_FLAGS
'
'    If Not (Class Is Nothing) Then
'        dw = Class.GetFlags()
'
'        If (dw And E_CLASS_PRIORITY_NOTIFICATION) Then _
'            uGetClassFlagsStr = uGetClassFlagsStr & "p"
'
'        If (dw And E_CLASS_NOTIFICATION_ONLY) Then _
'            uGetClassFlagsStr = uGetClassFlagsStr & "n"
'
'    End If
'
'    If ForcePriority Then _
'        uGetClassFlagsStr = uGetClassFlagsStr & "p"

End Function

'Private Function uTranslateIconPath(ByVal Icon As String, ByVal StylePath As String) As String
'Dim pbm As mfxBitmap
'Dim pIcon As BIcon
'Dim sz As String
'Dim dw As Long
'Dim i As Long
'
'
'    If g_SafeLeftStr(Icon, 1) = "!" Then
'        ' /* convert the icon into it's corresponding file */
'        Icon = g_SafeRightStr(Icon, Len(Icon) - 1) & ".png"
'
'        ' /* default location */
'
'        uTranslateIconPath = g_MakePath(App.Path) & "etc\default_theme\icons\" & Icon
'
'        ' /* here, 'StylePath' sould be either empty or contain the full path to the
'        '    current style.  If it's the former we use the built-in icon set; if
'        '    it's the latter we try to get the icon from the style */
'
'        If (g_ConfigGet("use_style_icons") = "1") And (StylePath <> "") Then
'            ' /* try to get it from the style... */
'            sz = g_MakePath(StylePath) & "icons\" & Icon
'            If g_Exists(sz) Then _
'                uTranslateIconPath = sz
'
'        End If
'
'    ElseIf g_SafeLeftStr(Icon, 1) = "%" Then
'        ' /* whatever is after the % should be a valid HICON */
'        dw = Val(g_SafeRightStr(Icon, Len(Icon) - 1))
'        If dw Then
'            Set pbm = uGetIcon(dw)
'            If Not (pbm Is Nothing) Then
'                uTranslateIconPath = g_GetSafeTempIconPath()
'                pbm.Save uTranslateIconPath, "image/png"
'
'            End If
'
'        Else
'            g_Debug "TNotificationRoster.uTranslateIconPath(): bad HICON '" & Icon & "'", LEMON_LEVEL_WARNING
'
'        End If
'
'    ElseIf LCase$(g_GetExtension(Icon)) = "ico" Then
'        uTranslateIconPath = uLoadICO(Icon)
'
'    Else
'
'        uTranslateIconPath = Icon
'
'        ' /* windows icon? */
'
'        i = InStr(Icon, ",")
'        If i <> 0 Then
'            dw = Val(g_SafeRightStr(Icon, Len(Icon) - i))
'            Icon = g_SafeLeftStr(Icon, i - 1)
'            uTranslateIconPath = uGetBestIcon(Icon, dw)
'
'        End If
'    End If
'
'End Function
'
'Private Function uLoadICO(ByVal IconPath As String) As String
'Dim pbm As mfxBitmap
'Dim pIcon As BIcon
'
'    On Error Resume Next
'
'    With New BIconContent
'        If Not .LoadFromICO(IconPath) Then _
'            Exit Function
'
'        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pIcon) Then
'            Set pbm = create_bitmap_from_image(pIcon.Render)
'            uLoadICO = g_GetSafeTempIconPath()
'            pbm.Save uLoadICO, "image/png"
'
'        End If
'
'    End With
'
'End Function
'
'Private Function uGetBestIcon(ByVal IconPath As String, ByVal Index As Long) As String
'Dim pbm As mfxBitmap
'Dim pIcon As BIcon
'
'    On Error Resume Next
'
'    With New BIconContent
'        If Not .LoadFromResource(IconPath, Index) Then _
'            Exit Function
'
'        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pIcon) Then
'            Set pbm = create_bitmap_from_image(pIcon.Render)
'            uGetBestIcon = g_GetSafeTempIconPath()
'            pbm.Save uGetBestIcon, "image/png"
'
'        End If
'
'    End With
'
'End Function

Private Function uFindbyStyleName(ByVal Style As String) As Long

    If (mItems = 0) Or (Style = "") Then _
        Exit Function

Static i As Long

    ' /* case insensitive */

    Style = LCase$(Style)

    For i = 1 To mItems
'        Debug.Print "#" & CStr(i) & " >> " & mItem(i).StyleName
        If LCase$(mItem(i).StyleName) = Style Then
            uFindbyStyleName = i
            Exit Function

        End If
    Next i

End Function

Private Function uGetNotificationPosition(ByRef Window As CSnarlWindow, ByVal Position As E_START_POSITIONS, ByRef Frame As BRect, ByRef Column As Long) As Boolean

    If (Window Is Nothing) Then _
        Exit Function

Dim pv As mfxView

    Set pv = Window.GetView()
    If (pv Is Nothing) Then _
        Exit Function

    Column = 1

Dim ps As SIZEL

    With pv
        ps.cx = .Width
        ps.cy = .Height

    End With

    ' /* TODO: needs to take into account styles which veto the dropshadow */

    ps.cx = ps.cx + (2 * Val(g_ConfigGet("margin_spacing")))
    ps.cy = ps.cy + (2 * Val(g_ConfigGet("margin_spacing")))

Dim rcwa As RECT
Const MARGIN = 4

    g_GetWorkArea rcwa
    g_RectInset rcwa, MARGIN, MARGIN

    ' /* determine starting position */

    Set Frame = new_BRect(0, 0, ps.cx - 1, ps.cy - 1)

    If Position = E_START_DEFAULT_POS Then _
        Position = Val(g_ConfigGet("default_position"))

    Select Case Position

    Case E_START_TOP_LEFT
        Frame.OffsetBy rcwa.Left, rcwa.Top

    Case E_START_TOP_RIGHT
        Frame.OffsetBy rcwa.Right - ps.cx, rcwa.Top

    Case E_START_BOTTOM_LEFT
        Frame.OffsetBy rcwa.Left, rcwa.Bottom - ps.cy

    Case Else
        ' /* bottom right */
        Frame.OffsetBy rcwa.Right - ps.cx, rcwa.Bottom - ps.cy

    End Select

    ' /* loop looking for free space */

Dim p As Long
Const STEP_SIZE = 8

    Do While Not uIsFreeSpace(Frame)

        If (Position = E_START_TOP_LEFT) Or (Position = E_START_TOP_RIGHT) Then
            ' /* notifications run from top to bottom */

            Frame.OffsetBy 0, STEP_SIZE
            If Frame.Bottom > rcwa.Bottom Then
                ' /* notification would extend past bottom of work area so move to top of next column... */

                Frame.Top = rcwa.Top
                Frame.Bottom = Frame.Top + ps.cy

                If Position = E_START_TOP_RIGHT Then
                    ' /* ...which is to the left */
                    p = uFindExtent(Column, True) - STEP_SIZE
                    Frame.Right = p
                    Frame.Left = Frame.Right - ps.cx

                    If Frame.Left < rcwa.Left Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                Else
                    ' /* ...which is to the right */
                    p = uFindExtent(Column, False) + STEP_SIZE
                    Frame.Left = p
                    Frame.Right = Frame.Left + ps.cx

                    If Frame.Right > rcwa.Right Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                End If

                Column = Column + 1     ' // column count always increments

            End If

        Else
            ' /* notifications run from bottom to top */

            Frame.OffsetBy 0, -STEP_SIZE
            If Frame.Top < rcwa.Top Then
                ' /* notification would extend past top of work area so move to bottom of next column... */

                Frame.Top = rcwa.Bottom - ps.cy
                Frame.Bottom = Frame.Top + ps.cy

                If Position = E_START_BOTTOM_RIGHT Then
                    ' /* ...which is to the left */
                    p = uFindExtent(Column, True) - STEP_SIZE
                    Frame.Right = p
                    Frame.Left = Frame.Right - ps.cx

                    If Frame.Left < rcwa.Left Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                Else
                    ' /* ...which is to the right */
                    p = uFindExtent(Column, False) + STEP_SIZE
                    Frame.Left = p
                    Frame.Right = Frame.Left + ps.cx

                    If Frame.Right > rcwa.Right Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                End If

                Column = Column + 1

            End If

        End If
    Loop

    ' /* for 2.1+ style can veto position */

Dim pt As POINTAPI

    pt.x = Frame.Left
    pt.y = Frame.Top

    Window.AdjustPosition pt.x, pt.y

    Frame.Left = pt.x
    Frame.Top = pt.y

    uGetNotificationPosition = True

End Function

'Private Function uGetIcon(ByVal hIcon As Long) As mfxBitmap
'
'    On Error Resume Next
'
'Dim pi As BIcon
'
'    Err.Clear
'    Set pi = New BIcon
'    If Err.Number <> 0 Then
'        g_Debug "TNofiticationRoster.uGetIcon(): no icon.resource", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
'    If Not pi.SetFromHICON(hIcon) Then
'        g_Debug "TNofiticationRoster.uGetIcon(): bad icon handle", LEMON_LEVEL_CRITICAL
'        Exit Function
'
'    End If
'
'    Set uGetIcon = create_bitmap_from_image(pi.Render())
'
'End Function

Private Sub uSafeGetStyle(ByVal StyleToUse As String, ByRef Style As TStyle)
Dim pStyle As TStyle

    ' /* pre-create a new TStyle object - this will act as a fail-safe as it automatically
    '    initialises itself as the internal style */

    Set Style = New TStyle
    If (g_StyleRoster Is Nothing) Then _
        Exit Sub

    If g_StyleRoster.Find(style_GetStyleName(StyleToUse), pStyle) Then _
        Set Style = pStyle

End Sub

Private Sub uCopyTo(ByRef Info As T_NOTIFICATION_INFO, ByRef Class As TAlert)

    ' /* get the appropriate instance for the notification before we do anything else
    '    we must pass a valid style instance to CSnarlWindow, even if the style
    '    roster is invalid */

Dim pStyle As TStyle

    ' /* get the style from the roster here - this is guaranteed to return a usable TStyle */

    uSafeGetStyle Info.StyleName & "/" & Info.SchemeName, pStyle

    ' /* get the path to the style - we need this for style-provided icons and sounds */

    Info.IconPath = g_TranslateIconPath(Info.IconPath, pStyle.Path)

    ' /* if S_STYLE_WANT_APP_NAME is set, prefix title with the sending app name */

    If (pStyle.Flags And S_STYLE_WANT_APP_NAME) Then
        g_Debug "TNotificationRoster.uCopyTo(): style wants application name"
        Info.Title = Class.App.Name & "#?" & Info.Title

    End If

    ' /* this is just a safety check */

    If (pStyle.Flags And S_STYLE_IS_WINDOWLESS) = 0 Then
        Debug.Print "TNotificationRoster.uCopyTo(): '" & pStyle.Name & "' is not a windowless style"
        Exit Sub

    End If

Dim psi As IStyleInstance

    Set psi = pStyle.CreateInstance()
    If (psi Is Nothing) Then
        Debug.Print "TNotificationRoster.uCopyTo(): failed to create instance for '" & pStyle.Name & "'"
        Exit Sub

    End If

    g_Debug "TNotificationRoster.uCopyTo(): sending to '" & pStyle.Name & "'"
    psi.UpdateContent pStyle.CreateNotificationInfo(Info.Title, Info.Text, Info.IconPath, _
                                                    Info.SchemeName, _
                                                    IIf(Info.Priority > 0, S_NOTIFICATION_IS_PRIORITY, 0))

End Sub

'Private Sub uUpdateQueueCount()
'Dim i As Long
'
'    i = uFind(gCurrentLowPriorityId)
'    If i Then _
'        mItem(i).Window.QueueCountChanged
'
'End Sub
'
'Public Sub QueueNext()
'Dim pq As T_QUEUED
'Dim i As Long
'
'    If mQueueCount > 0 Then
'        LSet pq = mQueue(1)
'
'        If mQueueCount > 1 Then
'            For i = 1 To mQueueCount - 1
'                LSet mQueue(i) = mQueue(i + 1)
'
'            Next i
'
'        End If
'
'        mQueueCount = mQueueCount - 1
'        ReDim Preserve mQueue(mQueueCount)
'
'        Me.Add pq.Detail
'
'    End If
'
'End Sub
'
'Public Function QueueCount() As Long
'
'    QueueCount = mQueueCount
'
'End Function

Private Function uFindMergeCandidate(ByRef Class As TAlert, ByVal Title As String) As Long

    If (Class Is Nothing) Or (mItems = 0) Then _
        Exit Function

Dim i As Long

    For i = 1 To mItems
        If mItem(i).Window.IsMergeCandidate(Class, Title) Then
            uFindMergeCandidate = i
            Exit Function

        End If

    Next i

End Function

Public Function AddAction(ByVal Token As Long, ByRef Data As BPackedData) As Long

    ' /* param checking... */

    If (Data.ValueOf("label") = "") Or (Data.ValueOf("cmd") = "") Then
        g_Debug "TNotificationRoster.AddAction(): missing label and/or command", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    ' /* TO-DO: if notification is in missed list, need to update record */

Dim i As Long

    ' /* uFind sets lasterror for us */

    i = uFind(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"))
    If i Then
        ' /* still visible */
        If uAddAction(mItem(i).Info.Actions, Data.ValueOf("label"), Data.ValueOf("cmd")) Then
            mItem(i).Window.RethinkActions
            AddAction = -1

        End If

    Else
        ' /* to-do: check missed list */
        i = uFindInMissedList(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"))
        If i Then
            If uAddAction(mMissed(i).Actions, Data.ValueOf("label"), Data.ValueOf("cmd")) Then
                ' /* update missed panel */
                AddAction = -1

            End If
        End If
    End If

End Function

Private Function uAddAction(ByRef List As BTagList, ByVal Label As String, ByVal Command As String) As Boolean

    With List
        If .IndexOf(Label) = 0 Then
            ' /* add new */
            .Add new_BTagItem(Label, Command)

        Else
            ' /* if duplicate label, update command */
            .Update Label, Command

        End If

        ' /* Need to check for duplicate command and fail */

    End With

    uAddAction = True

End Function

Public Function ClearActions(ByVal Token As Long, ByRef Data As BPackedData) As Long
Static i As Long

    ' /* uFind sets lasterror for us */

    i = uFind(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"))
    If i Then
        If mItem(i).Info.Actions.CountItems <> 0 Then
            mItem(i).Info.Actions.MakeEmpty
            mItem(i).Window.RethinkActions

        End If

        ClearActions = -1

    Else
        i = uFindInMissedList(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"))
        If i Then
            Debug.Print "TNotificationRoster.ClearActions(): in missed list"
            mMissed(i).Actions.MakeEmpty
            '//update missed panel if open

            ClearActions = -1

        End If

    End If

End Function

Private Sub uAddToMissedList(ByRef Info As T_NOTIFICATION_INFO)

    If Info.Token = 0 Then
        ' /* allocate a token now */
        mToken = mToken + 1
        Info.Token = mToken

    End If

    mActualMissed = mActualMissed + 1
    mMissedCount = mMissedCount + 1
    ReDim Preserve mMissed(mMissedCount)

Dim i As Long

    If mMissedCount > 1 Then
        ' /* push existing enties out one */
        For i = mMissedCount To 2 Step -1
            LSet mMissed(i) = mMissed(i - 1)

        Next i

    End If

    ' /* add to head */
    LSet mMissed(1) = Info
    Set mMissed(1).Icon = style_MakeSquareImage(load_image_obj(g_TranslateIconPath(Info.IconPath, "")))

    If Not (mMissedPanel Is Nothing) Then _
        mMissedPanel.UpdateList

    frmAbout.bMissedNotificationsChanged

End Sub

Private Sub uRemoveFromMissedList(ByVal Index As Long)

    If (Index < 1) Or (Index > mMissedCount) Then _
        Exit Sub

Static i As Long

    If Index < mMissedCount Then
        For i = Index To (mMissedCount - 1)
            LSet mMissed(i) = mMissed(i + 1)

        Next i

    End If

    mMissedCount = mMissedCount - 1
    ReDim Preserve mMissed(mMissedCount)


    If Not (mMissedPanel Is Nothing) Then _
        mMissedPanel.UpdateList

'    frmAbout.bMissedNotificationsChanged

End Sub

Public Function Update(ByVal Token As Long, ByRef Args As BPackedData) As Boolean
Dim sz As String
Dim i As Long

    g_Debug "TNotificationRoster.Update()", LEMON_LEVEL_PROC_ENTER

    i = uFind(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"))
    If i Then
        g_Debug "found visible"

        With mItem(i).Info

            If Args.Exists("title") Then
                sz = Replace$(Args.ValueOf("title"), "\n", vbCrLf)
                sz = Replace$(sz, "\+", .Title)                         ' // specfic to Update() - replace '\+' with existing title
                .Title = g_TrimLines(sz)

            End If

            If Args.Exists("text") Then
                sz = Replace$(Args.ValueOf("text"), "\n", vbCrLf)
                sz = Replace$(sz, "\+", .Text)                          ' // specfic to Update() - replace '\+' with existing text
                .Text = g_TrimLines(sz)

            End If

            If Args.Exists("icon") Then _
                .IconPath = Args.ValueOf("icon")

            If Args.Exists("timeout") Then _
                .Timeout = g_SafeLong(Args.ValueOf("timeout"))          ' // g_SafeLong() returns 0 on error - not good...

        End With

        uFixupContent mItem(i).Info                     ' // re-creates T_NOTIFICATION->OriginalContent
        Update = mItem(i).Window.Update(mItem(i).Info)

    Else
        i = uFindInMissedList(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"))
        If i Then
            g_Debug "in missed list"
            ' /* TO-DO: update in missed list */
'            Update = True

        Else
            g_Debug "not found", LEMON_LEVEL_WARNING

        End If

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Public Function Merge(ByVal Token As Long, ByRef Args As BPackedData) As Boolean
Dim sz As String
Dim i As Long

    g_Debug "TNotificationRoster.Merge()", LEMON_LEVEL_PROC_ENTER

    i = uFind(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"))
    If i Then
        g_Debug "found visible"

        With mItem(i).Info

'            If Args.Exists("title") Then
'                sz = Replace$(Args.ValueOf("title"), "\n", vbCrLf)
'                sz = Replace$(sz, "\+", .Title)                         ' // specfic to Update() - replace '\+' with existing title
'                .Title = g_TrimLines(sz)
'
'            End If

            If Args.Exists("text") Then _
                .Text = g_TrimLines(.Text & " " & Replace$(Args.ValueOf("text"), "\n", vbCrLf))

            If Args.Exists("icon") Then _
                .IconPath = Args.ValueOf("icon")

'            If Args.Exists("timeout") Then _
'                .Timeout = g_SafeLong(Args.ValueOf("timeout"))          ' // g_SafeLong() returns 0 on error - not good...

        End With

        uFixupContent mItem(i).Info                     ' // re-creates T_NOTIFICATION->OriginalContent
        Merge = mItem(i).Window.Update(mItem(i).Info)

    Else
        i = uFindInMissedList(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"))
        If i Then
            g_Debug "in missed list"
            ' /* TO-DO: update in missed list */
'            Update = True

        Else
            g_Debug "not found", LEMON_LEVEL_WARNING

        End If

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Sub uFixupContent(ByRef Info As T_NOTIFICATION_INFO)

    If (Info.ClassObj Is Nothing) Then
        Debug.Print "########### uFixupContent missing class object ###########"
        Exit Sub

    End If

    With Info
        .Title = Replace$(.Title, "\n", vbCrLf)
        .Text = Replace$(.Text, "\n", vbCrLf)
    
    End With

Dim ppn As BPackedData

    ' /* new */
    Set ppn = New BPackedData

    With ppn
        .Add "id", Info.ClassObj.Name
        .Add "title", Info.Title
        .Add "text", Info.Text
        .Add "timeout", CStr(Info.Timeout)
        .Add "icon", Info.IconPath
        .Add "priority", CStr(Info.Priority)
        .Add "callback", Info.DefaultAck
        .Add "value", Info.Value

'        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
            .Add "merge", "1"

    End With

Dim ppd As BPackedData
Dim szn As String
Dim szv As String

    ' /* add in all other content */

    Set ppd = New BPackedData
    With ppd
        If .SetTo(Info.OriginalContent) Then
            .Rewind
            Do While .GetNextItem(szn, szv)
                If Not ppn.Exists(LCase$(szn)) Then _
                    ppn.Add szn, szv

            Loop
        End If
    End With

    Info.OriginalContent = ppn.AsString()

End Sub

Public Function TokenFromUID(ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Long
Static i As Long

    i = uFind(0, UID, AppSig, Password)
    If i Then
        TokenFromUID = mItem(i).Info.Token

    Else
        i = uFindInMissedList(0, UID, AppSig, Password)
        If i Then _
            TokenFromUID = mMissed(i).Token

    End If



'    ' /* returns token or zero if not found / not authorised */
'
'    gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND
'
'    Debug.Print mItems & " " & AppSig & " " & UID
'
'    If (mItems = 0) Or (AppSig = "") Or (UID = "") Then _
'        Exit Function
'
'Static i As Long
'
'    For i = 1 To mItems
'        If (mItem(i).Info.ClassObj.App.Signature = AppSig) And (mItem(i).Info.CustomUID = UID) Then
'            If mItem(i).Info.ClassObj.App.Password = Password Then
'                UIDToToken = mItem(i).Info.Token
'                gSetLastError SNARL_SUCCESS
'
'            Else
'                gSetLastError SNARL_ERROR_AUTH_FAILURE
'
'            End If
'
'            Exit Function
'
'        End If
'    Next i

End Function

