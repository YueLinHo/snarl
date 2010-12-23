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
    hWndReply As Long
    Window As CSnarlWindow
    Column As Integer           ' // used to help determine placement of other notifications
    pid As Long
'    Class As String
    ClassObj As TAlert          ' // R2.31: object is saved
    StyleName As String         ' // R2.31: expanded
    SchemeName As String        ' // R2.31: expanded
    AppToken As Long            ' // V40.47: creating app - all registered apps should now have a token

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
                mItem(i).Window.Zap

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

Friend Function Add(ByRef Class As TAlert, ByRef Info As T_NOTIFICATION_INFO) As Long

    gSetLastError SNARL41_ERROR_FAILED
    If g_IsQuitting Then _
        Exit Function               ' // shutting down - can't do anything

    ' /* log only? */

    If g_ConfigGet("log_only") = "1" Then
        g_WriteToLog Info.Title, Info.Text
        Add = -1                    ' // V40.7 - changed from M_OK to -1 (indicates success but no token available)
        gSetLastError SNARL41_SUCCESS
        Exit Function               ' // TO-DO: determine if we should return 0 and set last_error instead

    End If

    ' /* for anything else, 'Class' *must* be valid */

    If (Class Is Nothing) Then
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

    If (Class.GetValue("block_spam") = "1") And ((GetTickCount() - Class.LastNotificationGeneratedTick) < lQuantum) Then
        Debug.Print "!!! SPAM !!! " & CStr(lQuantum)

        ' /* R2.31 - for now, block and return a failure code.  For the future, it would be lovely to
        '    allow more functionality (e.g. update notification, append to notification, send to missed log, etc. */

        gSetLastError SNARL41_ERROR_FLOODING
        Exit Function

    End If

    ' /* only record the tick of the last _generated_ notification */

    Class.LastNotificationGeneratedTick = GetTickCount()

    gSetLastError SNARL41_SUCCESS

    ' /* R2.31 - copy these across from the class object */

    Info.Sender = Class.App.Name
    Info.Class = Class.Description

    ' /* R2.3: set datestamp */

    Info.DateStamp = Now()

    ' /* R2.3: send to any selected non-window styles */

Dim pCopyList As BPackedData
Dim pCopyInfo As T_NOTIFICATION_INFO
Dim itemName As String

    Set pCopyList = Class.ForwardList()
    If Not (pCopyList Is Nothing) Then
        With pCopyList
            .Rewind
            Do While .GetNextItem(itemName, "")
                LSet pCopyInfo = Info
                pCopyInfo.StyleName = style_GetStyleName(itemName)
                pCopyInfo.SchemeName = style_GetSchemeName(itemName)
                uCopyTo pCopyInfo, Class

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
            pDelayed.SetTo Info, Class

            mDelayedCount = mDelayedCount + 1
            ReDim Preserve mDelayed(mDelayedCount)
            Set mDelayed(mDelayedCount) = pDelayed      '// add to the queue

            Add = mToken                                '// return the token that was allocated
            Exit Function

        Else
            g_Debug "TNotificationRoster.Add2(): remote icon specified but couldn't start icon download", LEMON_LEVEL_WARNING

        End If

    End If


    ' /* Do Not Disturb mode */


Dim pIcon As mfxBitmap

    ' /* while DnD is in effect, normal priority notifications are added
    '    to the missed list (unless the user has blocked them); high
    '    priority notifications are still displayed */

    If (g_IsDNDModeEnabled()) And (Info.Priority < 1) Then

        ' /* R2.31: only add if "Log During DND" is enabled for the notification class */

        gSetLastError SNARL42_ERROR_DO_NOT_DISTURB

        If (Info.Priority = 0) And (Class.GetValue("log_during_dnd") = "1") Then

            uAddToMissedList Info

            ' /* special case: we return a valid token but we set LastError
            '    to ERROR_DO_NOT_DISTURB to indicate the notification wasn't
            '    actually displayed */

            Add = Info.Token
            Exit Function

        Else
            ' /* otherwise, it's discarded */
            Add = 0
            Exit Function

        End If

    End If


    ' /* R2.4: for low priority notifications only one may be visible
    '    at any one time, so zap any existing */

    If Info.Priority < 0 Then

        If sn41IsVisible(gCurrentLowPriorityId) Then _
            Me.Hide gCurrentLowPriorityId

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

    ' /* if we're here we need to create an on-screen instance of the notification */

    If (g_IsSticky()) And (Info.Priority > 0) Then
        ' /* R2.3: if sticky, force timeout to 0 for priority notifications */
        Info.Timeout = 0

    ElseIf Info.Timeout = -1 Then
        ' /* V41: (not sure it's documented) - if timeout is -1 then use global default timeout */
        Info.Timeout = Val(g_ConfigGet("default_duration"))

    ElseIf (Info.Priority < 0) And (Info.Timeout = 0) Then
        ' /* R2.4: low-priority notifications cannot be sticky */
        Info.Timeout = Val(g_ConfigGet("default_duration"))

    End If

    ' /* R2.4: merging */

Dim i As Long

    If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then
        ' /* is there an existing notification we can merge with? */
        i = uFindMergeCandidate(Class, Info.Title)
        If i Then
            ' /* update content */
            mItem(i).Window.MergeWith Info
            gSetLastError SNARL42_WAS_MERGED                    ' // set LastError to indicate the merge
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
        Info.Title = Class.App.Name & "#?" & Info.Title

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
'        MsgBox Info.StyleToUse

        p = uFindbyStyleName(Info.StyleName)
        If p Then

'            MsgBox p, , Info.StyleName

            ' /* notification found, so decide what to do next... */
            If (pStyle.Flags And S_STYLE_DONT_QUEUE) Then
                ' /* don't queue notifications - so update and return the existing id */

                mItem(p).Window.Hide

''                Info.StyleToUse = style_GetSchemeName(Info.StyleToUse)
'                mItem(p).Window.Update2 Info
'                Add = mItem(p).Window.Id
'                Exit Function

            Else
                ' /* TO DO: add to queue */
                Exit Function

            End If
        End If
    End If

    ' /* figure out sound to play */

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
    hr = psw.Create(Class, Info, psi, dwStyleFlags, pStyle.Path)
    
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
        gSetLastError SNARL42_ERROR_COULD_NOT_DISPLAY
        
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
        .hWndReply = Info.hWndReply
        ' /* for V40 (V39.13) */
        .pid = Info.pid

        If Not (Class Is Nothing) Then
            Set .ClassObj = Class
            .AppToken = Class.App.Token

        End If

'        MsgBox "added: " & Info.StyleToUse

        ' /* V39.23 */
'        .Style = style_GetStyleName(Info.StyleToUse)

        ' /* R2.31 */
        .StyleName = Info.StyleName
        .SchemeName = Info.SchemeName

        .Window.Window.MoveTo ar.Left, ar.Top
        .Window.Show

        Add = Info.Token

    End With

    ' /* R2.4: if this was a low priority notification, remember the token */

    If Info.Priority < 0 Then _
        gCurrentLowPriorityId = Add

End Function

Public Function Remove(ByVal Id As Long) As Boolean
Dim i As Long

    i = uFind(Id)
    If i Then _
        Remove = uRemove(i)

End Function

Private Function uRemove(ByVal Index As Long) As Boolean

    If (Index < 1) Or (Index > mItems) Then
        g_Debug "TNotificationRoster.uRemove(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim i As Long

    Debug.Print "-- removing index " & Index & " --"

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

Private Function uFind(ByVal Id As Long) As Long
Dim i As Long

    If mItems Then
        For i = 1 To mItems
            If mItem(i).Window.Id = Id Then
                uFind = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Function uFindInMissedList(ByVal Token As Long) As Long

    If mMissedCount = 0 Then _
        Exit Function

Static i As Long

    For i = 1 To mMissedCount
        If mMissed(i).Token = Token Then
            uFindInMissedList = i
            Exit Function

        End If
    Next i

End Function

Public Function Hide(ByVal Id As Long) As Boolean
Static i As Long

    i = uFind(Id)
    If i Then
        mItem(i).Window.Hide
        gSetLastError SNARL41_SUCCESS
        Hide = True

    Else

        i = uFindInMissedList(Id)
        If i Then
            g_Debug "TNotificationRoster.Hide(): '" & g_HexStr(Id) & " is in missed list"
            uRemoveFromMissedList i

        Else
            g_Debug "TNotificationRoster.Hide(): '" & g_HexStr(Id) & " not found", LEMON_LEVEL_WARNING
            gSetLastError SNARL41_ERROR_NOTIFICATION_NOT_FOUND

        End If

    End If

End Function

Public Function IsVisible(ByVal Id As Long) As Boolean

    g_Debug "TNotificationRoster.IsVisible(" & g_HexStr(Id) & ")", LEMON_LEVEL_PROC_ENTER

    If uFind(Id) > 0 Then
        g_Debug "found/visible"
        gSetLastError SNARL41_SUCCESS
        IsVisible = True

    ElseIf uFindInMissedList(Id) Then
        g_Debug "found/missed"
        gSetLastError SNARL41_SUCCESS
        IsVisible = False

    Else
        g_Debug "not found"
        gSetLastError SNARL41_ERROR_NOTIFICATION_NOT_FOUND
        IsVisible = False

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

'Public Function Update(ByVal Token As Long, ByVal Title As String, ByVal Text As String, ByVal IconPath As String, Optional ByVal Content As String) As M_RESULT
'Dim i As Long
'
'    i = uFind(Token)
'    If i = 0 Then
'        g_Debug "TNotificationRoster.Update(): #" & g_HexStr(Token) & " not found", LEMON_LEVEL_WARNING
'        gSetLastError SNARL41_ERROR_NOTIFICATION_NOT_FOUND
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
'    gSetLastError SNARL41_SUCCESS
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
                    mItem(i).Window.Zap

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
                        If .AppToken = AppToken Then _
                            .Window.bSetAlpha n

                    End With
                Next i
            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mItems
                With mItem(i)
                    If .AppToken = AppToken Then _
                        .Window.bSetAlpha 0

                End With
            Next i
            
            For i = mItems To 1 Step -1
                If mItem(i).AppToken = AppToken Then
                    If Not (mItem(i).Window Is Nothing) Then _
                        mItem(i).Window.Zap

                    uRemove i

                End If
            Next i
        End If
    End If

End Sub

Public Function SetAttribute(ByVal Id As Long, ByVal Attr As SNARL_ATTRIBUTES, ByVal Value As String) As M_RESULT
Dim pw As CSnarlWindow
Dim i As Long

    i = uFind(Id)
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

Public Function GetCount(ByVal pid As Long, ByVal ByClass As Boolean, ByVal Class As String) As Long
Dim i As Long

    g_Debug "TNotificationRoster.GetCount(pid=" & CStr(pid) & " ByClass=" & ByClass & " Class=" & Class & " ) [V40]", LEMON_LEVEL_PROC

    If mItems = 0 Then
        g_Debug "TNotificationRoster.GetCount(): no notifications"
        Exit Function

    End If

    For i = 1 To mItems

        g_Debug ">> " & CStr(mItem(i).pid) & " >> " & mItem(i).ClassObj.Name

        If mItem(i).pid = pid Then
            If ByClass Then
                If Class = mItem(i).ClassObj.Name Then _
                    GetCount = GetCount + 1

            Else
                GetCount = GetCount + 1

            End If
        End If

    Next i

End Function

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

Private Sub uUpdateQueueCount()
Dim i As Long

    i = uFind(gCurrentLowPriorityId)
    If i Then _
        mItem(i).Window.QueueCountChanged

End Sub

Public Sub QueueNext()
Dim pq As T_QUEUED
Dim i As Long

    If mQueueCount > 0 Then
        LSet pq = mQueue(1)

        If mQueueCount > 1 Then
            For i = 1 To mQueueCount - 1
                LSet mQueue(i) = mQueue(i + 1)

            Next i

        End If

        mQueueCount = mQueueCount - 1
        ReDim Preserve mQueue(mQueueCount)

        Me.Add pq.Class, pq.Detail

    End If

End Sub

Public Function QueueCount() As Long

    QueueCount = mQueueCount

End Function

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

    If Data.ValueOf("label") = "" Then
        g_Debug "TNotificationRoster.AddAction(): missing 'label' from action", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL41_ERROR_ARG_MISSING
        Exit Function

    End If

Dim i As Long

    i = uFind(Token)
    If i = 0 Then
        g_Debug "TNotificationRoster.AddAction(): token '" & g_HexStr(Token) & " not found", LEMON_LEVEL_WARNING
        gSetLastError SNARL41_ERROR_NOTIFICATION_NOT_FOUND
        Exit Function

    End If

    If Not mItem(i).Window.AddAction(Data.ValueOf("label"), Data.ValueOf("cmd")) Then
        gSetLastError SNARL41_ERROR_FAILED

    Else
        gSetLastError SNARL41_SUCCESS
        AddAction = -1

    End If

End Function

Public Function ClearActions(ByVal Token As Long) As Long
Dim i As Long

    i = uFind(Token)
    If i = 0 Then
        g_Debug "TNotificationRoster.AddAction(): token '" & g_HexStr(Token) & " not found", LEMON_LEVEL_WARNING
        gSetLastError SNARL41_ERROR_NOTIFICATION_NOT_FOUND
        Exit Function

    End If

    mItem(i).Window.ClearActions
    ClearActions = -1

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
    LSet mMissed(mMissedCount) = Info
    Set mMissed(mMissedCount).Icon = style_MakeSquareImage(load_image_obj(g_TranslateIconPath(Info.IconPath, "")))

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
Dim i As Long

    i = uFind(Token)
    If i Then
        g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' found"
        Update = mItem(i).Window.Update(Args)

    Else
        i = uFindInMissedList(Token)
        If i Then
            g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' is in missed list"

        Else
            g_Debug "TNotificationRoster.Update(): '" & g_HexStr(Token) & "' not found", LEMON_LEVEL_WARNING
            gSetLastError SNARL41_ERROR_NOTIFICATION_NOT_FOUND
            Update = False

        End If
    
    End If

End Function

