VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_CREATE = &H1
Private Const WM_DESTROY = &H2
'Private Const WM_CLOSE = &H10

Private Const TIMER_WAIT_FOR_OUTLOOK = 1
Private Const TIMER_WAIT_FOR_QUIT = 2
'Private Const TIMER_MAIL_CHECK = 3
Private Const TIMER_WAIT_FOR_INIT = 4

Private Const WM_TIMER = &H113
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_NOTIFICATION = &H402

Dim mhWnd As Long
Dim WithEvents myOutlook As Outlook.Application
Attribute myOutlook.VB_VarHelpID = -1
'Dim WithEvents myItems As Outlook.Items

Dim mToken As Long
'Dim WithEvents myPanel As BPrefsPanel
'Dim mConfig As BPersistentSettings
'Dim mContacts As Outlook.Items

Private Const CLASS_EMAIL = "mail"
Private Const CLASS_EMAIL_HIGH = "mail_high"
Private Const CLASS_EMAIL_LOW = "mail_low"

'Private Const CLASS_MEETING = "appt"
'Private Const CLASS_URGENT_MEETING = "aurg"
'Private Const CLASS_BORING_MEETING = "alow"
'
'Private Type T_ITEM
'    NotifyToken As Long
'    IsMeeting As Boolean
'    Item As Object
'    ItemId As String
'
'End Type
'
'Dim mItem() As T_ITEM
'Dim mItems As Long

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Dim pObj As Object
Dim n As Long

    On Error Resume Next

    Select Case uMsg
    Case WM_CREATE
        Form1.Add "WM_CREATE"
'        uGetConfig
        SetTimer hWnd, TIMER_WAIT_FOR_OUTLOOK, 2000, 0
        mhWnd = hWnd


    Case WM_DESTROY
        Form1.Add "WM_DESTROY"
'        If Not (myPanel Is Nothing) Then _
            myPanel.Quit

        uTidyUp


    Case WM_CLOSE
        PostQuitMessage 0


    Case WM_TIMER
        Select Case wParam
        Case TIMER_WAIT_FOR_OUTLOOK
'            Debug.Print "outlook.exe:" & g_IsProcessRunning(, "outlook.exe")
'            Debug.Print "snarl:" & snIsSnarlRunning()
            If (g_IsProcessRunning(, "outlook.exe")) And (snIsSnarlRunning()) Then
                g_Debug "got Outlook and Snarl, waiting..."
                KillTimer hWnd, wParam
                SetTimer hWnd, TIMER_WAIT_FOR_INIT, 1000, 0

            End If

        Case TIMER_WAIT_FOR_INIT
            ' /* settle time between outlook/snarl launching has completed */
            g_Debug "init wait completed..."
            KillTimer hWnd, wParam
            uInitialize

        Case TIMER_WAIT_FOR_QUIT
            g_Debug "quit wait completed..."
            KillTimer hWnd, wParam
            SetTimer hWnd, TIMER_WAIT_FOR_OUTLOOK, 2000, 0

        End Select


    Case snBroadcastMsg()
        Select Case wParam
        Case SNARL_BROADCAST_LAUNCHED
            g_Debug "[SNARL_BROADCAST_LAUNCHED]"
            Form1.Add "Snarl launched"

        Case SNARL_BROADCAST_QUIT
            g_Debug "[SNARL_BROADCAST_QUIT]"
            Form1.Add "Snarl quit"
            uTidyUp

        End Select

    Case snAppMsg()
        Select Case wParam
        Case SNARLAPP_DO_PREFS
            g_Debug "[SNARLAPP_DO_PREFS]"
            Form1.Add "_APP_PREFS"
'            uDoPrefs

        Case SNARLAPP_DO_ABOUT
            g_Debug "[SNARLAPP_DO_ABOUT]"
            snDoRequest "app-sig=" & App.ProductName & _
                        "&title=SnarlMail " & CStr(App.Major) & "." & CStr(App.Minor) & " " & App.Comments & _
                        "&text=" & App.LegalCopyright & vbCrLf & "Build " & CStr(App.Revision) & _
                        "&icon=" & g_MakePath(App.Path) & "icon.png"

        End Select


'    Case WM_NOTIFICATION
'        Form1.Add "NOTIFICATION: " & g_HexStr(wParam) & " " & CStr(lParam)
'        n = LoWord(wParam)
'
'        Select Case n
'        Case SNARL_NOTIFICATION_ACK, SNARL_NOTIFICATION_MENU
'            If Not uFind(lParam, pObj) Then _
'                Exit Function
'
'            Form1.Add "item found: " & pObj.Class
'
'            If (n = SNARL_NOTIFICATION_ACK) And (as_boolean(mConfig.GetValue("open_item_on_ack"))) Then
'                pObj.Display
'
'            ElseIf n = SNARL_NOTIFICATION_MENU Then
'
'                sn41Hide lParam
'                pObj.UnRead = False
'
''auto_mark_as_read
'
'                Select Case HiWord(wParam)
'                Case 1
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    pObj.Display
'
'                Case 2
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    pObj.Reply.Display
'
'                Case 3
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    pObj.ReplyAll.Display
'
'                Case 4
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    pObj.Forward.Display
'
'                Case 5
'                    pObj.Delete
'
'                Case 6
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    uAsAppointment(pObj).Respond(olMeetingAccepted).Display
'
'                Case 7
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    uAsAppointment(pObj).Respond(olMeetingTentative).Display
'
'                Case 8
'                    If as_boolean(mConfig.GetValue("auto_mark_as_read")) Then _
'                        pObj.UnRead = False
'
'                    uAsAppointment(pObj).Respond(olMeetingDeclined).Display
'
'                Case 9
''                    uAsAppointment(pObj).().Display
'
'
'
'                End Select
'
'            End If
'
'
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
'
'        End Select


    Case WM_TEST
        Form1.Add "TEST: " & wParam

'        If wParam = 0 Then
'            uFindFirst olMail, pObj
'            uNotifyMail pObj
'
'        ElseIf wParam = 1 Then
'            uFindFirst olMeetingRequest, pObj
'            uNotifyMeeting pObj, pObj.Class
'
'        End If

'    Case Else
'        form1.add "WM_" & g_HexStr(uMsg)

    End Select

End Function

'Private Sub uCheckOutlook()
'
'
'
'
'
''Dim pOutlook As Outlook.Application
''
''    On Error Resume Next
''
''    If uGetOutlook(pOutlook) Then
''        Form1.Add "Got Outlook " & pOutlook.Version
''        Set myOutlook = pOutlook
''        mGotOutlook = True
''        KillTimer mhWnd, TIMER_WAIT_FOR_OUTLOOK
''
''        ' /* get inbox content [not needed now as we use NewMailEx()?] */
''        err.Clear
''        Set myItems = myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
''        If (myItems Is Nothing) Or (err.Number <> 0) Then
''            Form1.Add "Couldn't retrieve inbox items object (" & err.Description & ")"
''
''        Else
''            myItems.Sort "Received", True
''
''        End If
''
''        ' /* get contacts list */
''
''        Set mContacts = myOutlook.GetNamespace("MAPI").GetDefaultFolder(olFolderContacts).Items
''
'''        If sn41IsSnarlRunning() Then
'''            uRegisterWithSnarl
'''
'''        Else
'''            Form1.Add "Snarl not running, waiting..."
'''
'''        End If
''
''    End If
'
'End Sub

'Private Sub myItems_ItemAdd(ByVal Item As Object)
'
''    On Error Resume Next
''
''Dim pMeeting As MeetingItem
''Dim pMail As MailItem
''Dim pItem As T_ITEM
''Dim hr As Long
''
''    form1.add "New message: class=" & Item.Class
''
''    Err.Clear
''    Set pMail = Item
''    If Err.Number = 0 Then
''        hr = uNotifyMail(pMail)
''
''    Else
''        Err.Clear
''        Set pMeeting = Item
''        If Err.Number = 0 Then
''            form1.add "meeting item"
'''            pItem.IsMailItem = False
'''            hr = g_NotifyMeeting(pMeeting)
''
''        Else
''            form1.add "unknown"
''
''        End If
''
''    End If
''
''    If hr > 0 Then
''        ' /* notification was displayed ok */
''        Set pItem.Item = Item
''        pItem.NotifyToken = hr
''
''        mItems = mItems + 1
''        ReDim Preserve mItem(mItems)
''        LSet mItem(mItems) = pItem
''
''        Debug.Print "added: " & mItems
''
''    End If
'
'End Sub
'
'Private Sub myOutlook_NewMail()
'
''    form1.add "NEW EMAIL"
'
'End Sub

Private Sub myOutlook_NewMailEx(ByVal EntryIDCollection As String)
Dim ItemId() As String
Dim oItem As Object
Dim i As Long

    On Error Resume Next

    g_Debug "[myOutlook::NewMailEx]"

    With myOutlook.Session
        Form1.Add "NewMailEx: " & EntryIDCollection
        ItemId = Split(EntryIDCollection, ",")
        For i = 0 To UBound(ItemId)
            err.Clear
            Set oItem = .GetItemFromID(ItemId(i))
            Form1.Add "Item #" & CStr(i) & " entry id=" & ItemId(i) & " " & err.Description
'            Select Case oItem.Class
'            Case olMail
'                uNotifyMail oItem
'
'            Case olMeetingRequest, olMeetingAccepted, olMeetingDeclined, olMeetingCancellation
'                uNotifyMeeting oItem, oItem.Class
'
'            End Select

            uNotify oItem

        Next i

    End With

End Sub

Private Sub myOutlook_Quit()

    g_Debug "[myOutlook::Quit] fired"
    uTidyUp

End Sub

'Public Function uNotifyMail(ByRef Mail As MailItem) As Boolean
'Dim nPriority As OlImportance
'Dim nSensitivity As OlSensitivity
'Dim szSubject As String
'Dim szSender As String
'Dim szBody As String
'Dim szClass As String
'Dim cxIcon As Long
'Dim c As Long
'
'    ' /* pre-set values: if 'Item' is Null, we show a preview notification */
'
'    If (Mail Is Nothing) Then
'        szSender = "Samuel Taylor Coleridge"
'        szSubject = "RE: Kubla Khan"
'        nPriority = olImportanceHigh
'        nSensitivity = olNormal
'        c = 1
'
'        If as_boolean(mConfig.GetValue("show_body")) Then _
'            szBody = "In Xanadu did Kubla Khan  A stately pleasure-dome decree:  Where Alph, the sacred river, ran  Through caverns measureless to man Down to a sunless sea.  So twice five miles of fertile ground  With walls and towers were girdled round:  And here were gardens bright with sinuous rills,  Where blossomed many an incense-bearing tree;  And here were forests ancient as the hills,  Enfolding sunny spots of greenery."
'
'    Else
'        szSender = Mail.SenderName
'        szSubject = Mail.Subject
'        nPriority = Mail.Importance
'        nSensitivity = Mail.Sensitivity
'        c = Mail.Attachments.Count
'
'        If as_boolean(mConfig.GetValue("show_body")) Then _
'            szBody = Mail.Body
'
'    End If
'
'    If (nSensitivity > olNormal) And (as_boolean(mConfig.GetValue("obscure_sensitive"))) And (szBody <> "") Then
'        szBody = vbCrLf & vbCrLf & "<Sensitive content>"
'
'    Else
'
''Dim i As Integer
''
''        For i = 1 To Len(szBody)
''            Debug.Print Mid$(szBody, i, 1) & "-" & Asc(Mid$(szBody, i, 1)) & " ";
''
''        Next i
'
'        If szBody <> "" Then
'            ' /* only want first 128 chars... */
'            szBody = g_SafeLeftStr(szBody, 128, True)
'
'            ' /* do some tidying up */
'            szBody = Replace$(szBody, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, "� ")
'            szBody = Replace$(szBody, vbCrLf & vbCrLf, "� ")
'            szBody = vbCrLf & vbCrLf & szBody
'
'        End If
'    End If
'
'    ' /* determine notification class *?
'
'    Select Case nPriority
'    Case olImportanceHigh
'        szClass = CLASS_URGENT_EMAIL
'
'    Case olImportanceLow
'        szClass = CLASS_BORING_EMAIL
'
'    Case Else
'        szClass = CLASS_EMAIL
'
'    End Select
'
'Dim sz As String
'
'    sz = "mail.png"
'
'    If as_boolean(mConfig.GetValue("show_body")) Then
'        Select Case g_SafeLeftStr(szSubject, 4)
'        Case "RE: "
'            sz = "mail-reply.png"
'
'        Case "FW: "
'            sz = "mail-forward.png"
'
'        End Select
'
'    End If
'
'    ' /* get the sender's contact picture */
'
'    If (Not (Mail Is Nothing)) And (as_boolean(mConfig.GetValue("use_contact_icon"))) Then
'        ' /* if uGetContactIcon() succeeds, it will drop a copy of the contact's picture
'        '    into the current working directory called '.temp_icon.png' */
'        If uGetContactIcon(uGetSMTPAddress(Mail.SenderEmailAddress)) Then _
'            sz = ".temp_icon.png"
'
'    End If
'
'    ' /* create icon */
'
'    uBuildIcon g_MakePath(App.Path) & sz, nPriority, nSensitivity, , c
'
'    ' /* add content preview? */
'
'Dim hr As Long
'
''    hr = sn41EZNotify(mToken, szClass, szSender, szSubject & szBody, , _
''                      g_MakePath(App.Path) & ".temp_icon.png")
''
''    If hr > 0 Then
''        ' /* notification was displayed ok */
''        mItems = mItems + 1
''        ReDim Preserve mItem(mItems)
''        With mItem(mItems)
''            .NotifyToken = hr
''            If Not (Mail Is Nothing) Then
''                .ItemId = Mail.EntryID
''                Set .Item = Mail
''
''            Else
''                .ItemId = "0000000"
''
''            End If
''
''        End With
''
''        ' /* set the right-click menu */
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
''        Form1.Add "Notification token: " & CStr(hr)
''
''    End If
''
'''    g_Notify = snShowMessageEx(szClass, _
''                               IIf(gConfig.layout = 2, szSubject, szSender), _
''                               IIf(gConfig.layout = 2, "From " & szSender, szSubject) & sz & szBody, _
''                               0, _
''                               g_MakePath(App.Path) & ".temp_icon.png", _
''                               ghWnd, _
''                               WM_NOTIFICATION_REPLY)
''
'
'
''        g_Debug "From: " & Mail.SenderName
''        g_Debug "Subject: " & Mail.Subject
''        g_Debug "Priority: " & Mail.Importance
''        g_Debug "Sensitivity: " & Mail.Sensitivity
''        g_Debug "Content: " & pMail.Body
'
'End Function
'
'Public Function uNotifyMeeting(ByRef Meeting As MeetingItem, ByVal Class As OlObjectClass) As Boolean
'Dim nPriority As OlImportance
'Dim nSensitivity As OlSensitivity
'Dim pAppt As AppointmentItem
'Dim szSubject As String
'Dim szSender As String
'Dim szBody As String
'Dim szClass As String
'Dim szContent As String
'Dim cxIcon As Long
'Dim sz As String
'Dim c As Long
'Dim i As Long
'
'    ' /* pre-set values: if 'Item' is Null, we show a preview notification */
'
'    If (Meeting Is Nothing) Then
'        szSender = "OPS/KOR/INTEL"
'        szSubject = "Operation Dinner Out"
'        nPriority = olImportanceHigh
'        nSensitivity = olNormal
'        szBody = vbCrLf & "Today 0900-1130" & vbCrLf & "Briefing Room A"
'        c = 1
'
'    Else
'        szSender = Meeting.SenderName
'        szSubject = Meeting.Subject
'        nPriority = Meeting.Importance
'        nSensitivity = Meeting.Sensitivity
'        c = Meeting.Attachments.Count
'        Set pAppt = Meeting.GetAssociatedAppointment(False)
'
''        If Meeting.Body <> "" Then
''            ' /* only want first 128 chars... */
''            szContent = g_SafeLeftStr(Meeting.Body, 128, True)
''
''            ' /* do some tidying up */
''            szContent = Replace$(szContent, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, "� ")
''            szContent = Replace$(szContent, vbCrLf & vbCrLf, "� ")
''            szContent = vbCrLf & vbCrLf & szContent
''
''        End If
'
'        ' /* figure out time */
'
'        i = DateDiff("d", Now, pAppt.Start)         ' // negative value = in the past
'
'        szBody = vbCrLf & IIf(i = 0, "Today, ", Format$(pAppt.Start, "short date")) & " " & _
'                 IIf(pAppt.AllDayEvent, " (all day)", Format$(pAppt.Start, "h:mmampm") & "-" & Format$(pAppt.End, "h:mmampm"))
'
'        ' /* if there's a location, add that now */
'
'        sz = pAppt.Location
'        If sz <> "" Then _
'            szBody = szBody & vbCrLf & sz
'
'        szBody = szBody & szContent
'
'    End If
'
''    If nSensitivity > olNormal Then
''        szBody = vbCrLf & vbCrLf & "<Sensitive content>"
''
''    Else
'
'    ' /* determine notification class */
'
'    Select Case nPriority
'    Case olImportanceHigh
'        szClass = CLASS_URGENT_MEETING
'
'    Case olImportanceLow
'        szClass = CLASS_BORING_MEETING
'
'    Case Else
'        szClass = CLASS_MEETING
'
'    End Select
'
'    ' /* default icon */
'
'    sz = ""
'
'    ' /* get the sender's contact picture */
'
'    If (Not (Meeting Is Nothing)) Then
'        ' /* if uGetContactIcon() succeeds, it will drop a copy of the contact's picture
'        '    into the current working directory called '.temp_icon.png' */
'        If uGetContactIcon(uGetSMTPAddress(Meeting.SenderEmailAddress)) Then _
'            sz = ".temp_icon.png"
'
'    End If
'
'Dim szBadge As String
'
'    Select Case Class
'    Case olMeetingResponsePositive
'        szBadge = "emblem-accept"
'
'    Case olMeetingResponseNegative
'        szBadge = "emblem-decline"
'
'    Case olMeetingResponseTentative
'        szBadge = "emblem-tentative"
'
'    Case Else
'        szBadge = "emblem-meeting"
'
'    End Select
'
'    ' /* create icon */
'
'    uBuildIcon g_MakePath(App.Path) & sz, nPriority, nSensitivity, szBadge, c
'
'    ' /* add content preview? */
'
'Dim hr As Long
'
''    hr = sn41EZNotify(mToken, szClass, szSubject, szSender & szBody, , g_MakePath(App.Path) & ".temp_icon.png")
''    If hr > 0 Then
''        ' /* notification was displayed ok */
''        mItems = mItems + 1
''        ReDim Preserve mItem(mItems)
''        With mItem(mItems)
''            .NotifyToken = hr
''            .IsMeeting = True
''
''            If Not (Meeting Is Nothing) Then
''                .ItemId = Meeting.EntryID
''                Set .Item = Meeting
''
''            Else
''                .ItemId = "0000000"
''
''            End If
''
''        End With
''
''        ' /* set the right-click menu */
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Accept#?6|Tentative#?7|Decline#?8||Propose New Time#?9||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
''        Form1.Add "Notification token: " & CStr(hr)
''
''    End If
''
'''    g_Notify = snShowMessageEx(szClass, _
''                               IIf(gConfig.layout = 2, szSubject, szSender), _
''                               IIf(gConfig.layout = 2, "From " & szSender, szSubject) & sz & szBody, _
''                               0, _
''                               g_MakePath(App.Path) & ".temp_icon.png", _
''                               ghWnd, _
''                               WM_NOTIFICATION_REPLY)
''
''
''
'''        g_Debug "From: " & Meeting.SenderName
'''        g_Debug "Subject: " & Meeting.Subject
'''        g_Debug "Priority: " & Meeting.Importance
'''        g_Debug "Sensitivity: " & Meeting.Sensitivity
'''        g_Debug "Content: " & pappointment.Body
'
'End Function



'Private Function uGetContactIcon(ByVal EmailAddress As String) As Boolean
'
''    Exit Function
'
'    DeleteFile g_MakePath(App.Path) & ".temp_icon.png"
'
'Dim pOutlook As Outlook.Application
'
'    If Not uGetOutlook(pOutlook) Then _
'        Exit Function
'
'Dim pContact As ContactItem
'
'    If Not uFindContactByEmail(EmailAddress, pContact) Then
'        g_Debug "uGetContactIcon(): '" & EmailAddress & "' not in contacts"
'        Exit Function
'
'    End If
'
'    If Not pContact.HasPicture Then
'        g_Debug "uGetContactIcon(): '" & EmailAddress & "' has no picture associated"
'        Exit Function
'
'    End If
'
'    If pContact.Attachments.Count = 0 Then _
'        Exit Function
'
'Dim i As Long
'
'    With pContact.Attachments
'        For i = 1 To .Count
'            If LCase$(.Item(i).DisplayName) = "contactpicture.jpg" Then
'                .Item(i).SaveAsFile g_MakePath(App.Path) & ".temp_icon.png"
'                uGetContactIcon = True
'                Exit Function
'
'            End If
'        Next i
'
'    End With
'
'End Function
'
'Private Function uFindContactByEmail(ByVal Email As String, ByRef Contact As ContactItem) As Boolean
'
'    If (mContacts Is Nothing) Then _
'        Exit Function
'
'    If mContacts.Count = 0 Then _
'        Exit Function
'
'    Email = LCase$(Email)
'
'Dim pContact As ContactItem
'Dim i As Long
'
'    With mContacts
'        For i = 1 To .Count
'            Set pContact = .Item(i)
'            If (LCase$(pContact.Email1Address) = Email) Or (LCase$(pContact.Email2Address) = Email) Or (LCase$(pContact.Email3Address) = Email) Then
'                g_Debug "uFindContactByEmail(): found contact '" & pContact.FullName & "'"
'                Set Contact = pContact
'                uFindContactByEmail = True
'                Exit Function
'
'            End If
'        Next i
'    End With
'
'End Function
'
'Private Sub uBuildIcon(ByVal Filename As String, ByVal Priority As OlImportance, ByVal Sensitivity As OlSensitivity, Optional ByVal Badge As String, Optional ByVal AttachmentCount As Long)
'Dim pbm As mfxBitmap
'Dim cxIcon As Long
'Dim sz As String
'Dim pt As BPoint
'
'    Set pbm = load_image_obj(Filename)
'    If Not is_valid_image(pbm) Then _
'        Set pbm = uCreateFakeImage()
'
'    Set pbm = uMakeSquareImage(pbm)
'
'    cxIcon = Fix(pbm.Width / 3)
'
'    With New mfxView
'        .SizeTo pbm.Width, pbm.Height
'        .DrawScaledImage pbm
'
'        Set pt = new_BPoint(0, .Height - cxIcon - 1)
'
'        ' /* priority emblem */
'        If Priority = olImportanceHigh Then
'            .DrawScaledImage uCreateEmblem("!"), pt, new_BPoint(cxIcon, cxIcon)
'            pt.OffsetBy cxIcon, 0
'
'        End If
'
'        If AttachmentCount > 0 Then
'            .DrawScaledImage uCreateEmblem(CStr(AttachmentCount)), pt, new_BPoint(cxIcon, cxIcon)
'            pt.OffsetBy cxIcon, 0
'
'        End If
'
'        ' /* badge (only one) top-right */
'
''Dim szBadge As String
''
''        Select Case Acceptance
''        Case olMeetingAccepted
''            szBadge = "emblem-accept"
''
''        Case olMeetingDeclined
''            szBadge = "emblem-decline"
''
''        Case olMeetingTentative
''            szBadge = "emblem-tentative"
''
'''        Case Else
'''            If IsMeeting Then _
'''                szBadge = "emblem-meeting"
''
''        End Select
'
'        If Badge <> "" Then
'            cxIcon = cxIcon * 1.5
'            .DrawScaledImage load_image_obj(g_MakePath(App.Path) & Badge & ".png"), _
'                             new_BPoint(.Width - cxIcon, .Height - cxIcon), new_BPoint(cxIcon, cxIcon)
'
'        End If
'
'        .WriteToFile g_MakePath(App.Path) & ".temp_icon.png", "image/png"
'
'    End With
'
'End Sub
'
'Private Function uCreateFakeImage() As mfxBitmap
'
'    With New mfxView
'        .SizeTo 128, 128
'        .SetHighColour rgba(255, 255, 255, 102)
'        .StrokeRoundRect .Bounds, 20, 20, 2
'        Set uCreateFakeImage = .ConvertToBitmap()
'
'    End With
'
'End Function
'
'Private Function uGetSMTPAddress(ByVal Email As String) As String
'
'    On Error Resume Next
'
'    Debug.Print "uGetSMTPAddress(): " & Email
'
'    If g_SafeLeftStr(Email, 3) <> "/o=" Then
'        uGetSMTPAddress = Email
'        Exit Function
'
'    End If
'
'    uGetSMTPAddress = Email
'
''NameSpace.GetRecipientFromID
'
'
''Dim oSession As MAPI.Session
''
''    Err.Clear
''    Set oSession = CreateObject("MAPI.Session")
''    If (Err.Number <> 0) Or (oSession Is Nothing) Then
''        g_Debug "uGetSMTPAddress(): couldn't create MAPI.Session", LEMON_LEVEL_CRITICAL
''        Exit Function
''
''    End If
''
'''strEMailAddress = objAddressEntry.Address
'''
'''' Check if it is an Exchange object
'''If Left(strEMailAddress, 3) = "/o=" Then
'''
'''  ' Get the SMTP address
'''  strAddressEntryID = objAddressEntry.Id
'''  strEMailAddress =_
'''    objSession.GetAddressEntry(strAddressEntryID).Fields(CdoPR_EMAIL).Value
'''End If
'''
'''' Display the SMTP address of current user
'''MsgBox "SMTP address of current user: " & strEMailAddress
'
'End Function

'Private Function uFind(ByVal Token As Long, ByRef Item As Object) As Boolean
'
'    On Error Resume Next
'
'    If mItems = 0 Then _
'        Exit Function
'
'Dim i As Long
'
'    For i = 1 To mItems
'        If mItem(i).NotifyToken = Token Then
'            Form1.Add "Item " & Token & " found at #" & CStr(i) & " itemID=" & mItem(i).ItemId
'            Set Item = mItem(i).Item
'            uFind = (Not (Item Is Nothing))
'            Form1.Add "IsValid()=" & uFind
'            Exit Function
'
'        End If
'    Next i
'
'    Form1.Add "Item " & Token & " not found"
'
'End Function

'Private Sub uDoPrefs()
'Dim pPage As BPrefsPage
'Dim pC As BControl
'
'    If (myPanel Is Nothing) Then
'
'        Set myPanel = New BPrefsPanel
'        With myPanel
'            .SetTitle "SnarlMail Preferences"
'            .SetWidth 380
'
'            Set pPage = new_BPrefsPage()
'            With pPage
'                .SetMargin 32
'                .Add new_BPrefsControl("banner", "", "Appearance")
'                .Add new_BPrefsControl("fancytoggle2", "show_body", "Include message body in notification?", , mConfig.GetValue("show_body"))
'                .Add new_BPrefsControl("fancytoggle2", "obscure_sensitive", "Obscure body content if item is marked as sensitive?", , mConfig.GetValue("obscure_sensitive"))
'                .Add new_BPrefsControl("fancytoggle2", "extra_icons", "Use forward/reply icons?", , mConfig.GetValue("extra_icons"))
'                .Add new_BPrefsControl("fancytoggle2", "use_contact_icon", "Use contact picture (if available)?", , mConfig.GetValue("use_contact_icon"))
'
'                .Add new_BPrefsControl("banner", "", "Behaviour")
'                .Add new_BPrefsControl("fancytoggle2", "open_item_on_ack", "Open item when notification clicked?", , mConfig.GetValue("open_item_on_ack"))
'                .Add new_BPrefsControl("fancytoggle2", "auto_mark_as_read", "Mark item as read when replying or forwarding?", , mConfig.GetValue("auto_mark_as_read"))
'
'                .Add new_BPrefsSeparator()
''                .Add new_BPrefsControl("label", "", "something interesting here:")
''
''                Set pC = new_BPrefsControl("fancyedit", "", "2048", "Limit:")
''                pC.Indent 4
''                pC.SizeTo 48, -1
''                .Add pC
''
''                Set pC = new_BPrefsControl("fancyedit", "", "2048", "Limit:")
''                .Add pC
'
'                .Add new_BPrefsControl("fancybutton2", "about", "About SnarlMail")
'                .Add new_BPrefsSeparator()
'                .Add new_BPrefsControl("fancybutton2", "quit", "Quit SnarlMail", , , new_BPackedData("button_of_doom::1"))
'
'            End With
'
'            .AddPage pPage
'
'            .Go
'
'            .Show True
'
'        ' /* options:
'        '       mark item as read when replying or forwarding?
'        '       group multiple new items into a single notification?
'        '       open item if notification is clicked?
'        '
'        ' */
'
'        End With
'
'    Else
'        myPanel.Show True
'
'    End If
'
'End Sub

'Private Function uCreateEmblem(ByVal Text As String) As mfxBitmap
'
'    If Len(Text) > 1 Then _
'        Text = "+"
'
'    With New mfxView
'        .SizeTo 64, 64
'        .EnableSmoothing True
'        .SetHighColour rgba(255, 0, 0)
'        .FillEllipse .Bounds
'        .SetHighColour rgba(0, 0, 0)
'        .StrokeEllipse .Bounds
'        .SetHighColour rgba(255, 255, 255)
'        .SetFont "Arial", 36, True
'        .DrawString Text, .Bounds, MFX_ALIGN_H_CENTER Or MFX_ALIGN_V_CENTER
'        Set uCreateEmblem = .ConvertToBitmap()
'
'    End With
'
'End Function

'Private Function uFindFirst(ByVal ItemClass As OlObjectClass, ByRef Item As Object, Optional ByVal MaxSearchCount As Long = 32) As Boolean
'
'    If (myItems Is Nothing) Then _
'        Exit Function
'
'Dim i As Long
'
'    With myItems
'        For i = 1 To MIN(MaxSearchCount, .Count)
'            Debug.Print .Item(i).Subject & " (" & .Item(i).Class & ")"
'            If .Item(i).Class = ItemClass Then
'                Set Item = .Item(i)
'                uFindFirst = True
'                Exit Function
'
'            End If
'        Next i
'    End With
'
'End Function
'
'Private Function uAsAppointment(ByRef Item As MeetingItem) As AppointmentItem
'
'    Set uAsAppointment = Item.GetAssociatedAppointment(False)
'
'End Function

'Private Sub uGetConfig()
'Dim pDefaults As BPackedData
'Dim szPath As String
'
'    Set pDefaults = New BPackedData
'    With pDefaults
'        .Add "open_item_on_ack", "1"
'        .Add "auto_mark_as_read", "1"
'        .Add "obscure_sensitive", "1"
'        .Add "show_body", "1"
'        .Add "extra_icons", "1"
'        .Add "use_contact_icon", "1"
'
'    End With
'
'    get_user_app_data szPath
'
'    Set mConfig = New BPersistentSettings
'    mConfig.SetTo g_MakePath(szPath) & ".snarlmail", pDefaults
'
'End Sub

'Private Sub myPanel_ControlChanged(Control As prefs_kit12.BControl, ByVal Value As String)
'
'    mConfig.SetValue Control.Name, Value
'
'End Sub
'
'Private Sub myPanel_ControlInvoked(Control As prefs_kit12.BControl)
'
'    Select Case Control.Name
'    Case "about"
''        PostMessage mhWnd, sn41AppMsg(), SNARL_APP_SHOW_ABOUT, ByVal 0&
'
'    Case "quit"
'        PostQuitMessage 0
'
'    End Select
'
'End Sub
'
'Private Sub myPanel_Quit()
'
'    Set myPanel = Nothing
'
'End Sub
'
'Private Function uMakeSquareImage(ByRef Img As MImage, Optional ByVal Maximum As Long) As mfxBitmap
'Dim pv As mfxView
'Dim c As Long
'
'    If Not is_valid_image(Img) Then _
'        Exit Function
'
'    c = MAX(Img.Width, Img.Height)
'    If Maximum > 0 Then _
'        c = MIN(c, Maximum)
'
'    Set pv = New mfxView
'    With pv
'        .SizeTo c, c
'        .DrawScaledImage Img, new_BPoint(Fix((.Width - Img.Width) / 2), Fix((.Height - Img.Height) / 2))
'        Set uMakeSquareImage = .ConvertToBitmap()
'
'    End With
'
'End Function

Private Sub uInitialize()

    On Error Resume Next

    g_Debug "TWindow.uInitialize()", LEMON_LEVEL_PROC_ENTER

    err.Clear
    Set myOutlook = GetObject(, "outlook.application")
    If err.Number <> 0 Then
        g_Debug "GetObject('outlook.application') failed: " & err.Description, LEMON_LEVEL_CRITICAL Or LEMON_LEVEL_PROC_EXIT
        PostQuitMessage 0
        Exit Sub

    End If

    ' /* register */

    If snarl_register(App.ProductName, App.Title, g_MakePath(App.Path) & "icon.png", , _
                      mhWnd, WM_NOTIFICATION, _
                      SNARLAPP_IS_WINDOWLESS Or SNARLAPP_HAS_ABOUT) > 0 Then

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_EMAIL_LOW & _
                    "&name=Low priority items"

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_EMAIL & _
                    "&name=Normal priority items"

        snDoRequest "addclass?app-sig=" & App.ProductName & _
                    "&id=" & CLASS_EMAIL_HIGH & _
                    "&name=High priority items"

    Else
        g_Debug "snarl_register() failed", LEMON_LEVEL_CRITICAL
        PostQuitMessage 0

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Private Sub uTidyUp()

    g_Debug "TWindow.uTidyUp()", LEMON_LEVEL_PROC_ENTER

    snarl_unregister App.ProductName
    Set myOutlook = Nothing
    SetTimer mhWnd, TIMER_WAIT_FOR_QUIT, 5000, 0

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Sub

Public Function uNotify(ByRef Mail As MailItem) As Boolean
Dim szClass As String

    Select Case Mail.Importance
    Case olImportanceHigh
        szClass = CLASS_EMAIL_HIGH

    Case olImportanceLow
        szClass = CLASS_EMAIL_LOW

    Case Else
        szClass = CLASS_EMAIL

    End Select

    snDoRequest "notify?app-sig=" & App.ProductName & _
                "&id=" & szClass & _
                "&title=" & Mail.SenderName & "&text=" & Mail.Subject & _
                "&icon=!message-new_mail"

'Dim nSensitivity As OlSensitivity
'Dim szSubject As String
'Dim szSender As String
'Dim szBody As String
'Dim szClass As String
'Dim cxIcon As Long
'Dim c As Long
'
'    ' /* pre-set values: if 'Item' is Null, we show a preview notification */
'
'    If (Mail Is Nothing) Then
'        szSender = "Samuel Taylor Coleridge"
'        szSubject = "RE: Kubla Khan"
'        nPriority = olImportanceHigh
'        nSensitivity = olNormal
'        c = 1
'
'        If as_boolean(mConfig.GetValue("show_body")) Then _
'            szBody = "In Xanadu did Kubla Khan  A stately pleasure-dome decree:  Where Alph, the sacred river, ran  Through caverns measureless to man Down to a sunless sea.  So twice five miles of fertile ground  With walls and towers were girdled round:  And here were gardens bright with sinuous rills,  Where blossomed many an incense-bearing tree;  And here were forests ancient as the hills,  Enfolding sunny spots of greenery."
'
'    Else
'        szSender = Mail.SenderName
'        szSubject = Mail.Subject
'        nPriority = Mail.Importance
'        nSensitivity = Mail.Sensitivity
'        c = Mail.Attachments.Count
'
'        If as_boolean(mConfig.GetValue("show_body")) Then _
'            szBody = Mail.Body
'
'    End If
'
'    If (nSensitivity > olNormal) And (as_boolean(mConfig.GetValue("obscure_sensitive"))) And (szBody <> "") Then
'        szBody = vbCrLf & vbCrLf & "<Sensitive content>"
'
'    Else
'
''Dim i As Integer
''
''        For i = 1 To Len(szBody)
''            Debug.Print Mid$(szBody, i, 1) & "-" & Asc(Mid$(szBody, i, 1)) & " ";
''
''        Next i
'
'        If szBody <> "" Then
'            ' /* only want first 128 chars... */
'            szBody = g_SafeLeftStr(szBody, 128, True)
'
'            ' /* do some tidying up */
'            szBody = Replace$(szBody, vbCrLf & vbCrLf & Chr$(&HA0) & vbCrLf & vbCrLf, "� ")
'            szBody = Replace$(szBody, vbCrLf & vbCrLf, "� ")
'            szBody = vbCrLf & vbCrLf & szBody
'
'        End If
'    End If
'
'    ' /* determine notification class *?
'
'    Select Case nPriority
'    Case olImportanceHigh
'        szClass = CLASS_URGENT_EMAIL
'
'    Case olImportanceLow
'        szClass = CLASS_BORING_EMAIL
'
'    Case Else
'        szClass = CLASS_EMAIL
'
'    End Select
'
'Dim sz As String
'
'    sz = "mail.png"
'
'    If as_boolean(mConfig.GetValue("show_body")) Then
'        Select Case g_SafeLeftStr(szSubject, 4)
'        Case "RE: "
'            sz = "mail-reply.png"
'
'        Case "FW: "
'            sz = "mail-forward.png"
'
'        End Select
'
'    End If
'
'    ' /* get the sender's contact picture */
'
'    If (Not (Mail Is Nothing)) And (as_boolean(mConfig.GetValue("use_contact_icon"))) Then
'        ' /* if uGetContactIcon() succeeds, it will drop a copy of the contact's picture
'        '    into the current working directory called '.temp_icon.png' */
'        If uGetContactIcon(uGetSMTPAddress(Mail.SenderEmailAddress)) Then _
'            sz = ".temp_icon.png"
'
'    End If
'
'    ' /* create icon */
'
'    uBuildIcon g_MakePath(App.Path) & sz, nPriority, nSensitivity, , c
'
'    ' /* add content preview? */
'
'Dim hr As Long
'
''    hr = sn41EZNotify(mToken, szClass, szSender, szSubject & szBody, , _
''                      g_MakePath(App.Path) & ".temp_icon.png")
''
''    If hr > 0 Then
''        ' /* notification was displayed ok */
''        mItems = mItems + 1
''        ReDim Preserve mItem(mItems)
''        With mItem(mItems)
''            .NotifyToken = hr
''            If Not (Mail Is Nothing) Then
''                .ItemId = Mail.EntryID
''                Set .Item = Mail
''
''            Else
''                .ItemId = "0000000"
''
''            End If
''
''        End With
''
''        ' /* set the right-click menu */
''        snChangeAttribute hr, SNARL_ATTRIBUTE_MENU, "Open#?1||Reply#?2|Reply All#?3|Forward#?4||Delete#?5"
''        Form1.Add "Notification token: " & CStr(hr)
''
''    End If
''
'''    g_Notify = snShowMessageEx(szClass, _
''                               IIf(gConfig.layout = 2, szSubject, szSender), _
''                               IIf(gConfig.layout = 2, "From " & szSender, szSubject) & sz & szBody, _
''                               0, _
''                               g_MakePath(App.Path) & ".temp_icon.png", _
''                               ghWnd, _
''                               WM_NOTIFICATION_REPLY)
''
'
'
''        g_Debug "From: " & Mail.SenderName
''        g_Debug "Subject: " & Mail.Subject
''        g_Debug "Priority: " & Mail.Importance
''        g_Debug "Sensitivity: " & Mail.Sensitivity
''        g_Debug "Content: " & pMail.Body

End Function


