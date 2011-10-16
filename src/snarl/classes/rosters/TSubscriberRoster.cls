VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSubscriberRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TSubscriberRoster.cls
    '/
    '/  Description:    Manages subscribed GNTP and SNP3 clients
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mSubscriberConfig As ConfigFile
Dim mSubscribers As BTagList
Dim mRefs As Long

Dim mListenerConfig As ConfigFile
Dim mListeners As BTagList

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */

    End If

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT
Dim pc As ConfigSection
Dim ps As TSubscriber
Dim pl As TWebListener

    g_Debug "TSubscriberRoster.Open()", LEMON_LEVEL_PROC_ENTER

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        Set mSubscribers = new_BTagList()
'        Set mGNTPSubscribers = new_BTagList()

        Set mSubscriberConfig = New ConfigFile
        With mSubscriberConfig
            .File = g_SettingsPath() & "network.snarl"
            .Load

            ' /* load up persistent entries - specifically:
            '       forwards from this computer to a remote computer
            '       subscriptions to a remote computer
            ' */

            .Rewind

            Do While .GetNextSection(pc)
                Set ps = New TSubscriber
                If ps.SetFrom(pc) Then
                    mSubscribers.Add ps
                    g_Debug "got " & ps.TypeStr & " to " & ps.Destination

                End If
            Loop

        End With

        Set mListeners = new_BTagList()
        Set mListenerConfig = New ConfigFile
        With mListenerConfig
            .File = g_SettingsPath() & "listeners.snarl"
            .Load

            ' /* load up persistent entries - specifically:
            ' */

            .Rewind

            Do While .GetNextSection(pc)
                Set pl = New TWebListener
                If pl.SetFrom(pc) Then
                    mListeners.Add pl
                    g_Debug "got weblistener '" & pl.Name & "' url=" & pl.URL

                End If
            Loop

        End With


        g_Debug "roster initialised"

    End If

    g_Debug "", LEMON_LEVEL_PROC_EXIT

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mSubscribers.CountItems '+ mGNTPSubscribers.CountItems

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
Dim i As Long

'    If mCount Then
'        Name = LCase$(Name)
'        For i = 1 To mCount
'            If LCase$(mApp(i).Name) = Name Then
'                MRoster_FindItem = i
'                Exit Function
'
'            End If
'        Next i
'
'        g_Debug "TApplicationRoster.FindItem(): app '" & Name & "' not found", LEMON_LEVEL_WARNING
'
'    Else
'        g_Debug "TApplicationRoster.FindItem(): no apps registered!", LEMON_LEVEL_CRITICAL
'
'    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.mObject

'    If (Index > 0) And (Index <= mCount) Then _
        Set MRoster_ItemAt = mApp(Index)

End Function

Private Function MRoster_NextItem(Item As melon.mObject) As Boolean

'    If mIndex <= mCount Then
'        Set Item = mApp(mIndex)
'        mIndex = mIndex + 1
'        MRoster_NextItem = True
'
'    End If

End Function

Private Sub MRoster_Rewind()

'    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "24-Jun-2011"           ' // creation date, not last modification date

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "subscriber.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

'Public Function Add(ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Socket As CSocket, ByRef Data As BPackedData) As Boolean
Public Function Add(ByVal SubType As SN_SUBSCRIBER_TYPE, ByVal Description As String, ByVal RemoteHost As String, ByVal Apps As String) As Boolean
Dim pc As ConfigSection
Dim ps As TSubscriber

    g_Debug "TSubscriberRoster.Add(): type=0x" & g_HexStr(SubType, 4) & " dest=" & RemoteHost

    Select Case SubType
    Case SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION
        ' /*
        '   This represents a persistent forward or subscription from this computer to a remote destination using SNP3
        '
        '   Description: optional
        '   RemoteHost: required, specifies where to forward the notifications to
        '   Apps: optional, identifies which app(s) to forward notifications for
        '
        '   On successful add: add to config and save
        '
        ' */

        ' /* error checking */

        If (RemoteHost = "") Or (InStr(get_ip_address_table(), RemoteHost) <> 0) Then
            g_Debug "TSubscriberRoster.Add(): RemoteHost cannot be ourself", LEMON_LEVEL_CRITICAL
            Exit Function
    
        End If

        ' /* check this forward/subscription doesn't already exist */

        If uFindByTypeAndDestination(SubType, RemoteHost, Nothing) Then
            g_Debug "TSubscriberRoster.Add(): already exists", LEMON_LEVEL_CRITICAL
            Exit Function
    
        End If

        ' /* create persistent config */

        Set pc = mSubscriberConfig.AddSectionObj(g_CreateGUID(True))
        With pc
            .Add "type", CStr(SubType)
            .Add "description", Description
            .Add "remotehost", RemoteHost
            .Add "apps", Apps
            .Add "enabled", "1"

        End With

        ' /* create and initialise the subscriber */

        Set ps = New TSubscriber
        If Not ps.SetFrom(pc) Then
            g_Debug "TSubscriberRoster.Add(): failed to initialise", LEMON_LEVEL_CRITICAL
            mSubscriberConfig.RemoveSection mSubscriberConfig.FindSection(pc.Name)      ' // remove from the config
            Exit Function

        End If

        ' /* success! so add it to the list */

        mSubscribers.Add ps
        mSubscriberConfig.Save
        Add = True

    End Select

'    ps.SetTo "", SubType, Description, RemoteHost, Apps, mConfig.AddSectionObj("")

End Function

Private Function uAddNewSubscriber(ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Socket As CSocket, ByRef Data As BPackedData) As Boolean

'    If (Socket Is Nothing) Or (Data Is Nothing) Then _
'        Exit Function
'
'Dim ps As TSubscriber
'
'    If mSubscribers.Find(Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort), ps) Then
'        ' /* already subscribed */
'        g_Debug "TSubscriberRoster.uAddNewSubscriber(): '" & Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort) & "' is already subscribed", LEMON_LEVEL_INFO
'
'        ' /* check password */
'        If ps.Password <> Data.ValueOf("password") Then
'            g_Debug "TSubscriberRoster.uAddNewSubscriber(): password mismatch", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_AUTH_FAILURE
'
'        ElseIf Data.Exists("app-sig") Then
'            ' /* add specific apps */
'            uAddNewSubscriber = ps.AddApps(Data)
'
'        Else
'            ' /* cannot do this if already subscribed */
'            g_Debug "TSubscriberRoster.uAddNewSubscriber(): cannot do wildcard subscribe", LEMON_LEVEL_CRITICAL
'            gSetLastError SNARL_ERROR_FAILED
'
'        End If
'
'    Else
'        ' /* not already subscribed */
'        Set ps = New TSubscriber
'        ps.SetTo Socket, SubscriberType, Data
'        mSubscribers.Add ps
'        uAddNewSubscriber = True
'
'    End If

End Function

Public Function Remove(ByRef Socket As CSocket, ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Data As BPackedData) As Boolean

    If (Socket Is Nothing) Or (Data Is Nothing) Then _
        Exit Function

Dim ps As TSubscriber
Dim i As Long

    i = mSubscribers.IndexOf(Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort))
    If i Then
        Debug.Print "TSubscriberRoster.Remove(): subscriber '" & Socket.RemoteHostIP & ":" & CStr(Socket.RemotePort) & "' found"

        Set ps = mSubscribers.TagAt(i)

        If ps.SubscriptionType <> SubscriberType Then
            ' /* wrong type */
            g_Debug "TSubscriberRoster.Remove(): '" & Socket.RemoteHostIP & "' type mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_INVALID_ARG

        ElseIf ps.Password <> Data.ValueOf("password") Then
            ' /* wrong password */
            g_Debug "TSubscriberRoster.Remove(): password mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_AUTH_FAILURE

        Else
            ' /* remove */
            mSubscribers.Remove i
            Remove = True

        End If

    Else
        g_Debug "TSubscriberRoster.Remove(): '" & Socket.RemoteHostIP & "' is not subscribed", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_NOT_SUBSCRIBED

    End If

End Function

Friend Sub Broadcast(ByRef Content As T_NOTIFICATION_INFO)
Dim szRequest As String
Dim ps As TSubscriber
Dim pa As TApp

    szRequest = uCreateSNP3Forward(Content)
    Set pa = Content.ClassObj.App

    With mSubscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            If ps.IsEnabled Then
                ' /* TO-DO: if the subscriber uses GNTP we will need to do
                '    a two-pass process.  Step 1 will do the REGISTER;
                '    STEP 2 will do the NOTIFY */

                If ps.IsSubscribedTo(pa.Signature) Then _
                    ps.Send szRequest

            End If
        Loop

    End With

End Sub

Public Function List() As BTagList

    Set List = mSubscribers

End Function

Public Function RemoveByGuid(ByVal Guid As String) As Boolean
Dim f As Boolean
Dim i As Long

    i = mSubscribers.IndexOf(Guid)
    If i <> 0 Then
        mSubscribers.Remove i
        f = True

    Else
        g_Debug "TSubscriberRoster.RemoveByGuid(): '" & Guid & "' not in roster", LEMON_LEVEL_WARNING

    End If

    i = mSubscriberConfig.FindSection(Guid)
    If i <> 0 Then
        mSubscriberConfig.RemoveSection i
        mSubscriberConfig.Save

    Else
        g_Debug "TSubscriberRoster.RemoveByGuid(): '" & Guid & "' not in config", LEMON_LEVEL_WARNING
        f = False

    End If

    RemoveByGuid = f

End Function

Public Sub WriteConfig()

    If Not (mSubscriberConfig Is Nothing) Then _
        mSubscriberConfig.Save


End Sub

Private Function uFindByTypeAndDestination(ByVal aType As SN_SUBSCRIBER_TYPE, ByVal Destination As String, ByRef Subscriber As TSubscriber) As Boolean
Dim ps As TSubscriber

    With mSubscribers
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            If (ps.SubscriptionType = aType) And (ps.Destination = Destination) Then
                Set Subscriber = ps
                uFindByTypeAndDestination = True
                Exit Function

            End If

        Loop
    End With

End Function

Private Function uTranslateIcon(ByVal Icon As String) As String

    If Icon = "" Then _
        Exit Function

Dim sz As String
Dim b As String

'    If (LCase$(g_SafeLeftStr(Icon, 7)) = "http://") Or _
'       (LCase$(g_SafeLeftStr(Icon, 8)) = "https://") Or _

       
    If g_SafeLeftStr(Icon, 1) = "!" Then
'        ' /* add verbatim */
'        sz = sz & "&icon=" & Icon

    Else
        ' /* encode it in a slightly modified Base64 format (CRLF's are replaced with #'s) */
        If uEncodeIcon(Icon, b) Then _
            sz = sz & "&icon-phat64=" & b

    End If

    uTranslateIcon = sz

End Function

Private Function uEncodeIcon(ByVal IconPath As String, ByRef Base64 As String) As Boolean

    If IconPath = "" Then _
        Exit Function

    If Not g_Exists(IconPath) Then _
        Exit Function

Dim sz As String
Dim i As Integer

    On Error Resume Next

    i = FreeFile()

    err.Clear
    Open IconPath For Binary Access Read Lock Write As #i
    If err.Number = 0 Then
        sz = String$(LOF(i), Chr$(0))
        Get #i, , sz
        Close #i

        sz = Encode64orig(sz)                   ' // encode as standard Base64
        If sz <> "" Then
            Base64 = Replace$(sz, vbCrLf, "#")  ' // replace CRLFs
            Base64 = Replace$(Base64, "=", "%") ' // replace ='s
            uEncodeIcon = True

        End If

    Else
        g_Debug "TSubscriberRoster.uEncodeIcon(): " & err.Description, LEMON_LEVEL_CRITICAL

    End If

End Function

Public Function Listeners() As BTagList

    Set Listeners = mListeners

End Function

Public Function RemoveListenerByGuid(ByVal Guid As String) As Boolean
Dim i As Long

    i = mListeners.IndexOf(Guid)
    If i = 0 Then
        g_Debug "TSubscriberRoster.RemoveListenerByGuid(): '" & Guid & "' not in roster", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mListeners.Remove i

    i = mListenerConfig.FindSection(Guid)
    If i <> 0 Then
        mListenerConfig.RemoveSection i
        mListenerConfig.Save

    Else
        g_Debug "TSubscriberRoster.RemoveListenerByGuid(): '" & Guid & "' not in config", LEMON_LEVEL_WARNING

    End If

    RemoveListenerByGuid = True

End Function

Public Function AddWebListener(ByVal Description As String, ByVal URL As String) As Boolean
Dim pc As ConfigSection
Dim pl As TWebListener

    g_Debug "TSubscriberRoster.AddWebListener(): name=" & Description & " URL=" & URL

    ' /* create and initialise the subscriber */

    Set pl = New TWebListener
    If Not pl.CreateNew(Description, URL) Then
        g_Debug "TSubscriberRoster.AddWebListener(): failed to initialise", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mListeners.Add pl

    ' /* create persistent config */

    Set pc = mListenerConfig.AddSectionObj(pl.Guid)
    With pc
        .Add "type", CStr(SN_ST_WEB_LISTENER)               ' // for now, in case we support other models in the future
        .Add "name", Description
        .Add "url", URL
        .Add "enabled", "1"

    End With

    mListenerConfig.Save
    AddWebListener = True

End Function

Private Function uCreateSNP3Forward(ByRef Content As T_NOTIFICATION_INFO) As String
Dim szc As String
Dim sz As String

    sz = Content.OriginalContent
    sz = Replace$(sz, "::", "=")
    sz = Replace$(sz, "#?", "&")
    sz = Replace$(sz, vbCrLf, "\n")

Dim pa As TApp

    Set pa = Content.ClassObj.App

    ' /* base content */

    szc = "SNP/3.0" & vbCrLf & _
          "register?app-sig=" & pa.Signature & "&title=" & pa.Name & uTranslateIcon(pa.Icon) & vbCrLf

    ' /* add classes */

Dim i As Long

    With pa
        If .CountAlerts Then
            For i = 1 To .CountAlerts
                With .AlertAt(i)
                    szc = szc & "addclass?app-sig=" & pa.Signature & "&id=" & .Name & "&name=" & .Description & vbCrLf

                End With
            Next i
        End If
    End With

    ' /* add notification content and footer */

    szc = szc & "notify?app-sig=" & pa.Signature & uTranslateIcon(Content.IconPath) & "&" & sz & vbCrLf & _
                "END" & vbCrLf

'        Debug.Print "---"
'        Debug.Print szc
'        Debug.Print "---"

    uCreateSNP3Forward = szc

End Function

Private Function uCreateGNTPForward(ByRef Content As T_NOTIFICATION_INFO) As String
Dim szc As String
Dim sz As String

    sz = Content.OriginalContent
    sz = Replace$(sz, "::", "=")
    sz = Replace$(sz, "#?", "&")
    sz = Replace$(sz, vbCrLf, "\n")

Dim pa As TApp

    Set pa = Content.ClassObj.App

    ' /* base content */

    szc = "GNTP/1.0 REGISTER NONE" & vbCrLf & _
          "Application-Name: " & pa.Name & vbCrLf & _
          "Notifications-Count: " & CStr(pa.CountAlerts) & vbCrLf & vbCrLf

    ' /* add classes */

Dim i As Long

    With pa
        If .CountAlerts Then
            For i = 1 To .CountAlerts
                With .AlertAt(i)
                    szc = szc & "Notification-Name: " & .Name & vbCrLf & _
                                "Notification-Display-Name: " & .Description & vbCrLf & _
                                "Notification-Enabled: " & IIf(.IsEnabled, "True", "False") & vbCrLf & vbCrLf

                End With
            Next i
        End If
    End With

    ' /* add notification content and footer */

    szc = szc & vbCrLf & vbCrLf

    uCreateGNTPForward = szc

End Function

Private Function uCreateGNTPNotify(ByRef Content As T_NOTIFICATION_INFO) As String
Dim szc As String
Dim sz As String

    sz = Content.OriginalContent
    sz = Replace$(sz, "::", "=")
    sz = Replace$(sz, "#?", "&")
    sz = Replace$(sz, vbCrLf, "\n")

Dim pa As TApp

    Set pa = Content.ClassObj.App

    ' /* base content */

    szc = "GNTP/1.0 NOTIFY NONE" & vbCrLf & _
          "Application-Name: " & pa.Name & vbCrLf


'Notification-Name: <string>
'Required - The name (type) of the notification (must match a previously registered notification name registered by the application specified in Application-Name)
'
'Notification-ID: <string>
'Optional - A unique ID for the notification. If used, this should be unique for every request, even if the notification is replacing a current notification (see Notification-Coalescing-ID)
'
'Notification-Title: <string>
'Required - The notification's title
'
'Notification-Text: <string>
'Optional - The notification's text. (defaults to "")
'
'Notification-Sticky: <boolean>
'Optional - Indicates if the notification should remain displayed until dismissed by the user. (default to False)
'
'Notification-Priority: <int>
'Optional - A higher number indicates a higher priority. This is a display hint for the receiver which may be ignored. (valid values are between -2 and 2, defaults to 0)
'
'Notification-Icon: <url> | <uniqueid>
'Optional - The icon to display with the notification.
'
'Notification-Coalescing-ID: <string>
'Optional - If present, should contain the value of the Notification-ID header of a previously-sent notification. This serves as a hint to the notification system that this notification should replace/update the matching previous notification. The notification system may ignore this hint.
'
'Notification-Callback-Context: <string>
'Optional - Any data (will be passed back in the callback unmodified)
'
'Notification-Callback-Context-Type: <string>
'Optional, but Required if 'Notification-Callback-Context' is passed - The type of data being passed in Notification-Callback-Context (will be passed back in the callback unmodified). This does not need to be of any pre-defined type, it is only a convenience to the sending application.
'
'Notification-Callback-Target: <string>
'Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the callback over the original socket will be ignored and the callback data will be passed


    ' /* end marker */

    szc = szc & vbCrLf & vbCrLf

    uCreateGNTPNotify = szc

End Function

