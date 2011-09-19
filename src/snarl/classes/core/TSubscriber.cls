VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TSubscriber.cls
    '/
    '/  Description:    Represents a subscribed GNTP and SNP3 client
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Enum SN_SUBSCRIBER_TYPE
    ' /* these are dynamic and thus not recorded in network.snarl */
    SN_ST_SNP3_SUBSCRIBER = &H7         ' // subscription made to this computer from a remote computer
    SN_ST_GNTP_SUBSCRIBER = &HE         ' // GNTP subscription made to this computer from a remote computer

    ' /* these are persistent and thus are recorded in network.snarl */
    SN_ST_SNP3_FORWARD = &H100          ' // persistent forward from this computer to a remote computer via SNP3
    SN_ST_GNTP_FORWARD = &H115
    SN_ST_SNP3_SUBSCRIPTION = &H10A     ' // persistent subscription from this computer to a remote computer via SNP3
    SN_ST_GNTP_SUBSCRIPTION = &H11E

    SN_ST_PROTOCOL_MASK = &H10
    SN_ST_DIRECTION_MASK = &HC          ' // <>0=subscription, 0=forward

    ' /* special types */
    
    SN_ST_WEB_LISTENER = &H3FFE

End Enum

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1
Dim WithEvents theTimer As BTimer
Attribute theTimer.VB_VarHelpID = -1

Dim mType As SN_SUBSCRIBER_TYPE
'Dim mData As BPackedData
Dim mApps As BTagList               ' // which app(s) we're subscribed to - an empty list means "all"

Dim mRemoteAddr As String           ' // where to forward notifications or computer to subscribe to
Dim mDescription As String
Dim mPassword As String
Dim mGuid As String

'Dim mContent As BTagList
Dim mForwarders As BTagList
Dim mSuccess As Boolean
'Dim mBusy As Boolean

Dim mEnabled As Boolean
Dim mConfig As ConfigSection

    ' /* a word on granularity: a subscriber is an IP address and a port and any number of apps - password is per subscriber */

Implements BTagItem

Private Function BTagItem_Name() As String

'    BTagItem_Name = theSocket.RemoteHostIP & ":" & CStr(theSocket.RemotePort)
    BTagItem_Name = mGuid

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mType

End Function

Private Sub Class_Initialize()

    mEnabled = True     ' // temp

End Sub

Private Sub theSocket_OnClose()

    Set theSocket = Nothing
'    MsgBox "Closed"

'    g_SubsRoster.Remove theSocket, mType, mData
'    frmAbout.SubscribersChanged

End Sub

Public Function SetFrom(ByRef Config As ConfigSection) As Boolean
Dim aType As SN_SUBSCRIBER_TYPE

    ' /* configure from an existing (or recently created ConfigSection).  This implies
    '    persistency and thus means we're either a forwarder or a subscription to a
    '    remote computer */

    aType = g_SafeLong(Config.GetValueWithDefault("type", ""))

    Select Case aType
    Case SN_ST_GNTP_FORWARD, SN_ST_GNTP_SUBSCRIPTION, SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION

    Case Else
        g_Debug "TSubscriber.SetFrom(): type must be one of SN_ST_GNTP_FORWARD, SN_ST_GNTP_SUBSCRIPTION, SN_ST_SNP3_FORWARD, SN_ST_SNP3_SUBSCRIPTION", LEMON_LEVEL_CRITICAL
        Exit Function

    End Select

Dim sDest As String

    sDest = Config.GetValueWithDefault("remotehost", "")
    mEnabled = (Config.GetValueWithDefault("enabled", "0") = "1")

    If Not (Config Is Nothing) Then _
        SetFrom = uInit(Config.Name, aType, Config.GetValueWithDefault("description", ""), _
                        sDest, Config.GetValueWithDefault("apps", ""), Config)

    If Not SetFrom Then _
        Exit Function

    ' /* if it's a subscription, we need to create the socket and send the request */

    Select Case aType

    Case SN_ST_SNP3_SUBSCRIPTION, SN_ST_GNTP_SUBSCRIPTION
        g_Debug "TSubscriber.SetFrom(): connecting to '" & sDest & "'..."
        Set theSocket = New CSocket
        theSocket.Connect sDest, IIf(aType = SN_ST_SNP3_SUBSCRIPTION, SNP_DEFAULT_PORT, GNTP_DEFAULT_PORT)

    End Select

End Function

Public Function SetTo(ByVal Guid As String, ByVal SubType As SN_SUBSCRIBER_TYPE, ByVal Description As String, ByVal RemoteHost As String, ByVal Apps As String, ByRef Settings As ConfigSection, Optional ByRef ProvidedSocket As CSocket) As Boolean

    ' /* configure from the provided information.  This implies transiency and thus
    '    means we're recording a subscription received from a remote computer */

    ' /* error checking */

    If (SubType <> SN_ST_GNTP_SUBSCRIBER) Or (SubType <> SN_ST_SNP3_SUBSCRIBER) Then
        g_Debug "TSubscriber.SetTo(): SubType can only be SN_ST_GNTP_SUBSCRIBER or SN_ST_SNP3_SUBSCRIBER", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    SetTo = uInit(Guid, SubType, Description, RemoteHost, Apps, Nothing, ProvidedSocket)

End Function

Public Function uInit(ByVal Guid As String, ByVal SubType As SN_SUBSCRIBER_TYPE, ByVal Description As String, ByVal RemoteHost As String, ByVal Apps As String, ByRef Settings As ConfigSection, Optional ByRef Socket As CSocket) As Boolean

    ' /* error checking */

    If (Guid = "") Or (SubType = 0) Or (RemoteHost = "") Or (InStr(get_ip_address_table(), RemoteHost) <> 0) Then _
        Exit Function

    mGuid = Guid
    mType = SubType
    mDescription = Description
    mRemoteAddr = RemoteHost
    Set mConfig = Settings

    Set theSocket = Socket
    Set mApps = new_BTagList()
    Set mForwarders = new_BTagList()

    uInit = True

End Function


'Public Sub SetTo(ByRef Socket As CSocket, ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Data As BPackedData)
'
'    Set theSocket = Socket
'    mType = SubscriberType
'    Set mData = Data
'    Set mApps = new_BTagList()
'
'    mDest = Socket.RemoteHostIP & " (" & theSocket.RemoteHost & ")"
'
'    If Not (mData Is Nothing) Then
'        mPassword = Data.ValueOf("password")
'        AddApps mData
'
'    Else
'        g_Debug "TSubscriber.SetTo(): Data missing", LEMON_LEVEL_WARNING
'
'    End If
'
'End Sub

Public Function Forward(ByVal Content As String) As Boolean
Dim pf As TForwarder

    Select Case mType
    
    Case SN_ST_SNP3_FORWARD
        ' /* create a new forwarder: this will send the content
        '    once it has connected successfully */
        Set pf = New TForwarder
        If pf.Init(g_CreateGUID(True), Me, Content, False) Then _
            mForwarders.Add pf

    Case SN_ST_SNP3_SUBSCRIBER

'        If theSocket.State <> sckConnected Then
'            g_Debug "TSubscriber.Forward(): no longer connected to subscriber", LEMON_LEVEL_CRITICAL
'            ' /* TO DO: remove subscriber */
'            Exit Function
'
'        End If
'
'        ' /* SNP 3.0 subscription (sent as FORWARD) */
'
'        sz = "SNP/3.0 FORWARD" & vbCrLf
'
'        ' /* add app info */
'
'        sz = sz & "app-signature: " & App.Signature & vbCrLf
'        sz = sz & "app-name: " & App.Name & vbCrLf
'
'        ' /* add selected parts of the content */
'
''        Set ppd = New BPackedData
''        content.SetTo Content
'
'        ' /* standard components */
'
'        sz = sz & "class: " & Content.ValueOf("id") & vbCrLf
'        sz = sz & "title: " & Replace$(Content.ValueOf("title"), vbCrLf, "\n") & vbCrLf
'        sz = sz & "text: " & Replace$(Content.ValueOf("text"), vbCrLf, "\n") & vbCrLf
'        sz = sz & "priority: " & Content.ValueOf("priority") & vbCrLf
'        sz = sz & "uid: " & Content.ValueOf("uid") & vbCrLf                 ' // TO DO: should we prefix with sender host name?
'
'        ' /* icon */
'
'        szn = Content.ValueOf("icon")
'        If szn <> "" Then
'            If (LCase$(g_SafeLeftStr(szn, 7)) = "http://") Or _
'               (LCase$(g_SafeLeftStr(szn, 8)) = "https://") Or _
'               (LCase$(g_SafeLeftStr(szn, 1)) = "!") Then
'                ' /* add verbatim */
'                sz = sz & "icon: " & szn & vbCrLf
'
'            Else
'                ' /* encode it in a slightly modified Base64 format (CRLF's are replaced with #'s) */
'                If uEncodeIcon(szn, szv) Then _
'                    sz = sz & "icon-phat64: " & szv & vbCrLf
'
'            End If
'
'            szn = ""
'            szv = ""
'
'        End If
'
'        ' /* add data-*, label-* and value-* elements */
'
'        With Content
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                If (LCase$(g_SafeLeftStr(szn, 5)) = "data-") Or _
'                   (LCase$(g_SafeLeftStr(szn, 6)) = "label-") Or _
'                   (LCase$(g_SafeLeftStr(szn, 6)) = "value-") Then _
'                    sz = sz & szn & ": " & szv & vbCrLf
'
'            Loop
'
'        End With
'
'        ' /* standard headers */
'
'        sz = sz & "x-timestamp: " & Format$(Now(), "d mmm yyyy hh:mm:ss") & vbCrLf
'        sz = sz & "x-daemon: " & "Snarl " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & vbCrLf
'        sz = sz & "x-host: " & LCase$(g_GetComputerName()) & vbCrLf
'
'        sz = sz & "END" & vbCrLf
'        theSocket.SendData sz
'
'
'    Case SN_ST_GNTP




    End Select

End Function

Public Function Destination() As String

    Destination = mRemoteAddr

End Function

Public Function UID() As String

    UID = BTagItem_Name()

End Function

Public Function SubscriptionType() As SN_SUBSCRIBER_TYPE

    SubscriptionType = mType

End Function

Public Function Detail() As String
Dim pt As BTagItem
Dim i As Long

    Detail = IIf(mType = SN_ST_SNP3_FORWARD, "Forwarding ", "Subscribed to ")

    If mApps.CountItems = 0 Then
        Detail = Detail & "everything"

    Else
        With mApps
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                i = i + 1
                If i > 4 Then
                    Detail = Detail & "and " & CStr(mApps.CountItems - 4) & " more, "
                    Exit Do

                Else
                    Detail = Detail & pt.Name & ", "

                End If

            Loop

        End With

        Detail = g_SafeLeftStr(Detail, Len(Detail) - 2)

    End If

    Detail = Detail & IIf((mType And SN_ST_DIRECTION_MASK) = 0, " to ", " from ") & mRemoteAddr _
                    & " via " & IIf((mType And SN_ST_PROTOCOL_MASK) = 0, "SNP", "GNTP")

End Function

Public Function Socket() As CSocket

    Set Socket = theSocket

End Function

Public Function AddApps(ByRef Data As BPackedData) As Boolean
Dim sPassword As String
Dim szn As String
Dim szv As String

    ' /* returns TRUE if at least one subscription succeeded */

    If Data.Exists("password") Then
        If Data.ValueOf("password") <> mPassword Then
            g_Debug "TSubscriber.AddApps(): password mismatch", LEMON_LEVEL_CRITICAL
            gSetLastError SNARL_ERROR_AUTH_FAILURE
            Exit Function

        End If

    End If

    With Data
        .Rewind
        Do While .GetNextItem(szn, szv)
            If szn = "app-sig" Then
                If mApps.IndexOf(szv) = 0 Then
                    mApps.Add new_BTagItem(szv, "")
                    g_Debug "TSubscriber.AddApps(): subscribed to '" & szv & "'" '& IIf(sPassword <> "", " (using password)", "")
                    AddApps = True

                Else
                    g_Debug "TSubscriber.AddApps(): already subscribed to '" & szv & "'", LEMON_LEVEL_WARNING

                End If
            End If
        Loop
    End With

End Function

Public Function Password() As String

    Password = mPassword

End Function

Public Function IsSubscribedTo(ByVal AppSig As String) As Boolean

    If mApps.CountItems = 0 Then
        IsSubscribedTo = True

    Else
        IsSubscribedTo = (mApps.IndexOf(AppSig) <> 0)

    End If

End Function

Public Function Apps() As BTagList

    Set Apps = mApps

End Function

Public Function TypeStr() As String

    Select Case mType
    Case SN_ST_GNTP_SUBSCRIBER

    Case SN_ST_SNP3_SUBSCRIBER

    Case SN_ST_SNP3_FORWARD
        TypeStr = "FWD/SNP"
    
    Case SN_ST_SNP3_SUBSCRIPTION

    Case Else
        TypeStr = "UNK?"

    End Select

End Function

Public Function Description() As String

    If mDescription <> "" Then
        Description = mDescription

    Else
        Description = mRemoteAddr

    End If

End Function

Public Function Guid() As String

    Guid = mGuid

End Function

Private Sub theSocket_OnConnect()

    Select Case mType
    Case SN_ST_SNP3_SUBSCRIPTION
        g_Debug "TSubscriber.OnConnect(): SNP3 subscription here", LEMON_LEVEL_INFO

    Case SN_ST_GNTP_SUBSCRIPTION
        g_Debug "TSubscriber.OnConnect(): GNTP subscriptions are not currently implemented", LEMON_LEVEL_WARNING

    Case Else
        g_Debug "TSubscriber.OnConnect(): spurious connection (dest=" & mRemoteAddr & " type=" & g_HexStr(mType, 2) & ")", LEMON_LEVEL_WARNING

    End Select
    
    theSocket.CloseSocket       ' // temp
    Set theTimer = Nothing

End Sub

Private Sub theSocket_OnSendComplete()

'///

End Sub

Private Sub theTimer_Pulse()

    g_Debug "TSubscriber.Pulse(): timed out connecting to '" & mRemoteAddr & "' (type=" & g_HexStr(mType, 2) & ")", LEMON_LEVEL_CRITICAL
    theSocket.CloseSocket
    Set theSocket = Nothing

End Sub



'Private Sub uSendNext()
'Dim sz As String
'
'    If mContent.CountItems > 0 Then
'        mBusy = True
'        sz = mContent.TagAt(1).Value
'        mContent.Remove 1
'        Debug.Print "^sending '" & sz & "'...^"
'        theSocket.SendData sz
'
'    Else
'        mBusy = False
''        theSocket.CloseSocket
''        Set theSocket = Nothing
'
'    End If
'
'End Sub

Public Sub SetEnabled(ByVal Enabled As Boolean)

    mEnabled = Enabled
    Debug.Print mDescription & " enabled: " & mEnabled

    If Not (mConfig Is Nothing) Then
        mConfig.Update "enabled", IIf(Enabled, "1", "0")
        g_SubsRoster.WriteConfig

    End If

End Sub

Public Function IsEnabled() As Boolean

    IsEnabled = mEnabled

End Function

Public Sub RemoveForward(ByVal Guid As String)

    mForwarders.Remove mForwarders.IndexOf(Guid)

End Sub
