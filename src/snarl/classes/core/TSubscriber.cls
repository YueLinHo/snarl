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
    SN_ST_SNP3_SUBSCRIBER = 0
    SN_ST_GNTP = 1

    SN_ST_SNP3_FORWARDER = &H80&     ' // means the socket is under our control

End Enum

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1
Dim mType As SN_SUBSCRIBER_TYPE
Dim mData As BPackedData
Dim mApps As BTagList               ' // which app(s) we're subscribed to - an empty list means "all"
Dim mDest As String

Dim mPassword As String

    ' /* a word on granularity: a subscriber is an IP address and a port and any number of apps - password is per subscriber */

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = theSocket.RemoteHostIP & ":" & CStr(theSocket.RemotePort)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mType

End Function

Public Sub SetTo(ByRef Socket As CSocket, ByVal SubscriberType As SN_SUBSCRIBER_TYPE, ByRef Data As BPackedData)

    Set theSocket = Socket
    mType = SubscriberType
    Set mData = Data
    Set mApps = new_BTagList()

    mDest = Socket.RemoteHostIP & " (" & theSocket.RemoteHost & ")"

    If Not (mData Is Nothing) Then
        mPassword = Data.ValueOf("password")
        AddApps mData

    Else
        g_Debug "TSubscriber.SetTo(): Data missing", LEMON_LEVEL_WARNING

    End If

End Sub

Public Function send(ByVal Content As String, ByRef App As TApp) As Boolean

    If theSocket.State <> sckConnected Then
        g_Debug "TSubscriber.Send(): no longer connected to subscriber", LEMON_LEVEL_CRITICAL
        ' /* TO DO: remove subscriber */
        Exit Function

    End If

Dim ppd As BPackedData
Dim szn As String
Dim szv As String
Dim sz As String

    Select Case mType
    Case SN_ST_SNP3_SUBSCRIBER

        ' /* SNP 3.0 subscription (sent as FORWARD) */

        sz = "SNP/3.0 FORWARD" & vbCrLf

        ' /* add app info */

        sz = sz & "app-signature: " & App.Signature & vbCrLf
        sz = sz & "app-name: " & App.Name & vbCrLf

        ' /* add selected parts of the content */

        Set ppd = New BPackedData
        ppd.SetTo Content

        ' /* standard components */

        sz = sz & "class: " & ppd.ValueOf("id") & vbCrLf
        sz = sz & "title: " & Replace$(ppd.ValueOf("title"), vbCrLf, "\n") & vbCrLf
        sz = sz & "text: " & Replace$(ppd.ValueOf("text"), vbCrLf, "\n") & vbCrLf
        sz = sz & "priority: " & ppd.ValueOf("priority") & vbCrLf
        sz = sz & "uid: " & ppd.ValueOf("uid") & vbCrLf                 ' // TO DO: should we prefix with sender host name?

        ' /* icon */

        szn = ppd.ValueOf("icon")
        If szn <> "" Then
            If (LCase$(g_SafeLeftStr(szn, 7)) = "http://") Or _
               (LCase$(g_SafeLeftStr(szn, 8)) = "https://") Or _
               (LCase$(g_SafeLeftStr(szn, 1)) = "!") Then
                ' /* add verbatim */
                sz = sz & "icon: " & szn & vbCrLf

            Else
                ' /* encode it in a slightly modified Base64 format (CRLF's are replaced with #'s) */
                If uEncodeIcon(szn, szv) Then _
                    sz = sz & "icon-phat64: " & szv & vbCrLf

            End If

            szn = ""
            szv = ""

        End If

        ' /* add data-*, label-* and value-* elements */

        With ppd
            .Rewind
            Do While .GetNextItem(szn, szv)
                If (LCase$(g_SafeLeftStr(szn, 5)) = "data-") Or _
                   (LCase$(g_SafeLeftStr(szn, 6)) = "label-") Or _
                   (LCase$(g_SafeLeftStr(szn, 6)) = "value-") Then _
                    sz = sz & szn & ": " & szv & vbCrLf

            Loop

        End With

        ' /* standard headers */

        sz = sz & "x-timestamp: " & Format$(Now(), "d mmm yyyy hh:mm:ss") & vbCrLf
        sz = sz & "x-daemon: " & "Snarl " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & vbCrLf
        sz = sz & "x-host: " & LCase$(g_GetComputerName()) & vbCrLf

        sz = sz & "END" & vbCrLf
        theSocket.SendData sz


    Case SN_ST_GNTP
    
    
    
    Case SN_ST_SNP3_FORWARDER
        ' /* connect to destination, craft message content and send */


    End Select

End Function

Public Function Destination() As String

    Destination = mDest

End Function

Public Function UID() As String

    UID = BTagItem_Name()

End Function

Public Function SubscriptionType() As SN_SUBSCRIBER_TYPE

    SubscriptionType = mType

End Function

Public Function SubscribedTo() As String
Dim pt As BTagItem
Dim i As Long

    If mApps.CountItems = 0 Then
        SubscribedTo = "Subscribed to all"

    Else

        SubscribedTo = "Subscribed to "

        With mApps
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                i = i + 1
                If i > 4 Then
                    SubscribedTo = SubscribedTo & "and " & CStr(mApps.CountItems - 4) & " more, "
                    Exit Do

                Else
                    SubscribedTo = SubscribedTo & pt.Name & ", "

                End If

            Loop

        End With

        SubscribedTo = g_SafeLeftStr(SubscribedTo, Len(SubscribedTo) - 2)

    End If

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

Private Sub theSocket_OnClose()

    g_SubsRoster.Remove theSocket, mType, mData
    frmAbout.SubscribersChanged

End Sub

Private Function uEncodeIcon(ByVal IconPath As String, ByRef Base64 As String) As Boolean
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
        g_Debug "TSubscriber.uEncodeIcon(): " & err.Description, LEMON_LEVEL_CRITICAL
    
    End If

End Function

Public Function Apps() As BTagList

    Set Apps = mApps

End Function

Public Function SubscriptionTypeStr() As String

    If mType = SN_ST_GNTP Then
        SubscriptionTypeStr = "GNTP"

    ElseIf mType = SN_ST_SNP3_SUBSCRIBER Then
        SubscriptionTypeStr = "SNP"

    ElseIf mType = SN_ST_SNP3_FORWARDER Then
        SubscriptionTypeStr = "FWD"

    Else
        SubscriptionTypeStr = "UNK?"

    End If

End Function
