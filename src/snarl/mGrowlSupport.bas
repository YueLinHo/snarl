Attribute VB_Name = "mGrowlSupport"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mGrowlSupport.bas
    '/
    '/  Description:    Growl UDP support routines
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Type T_GROWL_NOTIFICATION_TYPE
    Name As String
    Enabled As Boolean

End Type

Private Type T_GROWL_NOTIFICATION
    Sticky As Boolean
    Priority As Integer
    TypeName As String
    Title As String
    Description As String
    AppName As String
    Sender As String

End Type

Private Type T_GROWL_REGISTRATION
    ApplicationName As String
    RemoteHost As String
    NumTypes As Integer
    NotificationType() As T_GROWL_NOTIFICATION_TYPE
    SnarlToken As Long

End Type

Dim mGrowlReg() As T_GROWL_REGISTRATION
Dim mRegCount As Long

Public Function g_ProcessGrowlUDP(ByRef Data() As Byte, ByVal cb As Long, ByVal Sender As String)
Dim b() As Byte
Dim w As Integer
Dim l As Long
Dim i As Long
Dim z As Byte
Dim p As Byte

    b() = Data()

    For i = 0 To cb - 1
'        Debug.Print b(i) & Chr$(b(i)) & " ";
        Debug.Print Chr$(b(i));

    Next i
    Debug.Print

    ' /* get version */

    CopyMemory z, b(0), 1
'    Debug.Print "version: " & CStr(z)
    If z <> 1 Then
        Debug.Print "bad version"
        Exit Function

    End If

    ' /* get packet type */

    CopyMemory z, b(1), 1
'    Debug.Print "type: " & CStr(z)

    Select Case z
    Case 0
        ' /* registration */

Dim pgr As T_GROWL_REGISTRATION

        ' /* Byte   Description
        '     0     Protocol version (always 1)
        '     1     Packet type (0 = registration, 1 = notification)
        '     2     Application name length (16 bits)
        '     4     Number of notification types
        '     5     Number of enabled notification types
        '     6     Application name
        '    ...
        '    ...    Notification Types
        '               0   Type name size
        '               2   Name
        '
        ' */

        l = VarPtr(b(2))            ' // app name length
        w = uGetInt(l)
        l = l + 2                   ' // skip types and enabled counts...

        pgr.ApplicationName = uGetString(l, w)
        pgr.RemoteHost = Sender
        Debug.Print "REGISTER: App=" & pgr.ApplicationName & " sender=" & pgr.RemoteHost

        ' /* get number of notification types */

        CopyMemory z, b(4), 1
        Debug.Print "Notification types: " & z

        ' /* read each notification type record */

        l = VarPtr(b(6)) + w
        For i = 0 To (z - 1)
            With pgr
                .NumTypes = .NumTypes + 1
                ReDim Preserve .NotificationType(.NumTypes)
                With .NotificationType(.NumTypes)
                    w = uGetInt(l)                  ' // name length (bytes)
                    .Name = uGetString(l, w)        ' // name (decoded from UTF8)
                    .Enabled = False
                    Debug.Print CStr(i + 1) & "='" & .Name & "'"

                End With
            End With
        Next i

        ' /* get number of enabled notifications */

        CopyMemory z, b(5), 1
        Debug.Print "Enabled notifications: " & z

        For i = 0 To (z - 1)
            CopyMemory p, ByVal l, 1
            Debug.Print "Notification '" & CStr(p) & "' is enabled"
            pgr.NotificationType(p + 1).Enabled = True
            l = l + 1

        Next i

        ' /* register with Snarl */
        uRegister pgr


    Case 1
        ' /* notification */

Dim pgn As T_GROWL_NOTIFICATION

        ' /* Byte   Description
        '     0     Protocol version (always 1)
        '     1     Packet type (0 = registration, 1 = notification)
        '     2     Flags (16 bits)
        '     4     Notification type name length
        '     6     Title length
        '     8     Description length
        '     A     Application name length
        '     C     Notification type name
        '
        ' */

        Debug.Print "NOTIFICATION"

        ' /* get flags */

        l = VarPtr(b(2))
        w = uGetInt(l)
        pgn.Sticky = (w And 1)

'        If (w And 1) Then
'            Debug.Print "sticky"
'
'        Else
'            Debug.Print "not sticky"
'
'        End If

Dim fNeg As Boolean

        fNeg = ((w And 8) = 8)
        w = (w And &HE) / 2
        If fNeg Then _
            w = (-(w Xor 7)) - 1

        pgn.Priority = w

Dim sptr As Long

        sptr = VarPtr(b(12))                        ' // start of string table

        w = uGetInt(l)
        Debug.Print "type name len: " & w
        pgn.TypeName = uGetString(sptr, w)

        w = uGetInt(l)
        Debug.Print "title len: " & w
        pgn.Title = uGetString(sptr, w)

        w = uGetInt(l)
        Debug.Print "description len: " & w
        pgn.Description = uGetString(sptr, w)

        w = uGetInt(l)
        Debug.Print "app name len: " & w
        pgn.AppName = uGetString(sptr, w)

        Debug.Print "sticky=" & CStr(pgn.Sticky) & " priority=" & CStr(pgn.Priority)
        Debug.Print "type name=" & pgn.TypeName & " app=" & pgn.AppName
        Debug.Print "title=" & pgn.Title & " desc=" & pgn.Description

        pgn.Sender = Sender

        i = uIndexOfApp(pgn.AppName)
        If i = 0 Then
            g_Debug "mGrowlSupport.g_ProcessGrowlUDP(): '" & pgn.AppName & "' is not registered", LEMON_LEVEL_CRITICAL

        Else
            uNotify mGrowlReg(i).SnarlToken, pgn

        End If


    Case Else
        Debug.Print "bad packet type"

    End Select

End Function

Public Sub g_GrowlUDPClose()
Static i As Long

    If mRegCount Then
        For i = mRegCount To 1 Step -1
            If mGrowlReg(i).SnarlToken Then _
                g_AppRoster.Unregister mGrowlReg(i).SnarlToken, ""

        Next i
    End If

End Sub

Private Function uGetInt(ByRef lptr As Long) As Integer

    CopyMemory uGetInt, ByVal lptr, 2
    SWAPBYTES uGetInt
    lptr = lptr + 2

End Function

Private Function uGetString(ByRef lptr As Long, ByVal Length As Long) As String
Dim sz As String

    sz = String$(Length, 0)
    CopyMemory ByVal StrPtr(sz), ByVal lptr, Length
    uGetString = g_toUnicodeUTF8(g_TrimStr(StrConv(sz, vbUnicode)))
    lptr = lptr + Length

End Function

Private Function uIndexOf(ByVal Token As Long) As Long
Static i As Long

    If mRegCount Then
        For i = 1 To mRegCount
            If mGrowlReg(i).SnarlToken = Token Then
                uIndexOf = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Function uIndexOfApp(ByVal Name As String) As Long
Dim i As Long

    If mRegCount Then
        For i = 1 To mRegCount
            If mGrowlReg(i).ApplicationName = Name Then
                uIndexOfApp = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Sub uRegister(ByRef pgr As T_GROWL_REGISTRATION)
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.Add "id", Replace$(pgr.ApplicationName, " ", "_") & "_" & pgr.RemoteHost
    ppd.Add "title", pgr.ApplicationName
    ppd.Add "icon", g_MakePath(App.Path) & "etc\icons\growl_app.png"

    g_Debug "mGrowlSupport.uRegister(): '" & pgr.ApplicationName & "_" & pgr.RemoteHost & "'"

    pgr.SnarlToken = g_AppRoster.Add41(ppd, (pgr.RemoteHost <> "localhost"))
    If pgr.SnarlToken = 0 Then
        g_Debug "mGrowlSupport.uRegister(): failed to register with Snarl", LEMON_LEVEL_CRITICAL

    Else
        If uIndexOf(pgr.SnarlToken) = 0 Then
            ' /* add to our list */
            mRegCount = mRegCount + 1
            ReDim Preserve mGrowlReg(mRegCount)
            LSet mGrowlReg(mRegCount) = pgr
            g_Debug "mGrowlSupport.uRegister(): registered as " & g_HexStr(pgr.SnarlToken)

            ' /* register notification types - think Growl does this each time, should we? */

            uAddClasses pgr

        Else
            g_Debug "mGrowlSupport.uRegister(): already registered as " & g_HexStr(pgr.SnarlToken)

        End If

    End If

End Sub

Private Sub uAddClasses(ByRef pgr As T_GROWL_REGISTRATION)
Dim ppd As BPackedData
Static i As Long
Dim pa As TApp

    If g_AppRoster.FindByToken(pgr.SnarlToken, pa, "") Then
        For i = 1 To pgr.NumTypes
            Set ppd = New BPackedData
            ppd.Add "id", pgr.NotificationType(i).Name
            ppd.Add "name", pgr.NotificationType(i).Name
            ppd.Add "enabled", IIf(pgr.NotificationType(i).Enabled, "1", "0")
            pa.AddClass ppd

        Next i

    Else
        g_Debug "mGrowlSupport.uAddClasses(): app '" & g_HexStr(pgr.SnarlToken) & "' not found", LEMON_LEVEL_CRITICAL
    
    End If

End Sub

Private Function uNotify(ByVal AppToken As Long, ByRef pgn As T_GROWL_NOTIFICATION) As Long
Dim ppd As BPackedData

    Set ppd = New BPackedData
    With ppd
        .Add "id", pgn.TypeName
        .Add "title", pgn.Title
        .Add "text", pgn.Description
        .Add "timeout", IIf(pgn.Sticky, "0", "-1")
        ' /* Growl supports two additional priorities that we don't */
'        .Add "priority", IIf(pgn.Priority > 0, "1", IIf(pgn.Priority < 0, "-1", "0"))
        .Add "priority", IIf(pgn.Priority > 0, "1", CStr(pgn.Priority))                 ' // R2.4 DR8: pass through unchanged

    End With

    ' /* R2.4.1 - include major version number in this */

    uNotify = g_DoNotify(AppToken, ppd, , NF_REMOTE Or App.Major, pgn.Sender)

End Function
