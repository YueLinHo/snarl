VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNetworkPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mPage As BPrefsPage

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg

    On Error Resume Next

    With mPage

        .SetMargin 96
        .Add new_BPrefsControl("banner", "", "Receiving")

        Set pm = New CTempMsg
        pm.Add "text", "Listen for incoming Growl or Snarl notifications?"
        pm.Add "align", 1
        .Add new_BPrefsControl("fancytoggle", "listen_for_snarl", , "", g_ConfigGet("listen_for_snarl"), pm)

        pm.Replace "text", "Listen for incoming JSON notifications?"
        .Add new_BPrefsControl("fancytoggle", "listen_for_json", , , g_ConfigGet("listen_for_json"), pm)

        .Add new_BPrefsControl("label", "", "Snarl will listen on ports 9887 and 23053 for incoming notifications, and port 9889 for incoming JSON notifications.  You may need to adjust your firewall settings in order to receive these.")

'        .Add new_BPrefsControl("fancytoggle2", "haycart", "Enable HAYBARN monitoring?", "", "0", , False)
'        .Add new_BPrefsControl("fancytoggle2", "haycart", "Enable CORNFLOWER monitoring?", "", "0", , False)


'        .Add new_BPrefsControl("separator", "")
''        pm.Replace "text", "Forward notifications?"
''        .Add new_BPrefsControl("fancytoggle", "ft>forward", , "Sending:", "0", pm, False)

        .Add new_BPrefsControl("banner", "", "Sending")
        .Add new_BPrefsControl("label", "", "Snarl will forward notifications to the following IP addresses:")
        Set pm = New CTempMsg
        pm.Add "plus_minus", 1
        Set pc = new_BPrefsControl("listbox", "lb>forward", g_GetRemoteComputers(), , , pm)
        .Add pc

        .Add new_BPrefsControl("fancyplusminus", "fpm>forward")

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pc As BControl
Dim i As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()
    Case "ft>forward"
        If Control.Page.Find("lb>forward", pc) Then _
            pc.SetEnabled (Value = "1")

        If Control.Page.Find("fpm>forward", pc) Then _
            pc.SetEnabled (Value = "1")

    Case "listen_for_json", "listen_for_snarl"
        g_ConfigSet Control.GetName, Value
        If Control.GetName = "listen_for_json" Then
            frmAbout.EnableJSON (Value = "1")

        Else
            frmAbout.EnableSNP (Value = "1")

        End If

    Case "fpm>forward"
        If Value = "+" Then
            ' /* display 'Add New Computer' dialog */
            With New TAddRemoteComputerPanel
                .Go Control.Page.Panel.hWnd

            End With

        Else
            ' /* remove selected entry */
            If mPage.Find("lb>forward", pc) Then
                i = Val(pc.GetValue())
                If i > 0 Then
                    gRemoteComputers.Remove i
                    frmAbout.bUpdateRemoteComputerList
                    g_WriteConfig

                End If
            End If

        End If

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

    Select Case Control.GetName()
    Case "lb>forward"
        If Notification = "update_list" Then
            Control.SetText g_GetRemoteComputers


        End If

    End Select

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
'    mPage.SetMargin 0

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

End Sub

Private Sub KPrefsPage_Detached()

End Sub

Private Sub KPrefsPage_GetPreferredSize(Width As Long, Height As Long)

End Sub

Private Function KPrefsPage_hWnd() As Long

End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)

End Sub








