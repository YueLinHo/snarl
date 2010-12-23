VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TExtPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TExtPage.cls
'/
'/  Description:    The loaded extensions preferences page
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

Dim mPage As BPrefsPage
Dim mCurrent As TExtension

Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()

End Sub

Private Sub KPrefsPage_Attached()
Dim pc As BControl
Dim pm As CTempMsg
Dim i As Long

    On Error Resume Next

    With mPage
        .Add new_BPrefsControl("label", "", "Extensions are modules loaded by Snarl during startup; " & _
                                            "they can register with Snarl and add alerts in exactly the " & _
                                            "same way applications can.  To disable an extension, " & _
                                            "untick it's entry in the list below:")

        Set pm = New CTempMsg
        pm.Add "checkboxes", 1
        pm.Add "item-height", 38
        Set pc = new_BPrefsControl("listbox", "lb>extensions", , , "1", pm)
        pc.SizeTo 0, 160
        .Add pc

        Set pc = new_BPrefsControl("fancytoolbar", "ftb>extension", "Details|Configure...|Show in Explorer")
        prefskit_SetItem pc, 1, "enabled", 0&
        prefskit_SetItem pc, 2, "enabled", 0&
        prefskit_SetItem pc, 3, "enabled", 0&
        .Add pc


'        .Add new_BPrefsControl("fancybutton2", "fb2>ext.cfg", "Configure", , , , False)

'        Set pc = new_BPrefsControl("label", "ext_desc", " ", "Description:")
'        pc.SizeTo 0, 26
'        mPage.Add pc

'        .Add new_BPrefsControl("label", "ext_desc", String$(92, " "), "Description:")
'        .Add new_BPrefsControl("label", "ext_info", String$(92, " "), "Version:")
        .Add new_BPrefsControl("label", "ext_stat", String$(92, " "), "")
'        .Add new_BPrefsControl("separator", "")
        
        .Add new_BPrefsControl("fancytoolbar", "toolbar-general", "Reload Extensions|Install Extension...")

'        .Add new_BPrefsControl("fancybutton2", "fb2>ext.reload", "Reload Extensions")
        .Add new_BPrefsControl("separator", "")

        Set pm = New CTempMsg
        pm.Add "image-file", g_MakePath(App.Path) & "etc\icons\info_emblem.png"
'        pm.Add "valign", "centre"
        Set pc = new_BPrefsControl("labelex", "", "To prevent extensions from being loaded, hold down either Control (Ctrl) key while Snarl loads.", , , pm)
        pc.SizeTo 0, 56
        .Add pc


        .Add new_BPrefsControl("label", "", "")

    End With

    uUpdateExtensions

End Sub

Private Sub KPrefsPage_ControlChanged(Control As BControl, ByVal Value As String)
Dim pmi As OMMenuItem
Dim pm As MMessage
Dim pc As BControl
Dim b As Boolean

    Debug.Print "[" & mPage.GetName() & "]: ControlChanged '" & Control.GetName() & "' >> '" & Value & "'"

    Select Case Control.GetName()
    Case "lb>extensions"
        ' /* extension selected */

        Set mCurrent = Nothing

        If (g_ExtnRoster Is Nothing) Then _
            Exit Sub

        Set mCurrent = melonRosterItemAt(g_ExtnRoster, Val(Value))
        If (mCurrent Is Nothing) Then _
            Exit Sub

'        If mPage.Find("ext_desc", pc) Then _
            pc.SetText px.Description

'        If mPage.Find("ext_info", pc) Then _
            pc.SetText "Version: " & uStr(px.InfoString())

'        If mPage.Find("ext_info", pc) Then _
            pc.SetText px.Copyright

        uUpdateStatus

'        If mPage.Find("fb2>ext.cfg", pc) Then
'            b = False
'            If px.IsEnabled Then _
'                b = (px.Flags And SNARL_EXT_IS_CONFIGURABLE)
'
'            pc.SetEnabled b
'
'        End If

    Case "ftb>extension"

        Select Case Val(Value)
        Case 1
            ' /* details */
            If Not (mCurrent Is Nothing) Then _
                mCurrent.ShowDetails

        Case 2
            ' /* configure */
            If Not (mCurrent Is Nothing) Then _
                mCurrent.DoPrefs mPage.Panel.hWnd

        Case 3
            If Not (mCurrent Is Nothing) Then _
                ShellExecute mPage.Panel.hWnd, "open", g_MakePath(mCurrent.Path), vbNullString, vbNullString, SW_SHOW
            
'            ' /* advanced */
'            With New OMMenu
'                .AddItem .CreateItem("expl", "Show in Explorer...")
'                .AddItem .CreateItem("inst", "Install Extension...")
'
'                Set pmi = .Track(mPage.Panel.hWnd)
'                If Not (pmi Is Nothing) Then
'                    Select Case pmi.Name
'                    Case "expl"
'
'                    Case "inst"
'                        uDoInstallExtension
'
'                    End Select
'                End If
'
'            End With

        End Select


    Case "toolbar-general"

        Select Case Val(Value)
        Case 1
            ' /* reload */
            If Not (g_ExtnRoster Is Nothing) Then
                g_ExtnRoster.Reload
                uUpdateExtensions

            End If

        Case 2
            ' /* install */
            uDoInstallExtension

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As BControl)
Dim px As TExtension
Dim pc As BControl
Dim rv As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlInvoked '" & Control.GetName() & "'"

    Select Case Control.GetName()
'    Case "fb2>ext.reload"
'        ' /* reload extensions */
'        If Not (g_ExtnRoster Is Nothing) Then
'            g_ExtnRoster.Reload
'            uUpdateExtensions
'
'        End If

    Case "fb2>ext.cfg"
        ' /* configure extension */
        If (g_ExtnRoster Is Nothing) Or (mPage Is Nothing) Then _
            Exit Sub

        If Not mPage.Find("lb>extensions", pc) Then _
            Exit Sub

        Set px = melonRosterItemAt(g_ExtnRoster, Val(pc.GetValue()))
        If Not (px Is Nothing) Then _
            px.DoPrefs Control.Page.Panel.hWnd


    Case "lb>extensions"
        ' /* extension double-clicked */

        If Not (mCurrent Is Nothing) Then _
            mCurrent.DoPrefs mPage.Panel.hWnd

'        If (g_ExtnRoster Is Nothing) Then _
'            Exit Sub
'
'        Set px = melonRosterItemAt(g_ExtnRoster, Val(Control.GetValue()))
'        If (px Is Nothing) Then _
'            Exit Sub
'
'        MsgBox px.Path, vbInformation Or vbOKOnly

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As BControl, ByVal Notification As String, Data As melon.MMessage)
Dim bEnable As Boolean
Dim pc As BControl
Dim px As TExtension
Dim pm As CTempMsg
Dim sz As String
Dim i As Long

    Debug.Print "[" & mPage.GetName() & "]: ControlNotify '" & Control.GetName() & "' >> '" & Notification & "'"

    If Control.GetName() = "lb>extensions" Then
        Select Case Notification
        Case "update_list"
            uUpdateExtensions

        Case "checked"
            If (mCurrent Is Nothing) Then _
                Exit Sub

            ' /* construct the extended command message */
            Set pm = New CTempMsg
            pm.Add "what", "checked"
            bEnable = (Control.DoExCmd(B_GET_ITEM, pm) <> "0")  ' // get enable/disable from checkbox setting

            mCurrent.SetEnabled bEnable
            g_ExtnRoster.WriteExcludeList
            uUpdateStatus

            pm.Add "image-file", IIf((mCurrent.IsEnabled And Not mCurrent.StartedOkay), _
                                     g_MakePath(App.Path) & "etc\icons\ext-error.png", mCurrent.Icon)

            pm.Add "subtext", IIf((mCurrent.IsEnabled And Not mCurrent.StartedOkay), _
                                    "Failed to start: " & mCurrent.LastError, mCurrent.Description)

            pm.Add "greyscale", IIf(mCurrent.IsEnabled, 0&, -1&)

            Control.DoExCmd B_SET_ITEM, pm

            If Control.Page.Panel.Find("cb>apps", pc) Then _
                pc.Notify "update_list", Nothing

'                Else
'                    ' /* failed to enable/disable the extension */
'
'
'                    If bEnable Then _
'                        MsgBox "The " & px.FriendlyName & " extension failed to start." & vbCrLf & _
'                               "The following information may be of use:" & vbCrLf & vbCrLf & _
'                               "Name: " & uStr(px.Name) & vbCrLf & _
'                               "Error: " & uStr(px.LastError) & vbCrLf & _
'                               "Version: " & uStr(px.VerString) & vbCrLf & _
'                               "Support contact: " & uStr(px.SupportURL), _
'                               vbCritical Or vbOKOnly, "Extension Failed to Load"
'
'                End If
'            End If
        End Select
    End If

End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 0

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

Private Sub uUpdateExtensions()

    If (g_ExtnRoster Is Nothing) Then _
        Exit Sub

Dim pc As BControl

    If Not (mPage.Find("lb>extensions", pc)) Then _
        Exit Sub

    pc.DoExCmd "lock"


Dim list_content As String
Dim px As TExtension
Dim pm As CTempMsg
Dim pr As MRoster
Dim z As String
Dim i As Long

    z = pc.GetValue()

    list_content = ""
    Set pr = g_ExtnRoster
    pr.Rewind
    Do While pr.NextItem(px)
        list_content = list_content & px.Name & " " & px.VerString(False) & "#?" & CStr(0) & "#?" & px.Description & "|"

    Loop

    ' /* set the content */

    pc.SetText g_SafeLeftStr(list_content, Len(list_content) - 1)

    ' /* set the checkboxes and icons */

Dim sz As String

    i = 0
    pr.Rewind
    Do While pr.NextItem(px)
        i = i + 1

        Set pm = New CTempMsg
        pm.Add "index", i
        pm.Add "checked", IIf(px.IsEnabled, 1&, 0&)
        pm.Add "greyscale", IIf(px.IsEnabled, 0&, 1&)

        sz = IIf((px.IsEnabled And Not px.StartedOkay), g_MakePath(App.Path) & "etc\icons\ext-error.png", px.Icon)

        pm.Add "image-file", sz
        pc.DoExCmd B_SET_ITEM, pm

    Loop

    pc.SetValue z
    pc.DoExCmd "unlock"

End Sub

Private Sub uUpdateStatus()

    If (mPage Is Nothing) Or (mCurrent Is Nothing) Then _
        Exit Sub

Dim pctb As BControl
Dim pm As MMessage
Dim pc As BControl
Dim sz As String

    Set pm = New CTempMsg

    sz = "Not enabled"

    If mCurrent.IsEnabled Then _
        sz = IIf(mCurrent.StartedOkay, "Running", "Failed to start: " & mCurrent.LastError)

    If mPage.Find("ext_stat", pc) Then _
        pc.SetText sz


    ' /* toolbar 'configure' button */

    If mPage.Find("ftb>extension", pc) Then
        prefskit_SetItem pc, 1, "enabled", -1&
        prefskit_SetItem pc, 2, "enabled", IIf(mCurrent.IsEnabled, IIf((mCurrent.Flags And SNARL_EXTN_IS_CONFIGURABLE), 1&, 0&), 0&)
        prefskit_SetItem pc, 3, "enabled", -1&

    End If

End Sub

Private Function uStr(ByVal str As String) As String

    If str = "" Then
        uStr = "<n/a>"

    Else
        uStr = str

    End If

End Function

Private Sub uDoInstallExtension()
Dim szError As String
Dim szPath As String
Dim sz As String

    sz = InputBox("Enter the name of the extension to install", "Install Extension")
    If sz = "" Then _
        Exit Sub

'    If LCase$(g_GetExtension(sz)) <> "extension" Then _
        sz = sz & ".extension"

    If Not (g_ExtnRoster Is Nothing) Then
        If g_ExtnRoster.LoadExtensionByClass(sz, szError) Then
            If g_GetUserFolderPath(szPath) Then
                uCreateMarkerFile g_MakePath(szPath) & "extensions\" & sz & ".extension"
                If Not (g_ExtnRoster Is Nothing) Then _
                    g_ExtnRoster.WriteExcludeList

            End If

            uUpdateExtensions

        Else
            MsgBox szError, vbOKOnly Or vbExclamation, "Error Installing Extension"
        
        End If

    End If

End Sub

Private Sub uCreateMarkerFile(ByVal File As String)
Dim n As Integer

    On Error Resume Next
    
    n = FreeFile()
    Open File For Output As #n
    Close #n

End Sub
