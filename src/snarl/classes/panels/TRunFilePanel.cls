VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRunFilePanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRunFilePanel.cls
    '/
    '/  Description:    Provides the Run File style configuration panel
    '/
    '/  Created:        13-Jan-2011
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Dim mPanel As BPrefsPanel
Dim mSelIndex As Long

Dim WithEvents mAddEditPanel As TRunFileEditPanel

Dim mIsRunFile As Boolean
Dim mConfFile As String

Implements KPrefsPanel
Implements KPrefsPage

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)
Dim pc As BControl
Dim sz As String
Dim i As Long

    Select Case Control.GetName()
    Case "list"
        mSelIndex = Val(Value)
'        Select Case gRunFiles.TagAt(mSelIndex).Value
'        Case "1"
'            sz = "Snarl will pass specific notification elements as individual arguments, as follows: " & vbCrLf & vbCrLf & _
'                 gRunFiles.TagAt(mSelIndex).Name & " ""<notification title>"" ""<notification text>"" ""<sender>""" & vbCrLf & vbCrLf & _
'                 "Any elements prefixed with either ""data-"" or ""value-"" will also be included as additional arguments."
'
'        Case "2"
'            sz = "variable template"
'
'        End Select

        prefskit_SafeSetText mPanel, "file_hint", sz


        ' /* toolbar */
'        prefskit_SafeEnable mPanel, "toolbar_options", (Value <> "0")

'        If mPanel.Find("toolbar_options", pc) Then
''            prefskit_SetItem pc, 1, "enabled", IIf(mSelected.DefaultAck <> "", 1&, 0&)
'            prefskit_SetItem pc, 2, "enabled", IIf(mSelIndex <> 0, 1&, 0&)
'
'        End If

    Case "toolbar_options"

        Select Case Val(Value)
        Case 1
            ' /* add */
            uAddNewEntry

        Case 2
            If mSelIndex <> 0 Then _
                uRemEntry mSelIndex

        End Select

    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
Dim prf As TRunFileScheme

    Select Case Control.GetName
    Case "list"
        Set prf = gRunFiles.TagAt(Val(Control.GetValue))
        Set mAddEditPanel = New TRunFileEditPanel
        mAddEditPanel.Go mPanel.hWnd, prf

'        ShellExecute Control.Page.Panel.hWnd, "open", prf.File, vbNullString, vbNullString, 5

    Case "go_user_guide"
        ShellExecute Control.Page.Panel.hWnd, "open", g_MakePath(App.Path) & "etc\RunFileUserGuide\index.html", vbNullString, vbNullString, 5

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hWnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Function Go(ByVal IsRunFile As Boolean) As Long
Dim pPage As BPrefsPage
Dim pc As BControl
Dim pm As CTempMsg

    mIsRunFile = IsRunFile

    If mIsRunFile Then
        mConfFile = "runfile"

    Else
        mConfFile = "image"

    End If


    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "Run File Settings"
            .SetWidth 400
            .SetWindow 1

            Set pPage = new_BPrefsPage("", , Me)

            With pPage
                .SetMargin 0

                Set pm = New CTempMsg
'                pm.Add "item-height", 36&
                Set pc = new_BPrefsControl("listbox", "list", , , , pm)
                pc.SizeTo 0, 160
                .Add pc

'                .Add new_BPrefsControl("fancyedit", "file_to_run")
                .Add new_BPrefsControl("fancytoolbar", "toolbar_options", "Add...|Remove")
''                .Add new_BPrefsControl("fancybutton2", "fb>ack", "Invoke Callback", , , , False)

'                If mIsRunFile Then
'                    Set pc = new_BPrefsControl("label", "file_hint")
'                    pc.SizeTo 0, 90
'                    .Add pc
'
''                    .Add new_BPrefsControl("label", "", "Each scheme represents an individual file to run")
'''                    .Add new_BPrefsControl("label", "", "Snarl will pass notification elements as arguments to the file it runs, as follows: " & vbCrLf & _
'''                                                        "<file> ""<notification title>"" ""<notification text>"" ""<sender>"".")
'''
'''                    .Add new_BPrefsControl("label", "", "Any notification elements prefixed with either ""data-"" or ""value-"" are also included as an additional arguments.")
'
'                    .Add new_BPrefsSeparator
'                    .Add new_BPrefsControl("fancybutton2", "go_user_guide", "User Guide")
'
'                End If

            End With

            .AddPage pPage
            .Go

            uRefreshList

        End With

    End If

    Go = mPanel.hWnd

End Function

Private Sub uAddNewEntry()
Dim prf As TRunFileScheme

    Set prf = New TRunFileScheme
    If prf.CreateNew Then
        Set mAddEditPanel = New TRunFileEditPanel
        mAddEditPanel.Go mPanel.hWnd, prf
        g_RunFileLoadSchemes
        uRefreshList

    End If

End Sub

Private Sub uRemEntry(ByVal Index As Long)
Dim prf As TRunFileScheme

    Set prf = gRunFiles.TagAt(Index)
    DeleteFile prf.File
    gRunFiles.Remove Index
    uRefreshList

End Sub

Private Sub uRefreshList()
Dim pc As BControl

    If Not mPanel.Find("list", pc) Then _
        Exit Sub

Dim sz As String
Dim n As Long

Dim pt As BTagItem

    n = Val(pc.GetValue())

    With gRunFiles
        .Rewind
        Do While .GetNextTag(pt) = B_OK
            sz = sz & pt.Name & "|"

        Loop

    End With

    pc.SetText g_SafeLeftStr(sz, Len(sz) - 1)
    pc.SetValue CStr(MAX(n, 1))

End Sub

Private Sub mAddEditPanel_Done()

    uRefreshList
    
End Sub
