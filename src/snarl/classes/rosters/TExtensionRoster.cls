VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TExtensionRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /* TExtensionRoster.cls - Manages our loaded extensions */

Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Dim mExt() As TExtension
Dim mCount As Long
Dim mIndex As Long

Dim mRefs As Long
Dim mExcluded As CConfFile          ' // list of extensions *not* to start  (R2.31, now a CConfFile)
Dim mVersions As CConfFile          ' // R2.2: list of extensions and the installed version

Dim WithEvents thePulseTimer As BTimer
Attribute thePulseTimer.VB_VarHelpID = -1

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT

    g_Debug "TExtensionRoster.Close()", LEMON_LEVEL_PROC

    mRefs = mRefs - 1
    If mRefs > 0 Then
        ' /* still open */
        Exit Function

    End If

    ' /* run tidyup code */

Dim i As Long

    g_Debug "TExtensionRoster.Close(): tidying up..."

    Set thePulseTimer = Nothing

    If mCount Then
        g_Debug "TExtensionRoster.Close(): count=" & CStr(mCount)

        For i = mCount To 1 Step -1
            mExt(i).SetEnabled False

        Next i

        mCount = 0
        ReDim mExt(0)

    End If

    g_Debug "TExtensionRoster.Close(): done"

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

'        If (g_IsPressed(VK_LCONTROL)) Or (g_IsPressed(VK_RCONTROL)) Then
'            g_Debug "TExtensionRoster.Open(): not opening: CTRL key pressed", LEMON_LEVEL_INFO
'            Exit Function
'
'        End If

        ' /* run through our extensions content loading each one up */

        ' /*
        '
        '   Two extension locations:
        '       "%ALL_USER_APP_DATA%\fullphat\snarl\extensions"
        '       "%APP_DATA%\fullphat\snarl\extensions"
        '
        '   - All-user extensions are loaded first, then personal ones.
        '   - If a conflict occurs, all-user extension wins.
        '
        '   Extensions:
        '       - Must have a creatable class called "Extension"
        '       - Extension class must support MWndProcSink and
        '         MVersionInfo interfaces
        '       - MVersionInfo.Name provides description
        '       - Extension name (minus .dll) provides name
        '       - For R2.2, can also support MSimpleEnum to provide
        '         more information
        '
        '   R2.2: single exclude file, in user's personal %APP_DATA%
        '   contains exclusions for all extensions
        '
        ' */

Dim pFolder As storage_kit.Node
Dim i As Long

        g_AddIgnoreLock         ' // we don't want loads of 'xxx registered with Snarl' notifications...

        ' /* R2.31: configure our exclude list */

        Set mExcluded = New CConfFile
        mExcluded.SetTo gPrefs.SnarlConfigPath & "extensions\exclude"

        uGetExcludeList
        uGetVersionList         ' // R2.2: keep a record of which versions are installed

        ' /* get all-user extensions first */

        g_GetSystemFolderNode CSIDL_COMMONAPPDATA, pFolder
        uGetExtensions pFolder

        ' /* get per-user extensions next */

        g_GetSystemFolderNode CSIDL_APPDATA, pFolder
        uGetExtensions pFolder



'        ' /* V38.139 -- process /etc/startup-script */
'
'Dim sz As String
'
'        If g_GetUserFolder(pFolder) Then
'            If pFolder.SetTo(g_MakePath(pFolder.File) & "etc") Then
'                With New CConfFile
'                    If .SetTo(g_MakePath(pFolder.File) & "startup-script") Then
'                        g_Debug "TExtensionRoster.Open(): processing startup-script..."
'                        .Rewind
'                        Do While .GetEntry(sz, "")
'                            i = WinExec(sz, SW_SHOWNOACTIVATE)
'                            g_Debug "TExtensionRoster.Open(): '" & sz & "' --> " & CStr(i)
'
'                        Loop
'                    End If
'                End With
'            End If
'        End If

        g_RemIgnoreLock

        Set thePulseTimer = New BTimer
        thePulseTimer.SetTo 250

    End If

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    MRoster_CountItems = mCount

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long
Dim i As Long

    Name = LCase$(Name)

    If mCount Then
        For i = 1 To mCount
            If LCase$(mExt(i).Class) = Name Then
                MRoster_FindItem = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.MObject

    If (Index > 0) And (Index <= mCount) Then _
        Set MRoster_ItemAt = mExt(Index)

End Function

Private Function MRoster_NextItem(Item As melon.MObject) As Boolean

    If mIndex <= mCount Then
        Set Item = mExt(mIndex)
        mIndex = mIndex + 1
        MRoster_NextItem = True

    End If

End Function

Private Sub MRoster_Rewind()

    mIndex = 1

End Sub

Private Property Get MVersionInfo_Date() As String

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "extension.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

'Private Function uNotExcluded(ByVal Filename As String, ByRef Exclude As ConfigSection) As Boolean
'
'    If (Exclude Is Nothing) Then
'        uNotExcluded = True
'        Exit Function
'
'    End If
'
'    uNotExcluded = Not Exclude.Find(Filename, "")
'    g_Debug "TExtensionRoster.uNotExcluded(): '" & Filename & "' = " & uNotExcluded
'
'End Function

Private Function uGetExtension(ByRef Folder As storage_kit.Node, Optional ByVal CheckVersion As Boolean) As Boolean

    ' /* 'folder' should be the folder (or marker file) which contains the extension */

    On Error Resume Next

    If (Folder Is Nothing) Then _
        Exit Function

Dim szExtension As String
Dim fLinked As Boolean

    If Folder.IsFolder Then
        ' /* old-style extension - is actually in this folder */

        szExtension = Folder.Filename
        fLinked = False

        If MRoster_FindItem(szExtension) <> 0 Then
            g_Debug "TExtensionRoster.uGetExtension(): '" & szExtension & "' is already loaded"
            Exit Function

        End If

    Else
        ' /* new style extension - installed elsewhere */

        If LCase$(Folder.Extension) <> "extension" Then _
            Exit Function

        szExtension = Folder.Filename
        fLinked = True

        If MRoster_FindItem(g_RemoveExtension(szExtension)) <> 0 Then
            g_Debug "TExtensionRoster.uGetExtension(): '" & g_RemoveExtension(szExtension) & "' is already loaded"
            Exit Function

        End If

    End If


    ' /* R2.2 - we don't bother validating the extension's location and we no longer register them - this
    '    must be done by the extension installer */

'Dim szLocation As String
'
'    szLocation = LCase$(g_GetAXLocation(ext_name & ".Extension"))
'    If szLocation <> LCase$(szFile) Then
'        g_Debug "TExtensionRoster.uGetExtension(): '" & ext_name & "' registration mismatch", LEMON_LEVEL_WARNING
'
'        ' /* try to register the component in this location */
'
'        If Not g_RegSvr32(szFile) Then
'            g_Debug "TExtensionRoster.uGetExtension(): '" & ext_name & "' failed to register", LEMON_LEVEL_CRITICAL
'            Exit Function
'
'        End If
'    End If

Dim px As TExtension

    Set px = New TExtension

    ' /* if SetTo() fails then this isn't a bona-fide extension */

    If Not px.SetTo(g_RemoveExtension(szExtension), Folder.File) Then _
        Exit Function

    ' /* calling bInit() on the extension checks the extension is valid and retrieves
    '    information about it */

'    If Not px.Init(Folder.File) Then _
        Exit Function

    ' /* initialised okay, so add it to our list */

    uAdd px, CheckVersion

    ' /* set enabled state */

    If (Not (mExcluded Is Nothing)) Or (Not gDebugMode) Then
        px.SetEnabled Not mExcluded.Exists(g_RemoveExtension(Folder.Filename))

        ' /* we return TRUE here even if the extension is disabled or failed to initialise */

        uGetExtension = True

    Else
        g_Debug "TExtensionRoster.uGetExtension(): exclude list not loaded", LEMON_LEVEL_WARNING

    End If

End Function

'Public Sub Unload()
'
'    ' /* stop them */
'
'    g_Debug "globalExtReload(): stopping extensions..."
''    SetEnabled False
'    DoEvents
'    DoEvents
'
'    ' /* unload them */
'
'    g_Debug "globalExtReload(): unloading extensions..."
'    MLibrary_Close
'    DoEvents
'    DoEvents
'
'End Sub

'Public Sub Load()
'
'    ' /* load them */
'
'    g_Debug "globalExtReload(): loading extensions..."
'    MLibrary_Open
'    DoEvents
'    DoEvents
'
'    ' /* start them */
'
'    g_Debug "globalExtReload(): starting extensions..."
''    SetEnabled True
'    DoEvents
'    DoEvents
'
'End Sub

Public Sub Reload()

    g_AddIgnoreLock

    MLibrary_Close

    ' /* wait */

    g_Debug "TExtensionRoster.Reload(): waiting..."
    Sleep 500

    MLibrary_Open

    g_RemIgnoreLock

End Sub

Public Function Find(ByVal Name As String, ByRef Extension As TExtension) As Boolean
Dim i As Long

    i = MRoster_FindItem(Name)
    If i Then
        Set Extension = mExt(i)
        Find = True

    End If

End Function

Public Function IndexOf(ByVal Name As String) As Long

    IndexOf = MRoster_FindItem(Name)

End Function

Public Function ItemAt(ByVal Index As Long) As TExtension

    If (Index > 0) And (Index <= mCount) Then _
        Set ItemAt = mExt(Index)

End Function

Public Sub WriteExcludeList()

    ' /* R2.31: only if not blocked */

    If (gDebugMode) Or (gSysAdmin.TreatSettingsAsReadOnly) Then _
        Exit Sub

Dim i As Long

    With mExcluded
        .MakeEmpty
        .Add "[exclude]"    ' // backwards compatability

        If mCount Then
            For i = 1 To mCount
                If Not mExt(i).IsEnabled Then _
                    .Add mExt(i).Class

            Next i
        End If

        .Save

    End With

    uGetExcludeList             ' // reload it to refresh mExcluded

End Sub

Private Sub uGetExcludeList()

    If Not (mExcluded Is Nothing) Then _
        mExcluded.Reload

End Sub

Public Function LoadExtension(ByVal Name As String) As M_RESULT
Dim pn As storage_kit.Node

    ' /* 'Name' should be the full path of the folder which contains the extension to be loaded */

    If g_GetPath(Name) = "" Then
        If g_GetUserFolder(pn, False) Then
            Name = g_MakePath(pn.File) & "extensions\" & Name

        Else
            g_Debug "TExtensionRoster.LoadExtension(): couldn't get user folder", LEMON_LEVEL_CRITICAL
            LoadExtension = M_FAILED
            Exit Function

        End If
    End If

    Set pn = get_node(Name)
    If uGetExtension(pn, True) Then
        frmAbout.bUpdateExtList
        frmAbout.bUpdateAppList
        LoadExtension = M_OK

    Else
        LoadExtension = M_NOT_FOUND

    End If

End Function

Public Function UnloadExtension(ByVal Name As String) As M_RESULT
Dim i As Long

    ' /* looks for the specified extension and, if it exists, unloads it cleanly */

    i = uFind(Name)
    If i = 0 Then
        g_Debug "TExtensionRoster.UnloadExtension(): '" & Name & "' not found", LEMON_LEVEL_CRITICAL
        UnloadExtension = M_NOT_FOUND
        Exit Function

    End If

    ' /* unload it but don't update the exclude file */

    g_Debug "TExtensionRoster.UnloadExtension(): stopping '" & Name & "'..."
'    mExt(i).SetEnabled2 False, , False

    mExt(i).SetEnabled False

    ' /* remove it from our list */

    uRemove i

    frmAbout.bUpdateExtList
    frmAbout.bUpdateAppList

End Function

Private Sub uRemove(ByVal Index As Long)
Dim i As Long

    If Index < mCount Then
        For i = Index To (mCount - 1)
            Set mExt(i) = mExt(i + 1)

        Next i

    End If

    mCount = mCount - 1
    ReDim Preserve mExt(mCount)

End Sub

Private Function uFind(ByVal Name As String) As Long
Dim i As Long

    Name = LCase$(Name)

    If mCount Then
        For i = 1 To mCount
            If LCase$(mExt(i).Class) = Name Then
                uFind = i
                Exit Function

            End If

        Next i
    End If

End Function

Private Sub uGetExtensions(ByRef Folder As storage_kit.Node)
Dim pf As storage_kit.Node

    If (Folder Is Nothing) Then _
        Exit Sub

    Set pf = New storage_kit.Node
    If Not pf.SetTo(g_MakePath(Folder.File) & "full phat\snarl\extensions") Then
        g_Debug "TExtensionRoster.uGetExtensions(): path '" & g_MakePath(Folder.File) & "full phat\snarl\extensions' doesn't exist", LEMON_LEVEL_WARNING
        Exit Sub

    End If

    If Not pf.IsFolder() Then
        g_Debug "TExtensionRoster.uGetExtensions(): '" & pf.File & "' is not a folder", LEMON_LEVEL_WARNING
        Exit Sub

    End If

Dim i As Long

    With pf
        .ReadContents
        For i = 1 To .CountNodes
            uGetExtension .NodeAt(i)

        Next i

    End With

    If g_IsWinXPOrBetter() Then _
        CoFreeUnusedLibrariesEx 0, 0

End Sub

Public Sub SendSnarlState(ByVal Running As Boolean)
Dim i As Long

    ' /* calls SNARL_EXT_START or SNARL_EXT_STOP on all enabled extensions */

    If mCount Then
        For i = 1 To mCount
            mExt(i).CallProc SNARL_EXT_STATUS, IIf(Running, 1, 0), 0

        Next i
    End If

End Sub

Private Sub uAdd(ByRef Extension As TExtension, ByVal CheckVersion As Boolean)
Dim szNewVer As String
Dim sz As String

    ' /* 'CheckVersion' is set by snPrivateLoadExtension().  Here we check the current
    '    version of the extension compared to the one we're loading.  If the versions
    '    are different we show a notification telling the user that the extension
    '    was upgraded (or, possibly, downgraded) */

    szNewVer = Extension.VerString(False)
    If Not mVersions.FindEntry(LCase$(Extension.Class), sz) Then
        ' /* new */
        mVersions.Add LCase$(Extension.Class), szNewVer

        ' /* R2.31: only if not blocked */
        If Not gSysAdmin.TreatSettingsAsReadOnly Then _
            mVersions.Save

        If CheckVersion Then _
            snShowMessage Extension.Name & " Installed", _
                          "Version: " & Extension.VerString(), -1, Extension.Icon
        

    Else
        ' /* known */
        If CheckVersion Then
            If Val(sz) <> Val(szNewVer) Then
                snShowMessage Extension.Name & " Upgraded", _
                              "New version: " & Extension.VerString(), -1, Extension.Icon

            End If

        End If

        mVersions.Update LCase$(Extension.Class), sz

        ' /* R2.31: only if not blocked */
        If Not gSysAdmin.TreatSettingsAsReadOnly Then _
            mVersions.Save

    End If

Dim i As Long
Dim j As Long

    ' /* add it alpha-sorted */

    If mCount Then
        For i = 1 To mCount
            If LCase$(Extension.Name) < LCase$(mExt(i).Name) Then
                mCount = mCount + 1
                ReDim Preserve mExt(mCount)
                For j = mCount To (i + 1) Step -1
                    Set mExt(j) = mExt(j - 1)

                Next j

                Set mExt(i) = Extension
                Exit Sub

            End If
        Next i
    End If

    ' /* add it */
    mCount = mCount + 1
    ReDim Preserve mExt(mCount)
    Set mExt(mCount) = Extension

End Sub

Private Sub uGetVersionList()
Dim szPath As String

    If Not g_GetUserFolderPath(szPath) Then _
        Exit Sub

    Set mVersions = New CConfFile
    If mVersions.SetTo(g_MakePath(szPath) & "etc\.extensions") Then
        g_Debug "TExtensionRoster.uGetVersionList(): extensions version list loaded"

    Else
        g_Debug "TExtensionRoster.uGetVersionList(): extensions version list not found", LEMON_LEVEL_WARNING

    End If

End Sub

Private Sub thePulseTimer_Pulse()

    If mCount = 0 Then _
        Exit Sub

Dim i As Long

    For i = 1 To mCount
        mExt(i).Pulse

    Next i

End Sub

Public Function LoadExtensionByClass(ByVal Class As String, ByRef Result As String) As Boolean
Dim anExt As TExtension

'    MsgBox MRoster_FindItem(Class)

    If MRoster_FindItem(Class) Then
        Result = "Extension is already loaded"
        Exit Function

    End If

    Set anExt = New TExtension
    If anExt.SetTo(Class, "") Then
        ' /* initialised okay, so add it to our list */
        uAdd anExt, False
        LoadExtensionByClass = True

    Else
        Result = Class & " is not a valid extension"

    End If

End Function



