VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyleRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    ' /* TStyleRoster.cls - Manages our loaded style packs */

Dim mEngine() As T_SNARL_STYLE_ENGINE_INFO
Dim mEngines As Long

    ' /* full list of all supported styles and schemes */
Dim mItem() As TStyle
Dim mItems As Long
Dim mIndex As Long

Dim mRefs As Long
Dim mVersions As CConfFile          ' // R2.2: list of engines and the installed version

Implements MLibrary
Implements MRoster
Implements MVersionInfo

Private Function MLibrary_Close() As melon.M_RESULT
Dim i As Long

    mRefs = mRefs - 1
    If mRefs = 0 Then
        ' /* run tidyup code */



        ' /* unload style engines */

        If mEngines Then
            For i = 1 To mEngines
                If Not (mEngine(i).Obj Is Nothing) Then _
                    mEngine(i).Obj.Tidyup

            Next i
        End If

        mEngines = 0
        ReDim mEngine(0)

        mItems = 0
        ReDim mItem(0)

    End If

    If g_IsWinXPOrBetter() Then _
        g_FlushLibs

End Function

Private Sub MLibrary_Initialize()

    ' /* not currently used */

End Sub

Private Function MLibrary_Magic() As Long

    MLibrary_Magic = &H23232323

End Function

Private Function MLibrary_Open() As melon.M_RESULT
Dim pFolder As storage_kit.Node
Dim sz As String
Dim i As Long

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */

        uGetVersionList
        uAddStyleEngine New TInternalStyle, sz, "<internal>.styleengine", &H80000000


        If g_GetUserFolder(pFolder) Then
            If pFolder.SetTo(g_MakePath(pFolder.File) & "styles") Then
                If pFolder.IsFolder() Then
                    pFolder.ReadContents
                    For i = 1 To pFolder.CountNodes
                        If LCase$(pFolder.NodeAt(i).Extension = "styleengine") Then _
                            uLoadStyleEngine pFolder.NodeAt(i).Filename

                    Next i

                Else
                    g_Debug "TStyleRoster.Open(): '" & g_MakePath(pFolder.File) & "full phat\snarl\styles' is not a folder", LEMON_LEVEL_WARNING

                End If

            Else
                g_Debug "TStyleRoster.Open(): '" & g_MakePath(pFolder.File) & "full phat\snarl\styles' does not exist", LEMON_LEVEL_WARNING

            End If

        Else
            g_Debug "TStyleRoster.Open(): Snarl user folder does not exist", LEMON_LEVEL_WARNING

        End If
    End If

End Function

Private Function MLibrary_OpenCount() As Long

    MLibrary_OpenCount = mRefs

End Function

Private Sub MLibrary_Uninitialize()

    ' /* not currently used */

End Sub

Private Function MRoster_CountItems() As Long

    g_Debug "TStyleRoster - CALL TO MRoster->CountItems", LEMON_LEVEL_WARNING

End Function

Private Function MRoster_FindItem(ByVal Name As String) As Long

    g_Debug "TStyleRoster - CALL TO MRoster->FindItem", LEMON_LEVEL_WARNING

End Function

Private Function MRoster_ItemAt(ByVal Index As Long) As melon.MObject

    g_Debug "TStyleRoster - CALL TO MRoster->ItemAt", LEMON_LEVEL_WARNING

End Function

Private Function MRoster_NextItem(Item As melon.MObject) As Boolean

    g_Debug "TStyleRoster - CALL TO MRoster->NextItem", LEMON_LEVEL_WARNING

End Function

Private Sub MRoster_Rewind()

    g_Debug "TStyleRoster - CALL TO MRoster->Rewind", LEMON_LEVEL_WARNING

End Sub

Private Property Get MVersionInfo_Date() As String

    MVersionInfo_Date = "18-Nov-2008"

End Property

Private Property Get MVersionInfo_Name() As String

    MVersionInfo_Name = "style.roster"

End Property

Private Property Get MVersionInfo_Revision() As Long

    MVersionInfo_Revision = App.Revision

End Property

Private Property Get MVersionInfo_Version() As Long

    MVersionInfo_Version = App.Major

End Property

Private Function uLoadStyleEngine(ByVal Object As String, Optional ByRef FailedReason As String) As Boolean
Dim pStyleEngine As IStyleEngine

    On Error Resume Next

    Err.Clear
    Set pStyleEngine = CreateObject(Object)
    If Err.Number = 0 Then
        uLoadStyleEngine = uAddStyleEngine(pStyleEngine, FailedReason, Object)
        If Not uLoadStyleEngine Then _
            g_Debug "TStyleRoster.uLoadStyleEngine(): '" & Object & "' failed to load: " & FailedReason, LEMON_LEVEL_CRITICAL

    Else
        FailedReason = "Couldn't create '" & Object & "'"
        g_Debug "TStyleRoster.uLoadStyleEngine(): " & Err.Description & "' creating '" & Object & "'", LEMON_LEVEL_CRITICAL

    End If

End Function

Public Function uAddStyleEngine(ByRef Engine As IStyleEngine, ByRef FailedReason As String, ByVal Library As String, Optional ByVal Flags As Long) As Boolean

    ' /* introduced in R2.2 - much neater style engine interface - returns False (and tries to set 'FailedReason' if
    '    something goes wrong), otherwise returns True (and adds engine to roster) */

    On Error Resume Next

    FailedReason = ""

    If (Engine Is Nothing) Then
        FailedReason = "Error creating engine object"
        Exit Function

    End If

    ' /* initialize it */

Dim hr As M_RESULT

    ' /* from R2.2 style can provide a human-readable reason for why it didn't initialize */

    hr = Engine.Initialize()
    If hr <> M_OK Then
        g_Debug "TStyleRoster.uAddStyleEngine(): style engine didn't initialize (" & g_HexStr(hr) & ")", LEMON_LEVEL_CRITICAL
        FailedReason = Engine.LastError()
        Exit Function

    End If

Dim c As Long

    ' /* must have at least one style... */

    c = Engine.CountStyles()
    If c < 1 Then
        g_Debug "uAddStyleEngine(): style engine '" & Engine.Name & "' (" & Library & ")" & "' has no styles", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* add it to the engine list */

    uCheckVersion Engine, False

    mEngines = mEngines + 1
    ReDim Preserve mEngine(mEngines)
    With mEngine(mEngines)
        Set .Obj = Engine
        .Name = IIf(Engine.Name = "", Library, Engine.Name)
        .Version = Engine.Version
        .Revision = Engine.Revision
        .Date = Engine.Date
        .Description = Engine.Description
        .Path = Engine.Path
        .Flags = Flags

    End With

    uAddStyles Engine
    uAddStyleEngine = True

End Function

Private Sub uAddStyles(ByRef Engine As IStyleEngine)

    ' /* get the style names - **assumes engine has been intialised and has at least one style** */

Dim pStyleInfo As style_info
Dim pStyle As TStyle
Dim i As Long

    For i = 1 To Engine.CountStyles
        Engine.StyleAt i, pStyleInfo
        If (pStyleInfo.Name <> "") And (IndexOf(pStyleInfo.Name) = 0) Then

            ' /* configure the TStyle object */
            Set pStyle = New TStyle
            pStyle.bInit Engine, pStyleInfo

            ' /* add it to the roster */
            uAddStyleToRoster pStyle

        Else
            g_Debug "uAddStyles(): style '" & pStyle.Name & "' is either invalid or already exists", LEMON_LEVEL_WARNING

        End If

    Next i

End Sub

Private Sub uAddStyleToRoster(ByRef Style As TStyle)
Dim i As Long
Dim j As Long

    ' /* add it alpha-sorted */

    If mItems Then
        For i = 1 To mItems
            If LCase$(Style.Name) < LCase$(mItem(i).Name) Then
                mItems = mItems + 1
                ReDim Preserve mItem(mItems)
                For j = mItems To (i + 1) Step -1
                    Set mItem(j) = mItem(j - 1)

                Next j

                Set mItem(i) = Style
                Exit Sub

            End If
        Next i
    End If

    mItems = mItems + 1
    ReDim Preserve mItem(mItems)
    Set mItem(mItems) = Style

    g_Debug "TStyleRoster.uAddStyleToRoster(): added style '" & Style.Name & "'"

End Sub

Public Sub Rewind()

    mIndex = 1

End Sub

Friend Function GetNextStyle(ByRef Style As TStyle) As Boolean

    If (mIndex < 1) Or (mIndex > mItems) Then _
        Exit Function

    Set Style = mItem(mIndex)
    mIndex = mIndex + 1
    GetNextStyle = True

End Function

Public Function CountStyles() As Long

    CountStyles = mItems

End Function

Friend Function StyleAt(ByVal Index As Long) As TStyle

    If (Index > 0) And (Index <= mItems) Then _
        Set StyleAt = mItem(Index)

End Function

Public Function GetStyleList(Optional ByRef Result As Boolean, Optional ByVal MarkDefault As Boolean = True) As String
Dim i As Long

    MarkDefault = False

    If mItems Then
        For i = 1 To mItems
            GetStyleList = GetStyleList & mItem(i).Name & IIf((MarkDefault) And (LCase$(mItem(i).Name) = uDefaultStyle), " (Default)", "") & "|"

        Next i

        GetStyleList = g_SafeLeftStr(GetStyleList, Len(GetStyleList) - 1)

    End If

    If GetStyleList = "" Then _
        GetStyleList = "<invalid list>"

    Result = (GetStyleList <> "<invalid list>")

End Function

Public Function GetNonWindowStyleList(Optional ByRef Result As Boolean) As String
Dim pt As TStyle
Dim i As Long
Dim j As Long

    If mItems Then
        For i = 1 To mItems
            Set pt = mItem(i)
            If (pt.Flags And S_STYLE_IS_WINDOWLESS) Then
                If pt.CountSchemes Then
                    For j = 1 To pt.CountSchemes
                        GetNonWindowStyleList = GetNonWindowStyleList & pt.Name & ": " & pt.SchemeAt(j) & "|"

                    Next j
                End If
            End If
        Next i

        GetNonWindowStyleList = g_SafeLeftStr(GetNonWindowStyleList, Len(GetNonWindowStyleList) - 1)

    End If

    If GetNonWindowStyleList = "" Then _
        GetNonWindowStyleList = "<None>"

    Result = (GetNonWindowStyleList <> "<None>")

End Function

Public Function IndexOf(ByVal Style As String) As Long
Dim i As Long

    ' /* expects just a style name - not a style/scheme pair! */

    If mItems = 0 Then _
        Exit Function

    Style = LCase$(Style)
    For i = 1 To mItems
        If LCase$(mItem(i).Name) = Style Then
            IndexOf = i
            Exit Function

        End If
    Next i

End Function

Public Function CountEngines() As Long

    CountEngines = mEngines

End Function

Friend Function EngineAt(ByVal Index As Long, ByRef EngineInfo As T_SNARL_STYLE_ENGINE_INFO) As Boolean

    If (Index > 0) And (Index <= mEngines) Then
        LSet EngineInfo = mEngine(Index)
        EngineAt = True

    End If

End Function

Public Function Find(ByVal Name As String, ByRef Style As TStyle) As Boolean
Dim i As Long

    i = Me.IndexOf(Name)
    If i Then
        Set Style = mItem(i)
        Find = True

    End If

End Function






'Friend Function RefreshStyle(ByVal Name As String, ByRef Style As T_SNARL_STYLE_2) As Boolean
'Dim i As Long
'
'    ' /* do we have the style? if not, can't do anything... */
'
'    i = Me.IndexOf(Name)
'    If i = 0 Then _
'        Exit Function
'
'Dim ps As SNARL_STYLE
'
'    ' /* got the style so call ISSEngine.GetStyle() but with the style's name
'    '    in SNARL_STYLE->Name.  This tells the engine that we're requesting
'    '    an update to a specific style; not just enumerating the list */
'
'    ps.Name = Name
'    If mItem(i).Engine.GetStyle(ps) Then
'        LSet mItem(i).Info = ps
'        uGetSchemes ps.Schemes, mItem(i)        ' // convert the schemes message
'        LSet Style = mItem(i)
'        RefreshStyle = True
'
'    End If
'
'End Function

Public Function Load(ByVal Engine As String, ByRef FailText As String, ByVal NotifyVersionChange As Boolean) As Boolean
Dim i As Long

    On Error Resume Next

    i = uFindEngine(Engine)
    If i = 0 Then
        ' /* not already in list */
        If Not uLoadStyleEngine(Engine, FailText) Then _
            Exit Function

    Else
        ' /* know about it but is it already loaded? */

        If Not (mEngine(i).Obj Is Nothing) Then
            g_Debug "TStyleRoster.Unload(): '" & Engine & "' is already loaded", LEMON_LEVEL_CRITICAL
            FailText = "Style Engine is already loaded"
            Exit Function

        End If

    End If

Dim pEngine As IStyleEngine

    Err.Clear
    Set pEngine = CreateObject(Engine)
    If (pEngine Is Nothing) Or (Err.Number <> 0) Then
        g_Debug "TStyleRoster.Load(): failed to create engine instance", LEMON_LEVEL_CRITICAL
        FailText = "Couldn't create instance"
        Exit Function

    End If

    ' /* initialize it */

Dim hr As M_RESULT

    ' /* from R2.2, style can provide a human-readable reason for why it didn't initialize */

    hr = pEngine.Initialize()
    If hr <> M_OK Then
        g_Debug "TStyleRoster.Load(): style engine didn't initialize (" & g_HexStr(hr) & ")", LEMON_LEVEL_CRITICAL
        FailText = pEngine.LastError()
        Exit Function

    End If

Dim c As Long

    ' /* must have at least one style... */

    If pEngine.CountStyles() < 1 Then
        g_Debug "TStyleRoster.Load(): style engine '" & pEngine.Name & "' has no styles", LEMON_LEVEL_CRITICAL
        FailText = "Engine doesn't provide any styles"
        Exit Function

    End If

    ' /* ok - let's set the object and get its styles */

    uCheckVersion pEngine, NotifyVersionChange

    With mEngine(i)
        Set .Obj = pEngine
        .Date = pEngine.Date
        .Description = pEngine.Description
        .Flags = 0                          ' // not used yet except to identify built-in engine
        .Path = pEngine.Path
        .Revision = pEngine.Revision
        .Version = pEngine.Version

    End With

    uAddStyles pEngine
    frmAbout.bUpdateStylesList

    Load = True

End Function

Public Function StylesForEngine(ByVal Engine As String, ByRef Styles As String) As Boolean
Dim i As Long

    i = uFindEngine(Engine)
    If i = 0 Then
        g_Debug "TStyleRoster.StylesForEngine(): '" & Engine & "' not in list", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If (mEngine(i).Obj Is Nothing) Then
        g_Debug "TStyleRoster.StylesForEngine(): '" & Engine & "' is already unloaded", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim sz As String
Dim j As Long

    ' /* find all styles associated with this engine */

    If mItems Then
        For j = 1 To mItems
            If mItem(j).Engine.Name = Engine Then
                sz = sz & mItem(j).Name & vbCrLf

            End If
        Next j
    End If

    Styles = sz
    StylesForEngine = True

End Function

Public Function Unload(ByVal Engine As String, ByRef StylesRemoved As String) As Boolean
Dim i As Long

    i = uFindEngine(Engine)
    If i = 0 Then
        g_Debug "TStyleRoster.Unload(): '" & Engine & "' not in list", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If (mEngine(i).Obj Is Nothing) Then
        g_Debug "TStyleRoster.Unload(): '" & Engine & "' is already unloaded", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* stop the style engine */

    mEngine(i).Obj.Tidyup

Dim sz As String
Dim j As Long

    ' /* find and remove all styles associated with this engine */

    If mItems Then
        For j = mItems To 1 Step -1
            If mItem(j).Engine.Name = Engine Then
                sz = sz & mItem(j).Name & vbCrLf
                uRemoveStyle j

            End If
        Next j
    End If

    ' /* zap and unload the engine code itself */

    Set mEngine(i).Obj = Nothing
    StylesRemoved = sz

    If g_IsWinXPOrBetter() Then _
        CoFreeUnusedLibrariesEx 0, 0

    frmAbout.bUpdateStylesList
    Unload = True

End Function

Private Function uFindEngine(ByVal Engine As String) As Long
Dim i As Long

    If mEngines Then
        For i = 1 To mEngines
            If mEngine(i).Name = Engine Then
                uFindEngine = i
                Exit Function

            End If
        Next i
    End If

End Function

Private Sub uRemoveStyle(ByVal Index As Long)

    If (Index < 1) Or (Index > mItems) Then _
        Exit Sub

    ' /* trash the pointer to the style engine */

    mItem(Index).bZapEngine

Dim i As Long

    If Index < mItems Then
        For i = Index To (mItems - 1)
            Set mItem(i) = mItem(i + 1)

        Next i

    End If

    mItems = mItems - 1
    ReDim Preserve mItem(mItems)

End Sub

Private Sub uGetVersionList()
Dim szPath As String

    If Not g_GetUserFolderPath(szPath) Then _
        Exit Sub

    Set mVersions = New CConfFile
    If mVersions.SetTo(g_MakePath(szPath) & "etc\.styleengines") Then
        g_Debug "TStyleRoster.uGetVersionList(): version list loaded"

    Else
        g_Debug "TStyleRoster.uGetVersionList(): version list not found", LEMON_LEVEL_WARNING

    End If

End Sub

Private Sub uCheckVersion(ByRef Engine As IStyleEngine, ByVal Notify As Boolean)

    If (Engine Is Nothing) Or (mVersions Is Nothing) Then _
        Exit Sub

Dim szVerStr As String
Dim bChanged As Boolean
Dim sz As String

    ' /* 'CheckVersion' is set by snPrivateLoadExtension().  Here we check the current
    '    version of the extension compared to the one we're loading.  If the versions
    '    are different we show a notification telling the user that the extension
    '    was upgraded (or, possibly, downgraded) */

    szVerStr = CStr(Engine.Version) & "." & CStr(Engine.Revision)

    If Not mVersions.FindEntry(Engine.Name, sz) Then
        ' /* new */
        mVersions.Add Engine.Name, szVerStr
        mVersions.Save

        If Notify Then _
            snShowMessage "Style Engine Installed", _
                          "Name: " & Engine.Name & vbCrLf & _
                          "Version: " & szVerStr, 20, g_MakePath(App.Path) & "etc\icons\style_engine.png"

    Else
        ' /* known */

'Dim nMaj As Long
'Dim nMin As Long
'Dim i As Long
'
'        ' /* decode version string */
'
'        i = InStr(sz, ".")
'        If i > 0 Then
'            nMaj = Val(g_SafeLeftStr(sz, i - 1))
'            nMin = Val(g_SafeRightStr(sz, Len(sz) - i))
'
'            Debug.Print nMaj & " > " & nMin

        If szVerStr <> sz Then
            Debug.Print "changed"

            If Notify Then
    
'                If (Engine.Version > nMaj) Or ((Engine.Version = nMaj) And (Engine.Revision > nMin)) Then
'                End If

                snShowMessage "Style Engine Updated", _
                              "Name: " & Engine.Name & vbCrLf & _
                              "Previous version: " & sz & vbCrLf & _
                              "New version: " & szVerStr, 20, g_MakePath(App.Path) & "etc\icons\style_engine.png"

            End If

            mVersions.Update Engine.Name, szVerStr
            mVersions.Save

        End If

    End If

End Sub

Private Function uDefaultStyle() As String

    uDefaultStyle = LCase$(style_GetStyleName(g_ConfigGet("default_style")))

End Function

Private Function uDefaultScheme() As String

    uDefaultScheme = LCase$(style_GetSchemeName(g_ConfigGet("default_style")))

End Function
