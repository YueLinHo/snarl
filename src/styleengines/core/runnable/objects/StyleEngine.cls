VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           StyleEngine.cls
    '/
    '/  Description:    Provides the Runnable style engine ingress point
    '/
    '/  © 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Private Const LIB_DATE = "31-Oct-2011"
Dim mStyles As BTagList
Dim mLastErr As String
Dim mTimeout As Long

Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mStyles.CountItems

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim pri As TRunnableInstance
Dim i As Long

    ' /* look it up (translates friendly name to signature) */

    i = mStyles.IndexOf(StyleName)
    If i Then
        Set pri = New TRunnableInstance
        pri.SetStyleX mStyles.TagAt(i), mTimeout
        Set IStyleEngine_CreateInstance = pri

    Else
        g_Debug "IStyleEngine.CreateInstance(): '" & StyleName & "' not found", LEMON_LEVEL_CRITICAL

    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Manages styles provided by standalone runnable processes"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
Dim ps As TRunnableStyle
Dim i As Long

    If StyleName = ".global" Then
        MsgBox "This style engine will support a user interface in a future release.  " & _
               "For now, configurable features can be modified directly via the settings " & _
               "file (located at '" & uGetSettingsFilePath() & "').", _
               vbInformation Or vbOKOnly, "Runnable Style Engine"

    Else
        i = mStyles.IndexOf(StyleName, False)
        If i > 0 Then
            Set ps = mStyles.TagAt(i)
            ShellExecute 0, "open", ps.Executable, "-configure", vbNullString, 5

        End If

    End If

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    l3OpenLog "%APPDATA%\full phat\snarl\" & "runnable.log"

Dim pcs As CConfSection
Dim pcf As CConfFile3
Dim sz As String

    sz = style_GetSnarlStylesPath(False)
    If sz <> "" Then
        sz = g_MakePath(sz) & "runnable\runnable.conf"
        g_Debug "Initialise(): looking for '" & sz & "'..."
        Set pcf = New CConfFile3
        pcf.SetFile sz
        If pcf.Load() Then
            g_Debug "Initialise(): loaded config"
            If pcf.Find("general", pcs) Then
                g_Debug "Initialise(): [general] section found"
                sz = pcs.GetValueWithDefault("timeout", "500")
                mTimeout = g_SafeLong(sz)

            Else
                g_Debug "Initialise(): 'timeout' entry not found", LEMON_LEVEL_WARNING

            End If
        
        Else
            g_Debug "Initialise(): config not found", LEMON_LEVEL_WARNING

        End If
    End If

    If mTimeout < 500 Then _
        mTimeout = 500

    g_Debug "Initialise(): using timeout of " & CStr(mTimeout) & "ms"

    IStyleEngine_Initialize = uGetStyles()

End Function

Private Function IStyleEngine_LastError() As String

    IStyleEngine_LastError = mLastErr

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = "runnable.styleengine"

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Revision

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)
Dim pr As TRunnableStyle

    If Index = 0 Then
        ' /* provide info about ourselves */
        With Style
            .Flags = S_STYLE_IS_CONFIGURABLE

        End With

    Else
        Set pr = mStyles.TagAt(Index)
        pr.Populate Style

    End If

End Sub

Private Sub IStyleEngine_TidyUp()

    Set mStyles = Nothing
    l3CloseLog

End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Function uGetStyles() As M_RESULT
Dim pr As TRunnableStyle
Dim sz As String

    Set mStyles = new_BTagList()
    mLastErr = ""

    ' /* load styles from All Users first */

    sz = style_GetSnarlStylesPath(True)
    If sz <> "" Then
        With New CFolderContent2
            If .SetTo(sz & "\runnable") Then
                .Rewind
                Do While .GetNextFile(sz)
                    If g_IsFolder(sz) Then
                        Set pr = New TRunnableStyle
                        If pr.SetTo(sz) Then _
                            mStyles.Add pr

                    End If
                Loop
            Else
                g_Debug "uGetStyles(): couldn't set to all user runnable path", LEMON_LEVEL_WARNING

            End If

        End With
    Else
        g_Debug "uGetStyles(): failed to get all user styles path", LEMON_LEVEL_WARNING

    End If

    ' /* load styles from user folder last */

    sz = style_GetSnarlStylesPath()
    If sz <> "" Then
        With New CFolderContent2
            If .SetTo(sz & "\runnable") Then
                .Rewind
                Do While .GetNextFile(sz)
                    If g_IsFolder(sz) Then
                        Set pr = New TRunnableStyle
                        If (pr.SetTo(sz)) And (Not uStyleExists(g_FilenameFromPath(sz))) Then _
                            mStyles.Add pr

                    End If
                Loop
            Else
                g_Debug "uGetStyles(): couldn't set to user runnable path", LEMON_LEVEL_WARNING

            End If

        End With
    Else
        g_Debug "uGetStyles(): failed to get user styles path", LEMON_LEVEL_WARNING

    End If

    If mStyles.CountItems = 0 Then _
        g_Debug "uGetStyles(): no styles found", LEMON_LEVEL_WARNING

    uGetStyles = M_OK

End Function

Private Function uStyleExists(ByVal ShortName As String) As Boolean
Dim ps As TRunnableStyle

    With mStyles
        .Rewind
        Do While .GetNextTag(ps) = B_OK
            If ps.ShortName = ShortName Then
                g_Debug "uStyleExists(): can't add '" & ShortName & "': a style with this name already exists", LEMON_LEVEL_CRITICAL
                uStyleExists = True
                Exit Do

            End If

        Loop
    End With

End Function

Private Function uGetSettingsFilePath() As String
Dim sz As String

    sz = style_GetSnarlStylesPath(False)
    If sz <> "" Then _
        uGetSettingsFilePath = g_MakePath(sz) & "runnable\runnable.conf"

End Function
