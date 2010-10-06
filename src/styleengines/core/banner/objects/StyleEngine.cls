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

Private Const LIB_DATE = "20-Sep-2010"

Dim mLastErr As String
Dim mPanel As BPrefsPanel

Implements IStyleEngine
Implements KPrefsPanel

'Private Function ISSEngine_ConfigRequested(ByVal StyleAndScheme As String) As melon.M_RESULT
'
'    MsgBox "Config: " & StyleAndScheme
'
''    If (mPanel Is Nothing) Then
''        Set mPanel = New BPrefsPanel
''        With mPanel
''
''            .SetHandler Me
''            .SetWidth 300
''            .AddPage new_BPrefsPage("General", , New TGeneralPage)
''
''
'''    load_image g_MakePath(App.Path) & "etc\icons\general.png", pbm
'''    mPanel.AddPage new_BPrefsPage("General", pbm, New TGeneralPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\apps.png", pbm
'''    Set mAppsPage = New TAppsPage
'''    mPanel.AddPage new_BPrefsPage("Apps", pbm, mAppsPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\display.png", pbm
'''    mPanel.AddPage new_BPrefsPage("Display", pbm, New TDisplayPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\styles.png", pbm
'''    mPanel.AddPage new_BPrefsPage("Styles", pbm, New TStylesPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\extensions.png", pbm
'''    mPanel.AddPage new_BPrefsPage("Extensions", pbm, New TExtPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\network.png", pbm
'''    mPanel.AddPage new_BPrefsPage("Network", pbm, New TNetworkPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\advanced.png", pbm
'''    mPanel.AddPage new_BPrefsPage("Advanced", pbm, New TAdvancedPage)
'''
'''    load_image g_MakePath(App.Path) & "etc\icons\about.png", pbm
'''    mPanel.AddPage new_BPrefsPage("About", pbm, New TAboutPage)
'''
'''    If gDebugMode Then
'''        load_image g_MakePath(App.Path) & "etc\icons\debug.png", pbm
'''        mPanel.AddPage new_BPrefsPage("Debug", pbm, New TDebugPage)
'''
'''    End If
''
'''    mPanel.SetTitle "Snarl Preferences"
'''    mPanel.SetWidth 500
''
''            .Go
'''    g_SetWindowIconToAppResourceIcon mPanel.hWnd
''
''        End With
''
''    End If
''
''    ISSEngine_ConfigRequested = mPanel.hWnd
'

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = 3

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance

    Select Case LCase$(StyleName)
    Case "music video"
        Set IStyleEngine_CreateInstance = New MTVInstance

    Case "info bar"
        Set IStyleEngine_CreateInstance = New InfoBarInstance

    Case "tiny"
        Set IStyleEngine_CreateInstance = New TinyInstance

    End Select

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = "Provides a number of single-instance styles which appear as banners"

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    Set gSettings = New CConfFile
    gSettings.SetTo sn41GetConfigPath() & "etc\banner.settings"

    IStyleEngine_Initialize = M_OK

End Function

Private Function IStyleEngine_LastError() As String

    IStyleEngine_LastError = mLastErr

End Function

Private Function IStyleEngine_Name() As String

    IStyleEngine_Name = App.Title

End Function

Private Function IStyleEngine_Path() As String

    IStyleEngine_Path = App.Path

End Function

Private Function IStyleEngine_Revision() As Long

    IStyleEngine_Revision = App.Revision

End Function

Private Sub IStyleEngine_StyleAt(ByVal Index As Long, Style As libSnarlStyles.style_info)

    With Style
        ' /* generic */

        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Path = App.Path
        .SupportEmail = "support@fullphat.net"
        .URL = "www.fullphat.net"

        Select Case Index
        Case 1
            .Name = "Music Video"
            .Description = "Displays the notification as an Music Video style banner along the bottom of the screen."
            .Major = 0
            .Minor = 2
            .Flags = S_STYLE_NO_DROPSHADOW Or S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_CLICK_THROUGH
            .Schemes = SCHEME_1 & "|" & SCHEME_2

        Case 2
            .Name = "Tiny"
            .Description = "Displays the notfication in a tiny bubble in the top right corner of the screen."
            .Major = 0
            .Minor = 2
            .Flags = S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE
            .Schemes = "Small|Extra Small"

        Case 3
            .Name = "Info Bar"
            .Description = "Displays the notfication in a thin bar along the top of the screen."
            .Major = 0
            .Minor = 1
            .Flags = S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE
            .Schemes = "Amiga|OS X|geOShell|Cloud:9ine|Amiga Guru"

        End Select

        .IconPath = g_MakePath(App.Path) & .Name & ".png"

    End With

End Sub

Private Sub IStyleEngine_TidyUp()
End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)

End Sub

Private Sub KPrefsPanel_Quit()

    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()

End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)

End Sub

