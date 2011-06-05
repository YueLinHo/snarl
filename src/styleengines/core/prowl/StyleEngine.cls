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

Private Const LIB_DATE = "27-Jul-2009"

Dim mLastErr As String
Dim mPanel As BPrefsPanel

Implements IStyleEngine
Implements KPrefsPanel

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = 1

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance

    Set IStyleEngine_CreateInstance = New StyleInstance

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = App.FileDescription

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
Dim bm As MImage

    If (mPanel Is Nothing) Then

'        mPanelStyle = StyleName

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle "Prowl Preferences"
            .SetWidth 300
            .SetWindow 1

            load_image g_MakePath(App.Path) & "general.png", bm
            .AddPage new_BPrefsPage("General", bm, New TPage)

            Set bm = Nothing
            load_image g_MakePath(App.Path) & "network.png", bm
            .AddPage new_BPrefsPage("Network", bm, New TPage)
            
            Set bm = Nothing
            load_image g_MakePath(App.Path) & "icon.png", bm
            .AddPage new_BPrefsPage("About", bm, New TPage)
            
            .Go
'            g_SetWindowIconToAppResourceIcon mPanel.hWnd

        End With
    End If

    IStyleEngine_GetConfigWindow = mPanel.hWnd

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    ' /* return code isn't really important so long as it's an error code (i.e. <> 0) */

    IStyleEngine_Initialize = M_FAILED
'
'    mLastErr = "Needs misc.resource V48d1 or better"
'    If Not uMiscResourceOk() Then _
'        Exit Function
'
'    ' /* look through %APP_DATA%\full phat\snarl\styles for any EZ styles */
'
'    mLastErr = "Couldn't get user folder"
'    If Not g_GetUserFolder(pn) Then _
'        Exit Function
'
    ' /* defaults */

    With gConfig
        .OnlyShowPriorityNotifications = False
        .UseProxyServer = False
        .ProxyServerPort = 80
        .ReplaceCRLFs = True
        .Timeout = 5

    End With

    ' /* load config here */

Dim sz As String
Dim i As Long

    With New ConfigFile
        .File = style_GetSnarlConfigPath("prowl")
        If .Load Then
            i = .FindSection("general")
            If i Then
                With .SectionAt(i)
                    If .Find("UserKey", sz) Then _
                        gConfig.UserKey = sz

                    If .Find("OnlyShowPriorityNotifications", sz) Then _
                        gConfig.OnlyShowPriorityNotifications = (Val(sz) <> 0)

                    If .Find("ReplaceCRLFs", sz) Then _
                        gConfig.ReplaceCRLFs = (Val(sz) <> 0)

                End With
            End If

            i = .FindSection("network")
            If i Then
                With .SectionAt(i)
                    If .Find("UseProxyServer", sz) Then _
                        gConfig.UseProxyServer = (Val(sz) <> 0)

                    If .Find("ProxyServer", sz) Then _
                        gConfig.ProxyServer = sz

                    If .Find("ProxyServerPort", sz) Then _
                        gConfig.ProxyServerPort = Val(sz)

                    If .Find("ProxyUsername", sz) Then _
                        gConfig.ProxyUsername = sz

                    If .Find("ProxyPassword", sz) Then _
                        gConfig.ProxyPassword = sz

                    If .Find("Timeout", sz) Then _
                        gConfig.Timeout = Val(sz)

                End With
            End If

            ' /* validate loaded settings */

            If gConfig.Timeout < 1 Then
                gConfig.Timeout = 1

            ElseIf gConfig.Timeout > 9 Then
                gConfig.Timeout = 9

            End If

        Else
            g_Debug "StyleEngine.Initialize(): couldn't load " & .File, LEMON_LEVEL_WARNING

        End If

    End With

    mLastErr = ""
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

    If (Index < 1) Or (Index > 1) Then _
        Exit Sub

    With Style
        .Copyright = App.LegalCopyright
        .Date = LIB_DATE
        .Description = App.FileDescription
        .Flags = S_STYLE_IS_CONFIGURABLE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME
        .IconPath = g_MakePath(App.Path) & "icon.png"
        .Major = 2
        .Minor = 0
        .Name = "Prowl"
        .Path = g_MakePath(App.Path)
        .Schemes = "Standard"
        .SupportEmail = "snarl@fullphat.net"
        .URL = "www.fullphat.net"

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

    g_WriteConfig
    Set mPanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub


