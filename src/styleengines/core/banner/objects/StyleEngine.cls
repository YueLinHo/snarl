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

    If StyleName <> "Tiny" Then _
        Exit Function

Dim pp As BPrefsPage
Dim sz As String

    If (mPanel Is Nothing) Then

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle StyleName & " Preferences"
            .SetWidth 300
            .SetWindow 1

            Set pp = new_BPrefsPage("", Nothing, New TTinyPage)

            With pp

                .Add new_BPrefsControl("banner", "", "Appearance")
                .Add new_BPrefsControl("fancycycle", "tiny.size", "Small|Medium|Large", "Size:", gSettings.GetValueWithDefault("tiny.size", "1"))
                
                sz = uCreatePackedFontStr("Tahoma", 7)
                .Add new_BPrefsControl("fancyfontpicker", "tiny.font", "", "Font:", gSettings.GetValueWithDefault("tiny.font", sz))

                .Add new_BPrefsControl("fancytoggle2", "tiny.cleartype", "Use ClearType™ font aliasing", "", gSettings.GetValueWithDefault("tiny.cleartype", "0"))

                .Add new_BPrefsControl("fancycolourpicker", "tiny.background", "", "Background:", gSettings.GetValueWithDefault("tiny.background", CStr(rgba(0, 0, 0))))
                .Add new_BPrefsControl("fancycolourpicker", "tiny.foreground", "", "Text:", gSettings.GetValueWithDefault("tiny.foreground", CStr(rgba(255, 255, 255))))


                ' /* test and reset buttons - all */
        
                .Add new_BPrefsControl("separator", "", "")
                .Add new_BPrefsControl("fancybutton2", "test", "Test")
                .Add new_BPrefsControl("separator", "", "")
                .Add new_BPrefsControl("fancybutton2", "reset", "Reset to Factory Settings")


'        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_FONT) Then
'            .Add new_BPrefsControl("fancyfontpicker", "text", "", "Text font:", uCreatePackedFontStr(mCurrent.ReadSetting("text_font"), mCurrent.ReadSetting("text_size")))
'
'        End If
'
'        ' /* icon */
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_HIDE_SHOW_ICON) Then
'            .Add new_BPrefsControl("fancytoggle2", "show_icon", "Show icon?", "Icon:", mCurrent.ReadSetting("show_icon"))
'
'            If (mCurrent.Flags And E_METER_STYLE_CAN_SET_ICON_POS) Then
'                Set pm = New CTempMsg
'                pm.Add "text", "Position:"
'                pm.Add "align", 1&
'                .Add new_BPrefsControl("fancytoggle", "icon_align", "Left|Right", " ", mCurrent.ReadSetting("icon_align"), pm)
'
'            End If
'        End If
'
'        ' /* background colour picker */
'
'Dim f As Boolean
'
'        f = ((mCurrent.Name = STYLE_NAME_JUST_BLACK) And (g_IsPressed(VK_LCONTROL)))
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_BKGND) Or (f) Then _
'            .Add new_BPrefsControl("fancycolourpicker", "background", "", IIf(f, "Not just black:", "Background:"), mCurrent.ReadSetting("background"))
'
'        ' /* foreground colour picker */
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_FGND) Then _
'            .Add new_BPrefsControl("fancycolourpicker", "foreground", "", "Text:", mCurrent.ReadSetting("foreground"))
'
'        ' /* border colour picker */
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_BORDER) Then _
'            .Add new_BPrefsControl("fancycolourpicker", "border", "", "Border:", mCurrent.ReadSetting("border"))
'
'        ' /* foreground opacity */
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_SET_FG_ALPHA) Then
'            Set pm = New CTempMsg
'            pm.Add "min", 0&
'            pm.Add "max", 100&
'            pm.Add "freq", 10&
'            pm.Add "suffix", "%"
'            .Add new_BPrefsControl("fancyslider", "fg_alpha", "", "Foreground opacity:", mCurrent.ReadSetting("fg_alpha"), pm)
'
'        End If
'
'        ' /* meter */
'
'        .Add new_BPrefsControl("banner", "", "Meter")
'
'        .Add new_BPrefsControl("fancytoggle2", "show_graphically", "Auto-detect numeric-only text?", "", mCurrent.ReadSetting("show_graphically"))
'        fUseMeter = (Val(mCurrent.ReadSetting("show_graphically")) <> 0)
'
'        If (mCurrent.Flags And E_METER_STYLE_VARIABLE_METERS) Then
'            .Add new_BPrefsControl("fancycycle", "meter_type", mCurrent.MeterTypes, "Type:", mCurrent.ReadSetting("meter_type"), , fUseMeter)
'
'            If (mCurrent.Flags And E_METER_STYLE_HAS_SPECTRUMS) Then
'
'                .Add new_BPrefsControl("fancycolourpicker", "bar_colour", "", "Bar colour:", mCurrent.ReadSetting("bar_colour"))
'
'                Set pm = New CTempMsg
'                pm.Add "item-height", 24&
'                Set pc = new_BPrefsControl("combo", "spectrum_type", mCurrent.Spectrums, "Spectrum type:", mCurrent.ReadSetting("spectrum_type"), pm, (InStr(mCurrent.MeterType(Val(mCurrent.ReadSetting("meter_type"))), "Spectrum")))
'                .Add pc
'
'                uSetSpectrumIcon pc, 1, 0
'                uSetSpectrumIcon pc, 2, 1
'
'            End If
'        End If
'
'        If (mCurrent.Flags And E_METER_STYLE_NO_SHOW_PERCENT) = 0 Then _
'            .Add new_BPrefsControl("fancytoggle2", "show_percent", "Display value as percentage?", "Options:", mCurrent.ReadSetting("show_percent"))
'
'
'        ' /* debugging section - Sony only ATM (and only if CTRL held down) */
'
'        If g_IsPressed(VK_LCONTROL) Then
''            If mPanelStyle = STYLE_NAME_SONY Then
''            .Add new_BPrefsControl("separator", "", "")
'                .Add new_BPrefsControl("banner", "", "Debugging")
'                Set pm = New CTempMsg
'                pm.Add "min", -1&
'                pm.Add "max", 1439&
'                pm.Add "freq", 120&
'                pm.Add "label_size", 50&
'                .Add new_BPrefsControl("fancyslider", "fs>dbgTime", "", "Time of day:", "0", pm)
'
''            End If
'        End If
'
'    End With

                
                
                
            
            End With

            .AddPage pp
            
            .Go

        End With

    End If

    IStyleEngine_GetConfigWindow = mPanel.hWnd

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
            .Flags = S_STYLE_CUSTOM_SHOW Or S_STYLE_CUSTOM_HIDE Or S_STYLE_SINGLE_INSTANCE Or S_STYLE_DONT_QUEUE Or S_STYLE_IS_CONFIGURABLE
            .Schemes = "Top|Bottom|Top Left|Top Right|Bottom Left|Bottom Right"

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

Private Function uCreatePackedFontStr(ByVal Typeface As String, ByVal TypeSize As String) As String
Dim ppd As BPackedData

    Set ppd = New BPackedData
    ppd.Add "name", Typeface
    ppd.Add "size", TypeSize
    uCreatePackedFontStr = ppd.AsString()

End Function
