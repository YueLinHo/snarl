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

    ' /*
    '
    '   StyleEngine.cls --
    '
    '   © 2010 full phat products.  Some Rights Reserved.
    '   Licenced as per the Simplified BSD (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   Standardised V42 entries accepted:
    '
    '   Generic groups
    '   --------------
    '   value-:         must be a numeric value, may be a decimal or integer
    '   label-:         may be alphanumeric
    '   colour-:        may use HTML notation (i.e. #RRGGBB) or rgb notation (i.e. r,g,b)
    '   data-:          free for custom use
    '
    '   Defined Entries
    '   ---------------
    '   value-percent:      must be a numeric value in the range 0 to 100.  May be fractional although styles may round to
    '                       nearest integer at their discretion.
    '
    '   label-subtext:      text which is typically displayed inferior to the notification text.
    '   label-tagline:      a short piece of text typically displayed in a prominent way.
    '
    '   colour-background:  requested background colour.
    '   colour-text:        requested text colour.
    '   colour-subtext:     requested subtext colour.
    '
    ' */

Private Const LIB_DATE = "19-Apr-2011"

Private Const STYLE_NAME_TOASTY = "Toasty"
Private Const STYLE_NAME_NMA = "Notify My Android"

Dim mLastErr As String
Dim mStyle() As TStyle
Dim mCount As Long

Dim mPanel As BPrefsPanel
Dim mPage As BPrefsPage
Dim mNoWriteConfig As Boolean

Dim mCurrent As TStyle

Implements KPrefsPage
Implements KPrefsPanel
Implements IStyleEngine

Private Function IStyleEngine_CountStyles() As Long

    IStyleEngine_CountStyles = mCount

End Function

Private Function IStyleEngine_CreateInstance(ByVal StyleName As String) As libSnarlStyles.IStyleInstance
Dim pStyle As TGenericPhoneInstance

    Select Case StyleName
    Case STYLE_NAME_TOASTY
        Set pStyle = New TGenericPhoneInstance

    Case STYLE_NAME_NMA
        Set pStyle = New TGenericPhoneInstance
        pStyle.IsNotifyMyAndroid = True

    End Select

    If Not (pStyle Is Nothing) Then
        pStyle.SetStyle mStyle(uIndexOf(StyleName))
        Set IStyleEngine_CreateInstance = pStyle

    End If

End Function

Private Function IStyleEngine_Date() As String

    IStyleEngine_Date = LIB_DATE

End Function

Private Function IStyleEngine_Description() As String

    IStyleEngine_Description = App.FileDescription

End Function

Private Function IStyleEngine_GetConfigWindow(ByVal StyleName As String) As Long
Dim pp As BPrefsPage

    If (mPanel Is Nothing) Then

        Set mCurrent = mStyle(uIndexOf(StyleName))

        If (mCurrent Is Nothing) Then
            MsgBox "Error setting style", vbCritical, App.Title
            Exit Function

        End If

        Set mPanel = New BPrefsPanel
        With mPanel
            .SetHandler Me
            .SetTitle StyleName & " Preferences"
            .SetWidth 300
            .SetWindow 1

            Set pp = new_BPrefsPage("", , Me)

            With pp
                If mCurrent.Name = "Notify My Android" Then
                    .Add new_BPrefsControl("fancyedit", "api_key", mCurrent.ReadSetting("api_key"), "API Key:")

                Else
                    .Add new_BPrefsControl("fancyedit", "device_id", mCurrent.ReadSetting("device_id"), "Device ID:")
            
                End If

            End With
            
            .AddPage pp
            .Go

'    g_SetWindowIconToAppResourceIcon mPanel.hWnd

        End With

    End If

    IStyleEngine_GetConfigWindow = mPanel.hWnd

End Function

Private Function IStyleEngine_Initialize() As melon.M_RESULT

    gID = 63

    ' /* add our styles */

    uAdd STYLE_NAME_TOASTY, _
         "Forwards the notification content to a Windows Phone 7 device via Toasty.  See http://supertoasty.com/ for more information.", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME, _
         "", _
         0, 3

    uAdd STYLE_NAME_NMA, _
         "Forwards the notification content to an Android device via Notify My Android.  You will need to register an API key at http://nma.usk.bz first.", _
         S_STYLE_IS_CONFIGURABLE Or S_STYLE_MULTI_INSTANCE Or S_STYLE_IS_WINDOWLESS Or S_STYLE_WANT_APP_NAME, _
         "", _
         0, 1


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

    If (Index < 1) Or (Index > mCount) Then _
        Exit Sub

    mStyle(Index).GetInfo Style

End Sub

Private Sub IStyleEngine_TidyUp()
End Sub

Private Function IStyleEngine_Version() As Long

    IStyleEngine_Version = App.Major

End Function

Private Sub KPrefsPage_AllAttached()
'Dim pc As BControl
'
'    If mPanel.Find("fs>dbgTime", pc) Then _
'        pc.SetValue "-1"

End Sub

Private Sub KPrefsPage_Attached()
'Dim pc As BControl
'Dim pm As CTempMsg
'Dim dw As Long
'Dim sz As String
'
'
'Dim fUseMeter As Boolean

    ' /* construct based on the style we're configuring */

    With mPage

''        .Add new_BPrefsControl("banner", "", "General Settings")
'
'        ' /* font pickers */
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_CHANGE_FONT) Then
'            .Add new_BPrefsControl("fancyfontpicker", "title", "", "Title font:", uCreatePackedFontStr(mCurrent.ReadSetting("title_font"), mCurrent.ReadSetting("title_size")))
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
'        ' /* foreground opacity */
'
'        If (mCurrent.Flags And E_METER_STYLE_CAN_SET_BG_ALPHA) Then
'            Set pm = New CTempMsg
'            pm.Add "min", 0&
'            pm.Add "max", 100&
'            pm.Add "freq", 10&
'            pm.Add "suffix", "%"
'            .Add new_BPrefsControl("fancyslider", "bg_alpha", "", "Background opacity:", mCurrent.ReadSetting("bg_alpha"), pm)
'
'        End If
'
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
'        ' /* test and reset buttons - all */
'
'        .Add new_BPrefsControl("separator", "", "")
'        .Add new_BPrefsControl("fancybutton2", "fb2>Test", "Test")
'        .Add new_BPrefsControl("separator", "", "")
'        .Add new_BPrefsControl("fancybutton2", "fb2>Reset", "Reset to Factory Settings")
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

    End With

End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Debug.Print Control.GetName() & " > " & Value

    ' /* write settings back to config file */

    mCurrent.WriteSetting Control.GetName, Value


'    Select Case Control.GetName()
'    Case "title", "text"
'        ' /* fonts need a little extra tweaking... */
'        With New BPackedData
'            If .SetTo(Value) Then
'                mCurrent.WriteSetting Control.GetName & "_font", .ValueOf("name")
'                mCurrent.WriteSetting Control.GetName & "_size", .ValueOf("size")
'
'            End If
'
'        End With
'
'    Case Else
'        ' /* ...everything else is straightforward */
'
'    End Select
'
'    ' /* deal with control setting */
'
'Dim pc As BControl
'
'    Select Case Control.GetName()
'
'    Case "meter_type"
'        If Control.Page.Find("spectrum_type", pc) Then _
'            pc.SetEnabled (InStr(mCurrent.MeterType(Val(Value)), "Spectrum"))
'
'        If Control.Page.Find("bar_colour", pc) Then _
'            pc.SetEnabled (InStr(mCurrent.MeterType(Val(Value)), "Coloured"))
'
''    ' /* Sony only */
''
''    Case "fs>dbgTime"
''        ' /* debug mode minute slider */
''        gDbgMinute = Val(Value)
''        If Val(Value) = -1 Then
''            Control.SetText "Auto"
''
''        Else
''            Control.SetText Format$(TimeSerial(0, Val(Value), 0), "h:mm AMPM")
''
''        End If
''
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)
'Dim pc As BControl
'
'    Select Case Control.GetName()
'
'    Case "fb2>Test"
'        ' /* test settings */
'        mCurrent.DoPreview
'
'
'    Case "fb2>Reset"
'        ' /* reset to factory defaults */
'
'        mNoWriteConfig = True
'
'        ' /* set our title and text controls - the correct values are determined by
'        '    uGetDefaultFont() and are stored in the right place as setting a
'        '    control's value manually still generates a 'changed' notification */
'
'
'Dim szn As String
'Dim szv As String
'
'        With mCurrent.Defaults
'            .Rewind
'
'            Do While .GetNextItem(szn, szv)
'                If mPage.Find(szn, pc) Then _
'                    pc.SetValue szv
'
'            Loop
'
'        End With
'
'        ' /* fonts */
'
'        If mPage.Find("title", pc) Then _
'            pc.SetValue uCreatePackedFontStr(mCurrent.GetDefault("title_font"), mCurrent.GetDefault("title_size"))
'
'        If mPage.Find("text", pc) Then _
'            pc.SetValue uCreatePackedFontStr(mCurrent.GetDefault("text_font"), mCurrent.GetDefault("text_size"))
'
''            Select Case mPanelStyle
''            Case STYLE_NAME_IPHONEY
''                pc.SetValue CStr(rgba(56, 89, 131))
''
''            Case STYLE_NAME_MINIMAL
''                pc.SetValue CStr(rgba(255, 255, 255))
''
''            End Select
''        End If
''
''        If mPage.Find("foreground", pc) Then
''            Select Case mPanelStyle
''            Case STYLE_NAME_IPHONEY
''                pc.SetValue CStr(rgba(255, 255, 255))
''
''            Case STYLE_NAME_MINIMAL
''                pc.SetValue CStr(rgba(168, 192, 225))
''
''            End Select
''        End If
'
'
'        ' /* set style-specific controls here */
'
''        Select Case mPanelStyle
''        Case STYLE_NAME_SONY
''            ' /* Sony */
''
''            If mPage.Find("fs>DropshadowStrength", pc) Then _
''                pc.SetValue "50"
'
'        mNoWriteConfig = False
''        uWriteConfig
'
'    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As BPrefsPage)

    Set mPage = Page
    mPage.SetMargin 140

End Sub

Private Sub KPrefsPage_Destroy()

    Set mPage = Nothing

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
    Set mCurrent = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Private Sub uAdd(ByVal StyleName As String, ByVal Description As String, ByVal Flags As S_STYLE_FLAGS, ByVal Schemes As String, ByVal Major As Long, ByVal Minor As Long) ', ByRef Defaults As BPackedData, Optional ByRef CustomPrefsPage As KPrefsPage)
Dim ps As TStyle

    Set ps = New TStyle
    ps.Init StyleName, Description, Flags, Schemes, Major, Minor, Nothing, LIB_DATE ', CustomPrefsPage

    mCount = mCount + 1
    ReDim Preserve mStyle(mCount)
    Set mStyle(mCount) = ps

End Sub

Private Function uIndexOf(ByVal Name As String) As Long
Static i As Long

    For i = 1 To mCount
        If mStyle(i).Name = Name Then
            uIndexOf = i
            Exit Function

        End If
    Next i

End Function

'Private Function uCreatePackedStr(ByVal Content As String) As BPackedData
'
'    Set uCreatePackedStr = New BPackedData
'    uCreatePackedStr.SetTo Content
'
'End Function

