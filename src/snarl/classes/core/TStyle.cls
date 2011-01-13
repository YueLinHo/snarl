VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type T_INFO
    Engine As IStyleEngine
    SchemeCount As Long
    Scheme() As String
    Info As style_info

End Type

Dim mContent As T_INFO

Private Sub Class_Initialize()

    Set mContent.Engine = New TInternalStyle
    mContent.Info.Path = g_MakePath(App.Path) & "etc\default_theme"

End Sub

Friend Sub bInit(ByRef Engine As IStyleEngine, ByRef Info As style_info)

    With mContent
        Set .Engine = Engine
        LSet .Info = Info

    End With

    ' /* get the schemes */

    uGetSchemes

End Sub

Friend Sub bZapEngine()

    Set mContent.Engine = Nothing

End Sub

Private Sub uGetSchemes()
Dim sz() As String
Dim c As Long
Dim i As Long

    On Error Resume Next

    With mContent
        .SchemeCount = 0
        ReDim .Scheme(0)

    End With

    sz() = Split(mContent.Info.Schemes, "|")
    c = UBound(sz()) + 1
    If c > 0 Then
        For i = 0 To (c - 1)
            sz(i) = trim(sz(i))
            If sz(i) <> "" Then
                With mContent
                    .SchemeCount = .SchemeCount + 1
                    ReDim Preserve .Scheme(.SchemeCount)
                    .Scheme(.SchemeCount) = sz(i)

                End With
            End If
        Next i
    End If

    With mContent
        If .SchemeCount = 0 Then
            g_Debug "TStyle.uGetSchemes(): style didn't provide any valid schemes", LEMON_LEVEL_WARNING
            .SchemeCount = 1
            ReDim .Scheme(.SchemeCount)
            .Scheme(.SchemeCount) = "<Default>"

        End If

    End With

End Sub

'Public Function Info() As style_info
'
'    LSet Info = mInfo
'
'End Function

Public Function CreateInstance() As IStyleInstance

    If Not (mContent.Engine Is Nothing) Then _
        Set CreateInstance = mContent.Engine.CreateInstance(mContent.Info.Name)

End Function

Public Function CreateNotificationInfo(ByVal Title As String, ByVal Text As String, ByVal Icon As String, ByVal Scheme As String, ByVal Flags As S_NOTIFICATION_FLAGS) As notification_info
Dim pInfo As notification_info

    With pInfo
        .Title = Title
        .Text = Text
        .Icon = Icon
        .Scheme = Scheme
        .Flags = Flags

    End With

    LSet CreateNotificationInfo = pInfo

End Function

Public Function CountSchemes() As Long

    CountSchemes = mContent.SchemeCount

End Function

Public Function SchemeAt(ByVal Index As Long) As String

    If (Index > 0) And (Index <= mContent.SchemeCount) Then _
        SchemeAt = mContent.Scheme(Index)

End Function

Public Function IconPath() As String

    IconPath = mContent.Info.IconPath

    ' /* backwards compatability */

    If Not g_Exists(IconPath) Then _
        IconPath = g_MakePath(mContent.Info.Path) & mContent.Info.Name & ".png"

    If Not g_Exists(IconPath) Then _
        IconPath = g_MakePath(mContent.Info.Path) & "icon.png"

    ' /* still no good? let caller decide what to do */

    If Not g_Exists(IconPath) Then _
        IconPath = ""

End Function

Public Function Name() As String

    Name = mContent.Info.Name

End Function

Public Function Flags() As S_STYLE_FLAGS

    Flags = mContent.Info.Flags

End Function

Public Function SchemeIndex(ByVal Scheme As String) As Long
Dim i As Long

    Scheme = LCase$(Scheme)

    With mContent
        If .SchemeCount Then
            For i = 1 To .SchemeCount
                If LCase$(.Scheme(i)) = Scheme Then
                    SchemeIndex = i
                    Exit Function

                End If
            Next i
        End If
    End With

End Function
'
'Friend Function IndexOfScheme2(ByVal StyleIndex As Long, ByVal Scheme As String) As Long
'Dim i As Long
'
'    If (StyleIndex < 1) Or (StyleIndex > mItems) Then _
'        Exit Function
'
'    Scheme = LCase$(Scheme)
'
'    With mItem(StyleIndex)
'        If .SchemeCount Then
'            For i = 1 To .SchemeCount
'                If LCase$(.Scheme(i)) = Scheme Then
'                    IndexOfScheme2 = i
'                    Exit Function
'
'                End If
'            Next i
'        End If
'    End With
'
'End Function


Public Function Path() As String

    Path = g_MakePath(mContent.Info.Path)

End Function

Public Function Description() As String

    Description = mContent.Info.Description

End Function

Public Function Engine() As IStyleEngine

    Set Engine = mContent.Engine

End Function

Public Function Copyright() As String

    Copyright = mContent.Info.Copyright

End Function

Public Function URL() As String

    URL = mContent.Info.URL

End Function

Public Function Release() As String

    Release = CStr(mContent.Info.Major) & "." & CStr(mContent.Info.Minor)

End Function

Public Function GetSchemeList(Optional ByVal Refresh As Boolean) As String

    If Refresh Then
        uUpdateStyleInfo
        uGetSchemes

    End If

Dim i As Long

    With mContent
        If .SchemeCount Then
            For i = 1 To .SchemeCount
                GetSchemeList = GetSchemeList & IIf(.Scheme(i) = "", "<default>", .Scheme(i)) & "|"

            Next i
        End If
    End With

    GetSchemeList = g_SafeLeftStr(GetSchemeList, Len(GetSchemeList) - 1)

    If GetSchemeList = "" Then _
        GetSchemeList = "<Default>"

End Function

Private Sub uUpdateStyleInfo()
Dim ps As style_info
Dim i As Long

    With mContent.Engine
        For i = 1 To .CountStyles
            .StyleAt i, ps
            If ps.Name = mContent.Info.Name Then
                LSet mContent.Info = ps
                Exit Sub

            End If
        Next i

    End With

    Debug.Print "TStyle.uUpdateStyleInfo(): couldn't find style '" & mContent.Info.Name & "'"

End Sub
