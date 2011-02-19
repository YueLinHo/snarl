Attribute VB_Name = "mSnarlStyleHelpers_i"
Option Explicit

    ' /*
    '
    '   mSnarlStyleHelpers_i.bas -- Snarl Visual Basic 5/6 style engine helpers
    '
    '   � 2010 full phat products.  All Rights Reserved.
    '
    '        Version: 41 (R2.31)
    '       Revision: 1
    '        Created: 20-Sep-2010
    '   Last Updated:
    '         Author: C. Peel (aka Cheekiemunkie)
    '        Licence: Simplified BSD License (http://www.opensource.org/licenses/bsd-license.php)
    '
    '   Notes
    '   -----
    '
    '   Simple style engine helpers
    '
    ' */

Private Const CSIDL_APPDATA = &H1A
Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWndOwner As Long, ByVal lpszPath As String, ByVal nFolder As Long, ByVal fCreate As Boolean) As Long

Public Function style_GetStyleName(ByVal StyleAndScheme As String) As String
Dim i As Long
    
    On Error Resume Next

    i = InStr(StyleAndScheme, "/")
    If i Then
        style_GetStyleName = Left$(StyleAndScheme, i - 1)

    Else
        style_GetStyleName = StyleAndScheme

    End If

End Function

Public Function style_GetSchemeName(ByVal StyleAndScheme As String) As String
Dim i As Long

    On Error Resume Next

    i = InStr(StyleAndScheme, "/")
    If i Then
        style_GetSchemeName = Right$(StyleAndScheme, Len(StyleAndScheme) - i)
        i = InStr(style_GetSchemeName, "|")
        If i Then _
            style_GetSchemeName = Left$(style_GetSchemeName, i - 1)
            
    End If

End Function

Public Function style_GetNotificationFlags(ByVal StyleAndScheme As String) As String
Dim i As Long

    On Error Resume Next

    i = InStr(StyleAndScheme, "|")
    If i Then _
        style_GetNotificationFlags = Right$(StyleAndScheme, Len(StyleAndScheme) - i)

End Function

Public Function styles_SchemesToMessage(ByVal Schemes As String) As MMessage
Dim pm As CTempMsg
Dim sz() As String
Dim i As Long

    If Schemes = "" Then _
        Exit Function

    sz() = Split(Schemes, "|")

    Set pm = New CTempMsg
    pm.What = UBound(sz()) + 1

    For i = 0 To pm.What - 1
        pm.Add CStr(i + 1), sz(i)

    Next i

    Set styles_SchemesToMessage = pm

End Function

Public Function style_GetSnarlConfigPath(ByVal StyleName As String) As String
Dim sz As String

    sz = String$(4096, 0)
    If SHGetSpecialFolderPath(0, sz, CSIDL_APPDATA, False) Then _
        style_GetSnarlConfigPath = g_MakePath(uTrimStr(sz)) & "full phat\snarl\etc\." & StyleName

End Function

Private Function uTrimStr(ByVal sz As String) As String
Dim i As Long

    i = InStr(sz, Chr$(0))
    If i Then
        uTrimStr = Left$(sz, i - 1)
    Else
        uTrimStr = sz
    End If

End Function

'Public Function style_IsValidImage(ByRef Image As MImage) As Boolean
'
'    If (Image Is Nothing) Then _
'        Exit Function
'
'    style_IsValidImage = ((Image.Width > 0) And (Image.Height > 0))
'
'End Function

Public Function style_MakeSquareImage(ByRef Img As MImage, Optional ByVal Maximum As Long) As mfxBitmap
Dim pv As mfxView
Dim c As Long

    If Not is_valid_image(Img) Then _
        Exit Function

    c = MAX(Img.Width, Img.Height)
    If Maximum > 0 Then _
        c = MIN(c, Maximum)

    Set pv = New mfxView
    With pv
        .SizeTo c, c
        .DrawScaledImage Img, new_BPoint(Fix((.Width - Img.Width) / 2), Fix((.Height - Img.Height) / 2))
        Set style_MakeSquareImage = .ConvertToBitmap()

    End With

End Function

