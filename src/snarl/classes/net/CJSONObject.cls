VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJSONObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mValue() As CJSONValue
Dim mValues As Long

Public Function SetFrom(ByVal Raw As String) As Boolean

    ' /* must start and end with { } */

    If (g_SafeLeftStr(Raw, 1) <> "{") Or (g_SafeRightStr(Raw, 1) <> "}") Then
        g_Debug "CJSON.SetFrom(): bad data"
        Exit Function

    End If

    ' /* trim { } off */

    Raw = mId$(Raw, 2, Len(Raw) - 2)

    ' /* is it complex?  that is, are objects and arrays defined? */

    If (InStr(Raw, "[")) Or (InStr(Raw, "{")) Then
        g_Debug "CJSON.SetFrom(): is complex"
        uParseObject Raw

    End If

    ' /* now split into values */

Dim sv() As String
Dim i As Long

    sv() = Split(Raw, ",")
    mValues = UBound(sv()) + 1
    ReDim mValue(mValues)
    For i = 1 To mValues
        Set mValue(i) = New CJSONValue
        mValue(i).bInit sv(i - 1)

    Next i

'    Debug.Print Raw
'    Debug.Print "Values: " & mValues
'
'    For i = 1 To mValues
'        Debug.Print "#" & CStr(i) & ": " & mValue(i).Name & " >> " & mValue(i).AsString & " / " & mValue(i).IsValid() & " / " & mValue(i).DataType
'
'    Next i

    SetFrom = True

End Function

Private Sub uParseObject(ByRef Raw As String)
Dim i As Integer
Dim j As Integer
Dim sz As String

    i = InStr(Raw, "{")
    Do While i
        ' /* find end */
        j = InStr(i + 1, Raw, "}")
        If j = 0 Then
            g_Debug "CJSON.uParseObject(): error!"
            Exit Sub

        End If

        sz = mId$(Raw, (i + 1), (j - 1) - i)
        sz = Replace$(sz, ",", Chr$(255))
        Raw = g_SafeLeftStr(Raw, i) & sz & g_SafeRightStr(Raw, Len(Raw) - j + 1)

        i = InStr(j + 1, Raw, "{")

    Loop

End Sub

Public Function Find(ByVal Name As String, ByRef Value As CJSONValue) As Boolean
Dim i As Long

    Name = LCase$(Name)

    If mValues Then
        For i = 1 To mValues
            If LCase$(mValue(i).Name) = Name Then
                Set Value = mValue(i)
                Find = True
                Exit Function

            End If
        Next i
    End If

End Function

Public Function GetBoolean(ByVal Name As String) As Boolean
Dim pv As CJSONValue

    If Find(Name, pv) Then _
        GetBoolean = pv.AsBool()

End Function

Public Function GetLong(ByVal Name As String) As Long
Dim pv As CJSONValue

    If Find(Name, pv) Then _
        GetLong = pv.AsLong()

End Function

Public Function GetString(ByVal Name As String) As String
Dim pv As CJSONValue

    If Find(Name, pv) Then _
        GetString = pv.AsString()

End Function

Public Function GetObject(ByVal Name As String) As CJSONObject

End Function

Public Function CountValues() As Long

    CountValues = mValues

End Function

Public Function ValueAt(ByVal Index As Long) As CJSONValue

    If (Index > 0) And (Index <= mValues) Then _
        Set ValueAt = mValue(Index)

End Function

Public Function SetFromHTTP(ByVal RawHTTP As String) As Boolean
Dim i As Integer

    ' /* first, remove leading and trailing HTTP clutter */

    i = InStr(RawHTTP, "{")
    If i = 0 Then
        g_Debug "CJSONObject.SetFromHTTP(): no JSON found", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    RawHTTP = g_SafeRightStr(RawHTTP, Len(RawHTTP) - (i - 1))
    RawHTTP = StrReverse(RawHTTP)

    i = InStr(RawHTTP, "}")
    If i = 0 Then
        g_Debug "CJSONObject.SetFromHTTP(): no JSON found", LEMON_LEVEL_CRITICAL
        Exit Function

    End If
        
    RawHTTP = g_SafeRightStr(RawHTTP, Len(RawHTTP) - (i - 1))
    RawHTTP = StrReverse(RawHTTP)

    ' /* now validate the JSON itself */

    i = uCount(RawHTTP, "{", "}")
    If i <> 0 Then
        g_Debug "CJSONObject.SetFromHTTP(): bad JSON structure", LEMON_LEVEL_CRITICAL
        Exit Function

    End If
        
    i = uCount(RawHTTP, "[", "]")
    If i <> 0 Then
        g_Debug "CJSONObject.SetFromHTTP(): bad JSON structure", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* lastly, replace any http '%nn' values with their ASCII equivalents */

    SetFromHTTP = SetFrom(uReplacePercent(RawHTTP))

End Function

Private Function uCount(ByVal sz As String, ByVal a1 As String, ByVal a2 As String) As Long
Dim i As Long
Dim a As String

    For i = 1 To Len(sz)
        a = mId$(sz, i, 1)
        If a = a1 Then
            uCount = uCount + 1

        ElseIf a = a2 Then
            uCount = uCount - 1

        End If

    Next i

End Function

Private Function uReplacePercent(ByVal sz As String) As String
Dim i As Integer
Dim j As Integer

    i = InStr(sz, "%")
    Do While i > 0
        j = Val("&H" & mId$(sz, i + 1, 2))
        uReplacePercent = uReplacePercent & g_SafeLeftStr(sz, i - 1) & Chr$(j)
        sz = g_SafeRightStr(sz, Len(sz) - (i + 2))
        i = InStr(sz, "%")

    Loop

    uReplacePercent = uReplacePercent & sz

End Function

