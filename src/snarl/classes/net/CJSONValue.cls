VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJSONValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type T_VALUE
    Name As String
    Value As String
    LongValue As Long
    BoolValue As Boolean
    DataType As E_JSON_DATATYPES
    IsValid As Boolean
    Object As CJSONObject

End Type

Dim mContent As T_VALUE

Friend Function bInit(ByVal Raw As String) As Boolean
Dim i As Integer

    Raw = Replace$(Raw, Chr$(255), ",")             ' // put the commas back

    i = InStr(Raw, ":")                             ' // first colon is name delimiter
    If i = 0 Then _
        Exit Function

    With mContent
        .Name = g_SafeLeftStr(Raw, i - 1)
        .Value = g_SafeRightStr(Raw, Len(Raw) - i)

        ' /* validate the name */

        If (g_SafeLeftStr(.Name, 1) <> Chr$(34)) Or (g_SafeRightStr(.Name, 1) <> Chr$(34)) Then _
            Exit Function

        .Name = mId$(.Name, 2, Len(.Name) - 2)

        ' /* determine the value type */

        If (g_SafeLeftStr(.Value, 1) = "{") And (g_SafeRightStr(.Value, 1) = "}") Then
            ' /* an object */
            .DataType = E_OBJECT
            Set .Object = New CJSONObject
            .IsValid = .Object.SetFrom(.Value)

        ElseIf (g_SafeLeftStr(.Value, 1) = Chr$(34)) Or (g_SafeRightStr(.Value, 1) = Chr$(34)) Then
            ' /* is a string */
            .DataType = E_STRING
            .Value = mId$(.Value, 2, Len(.Value) - 2)
            .IsValid = True

        ElseIf g_IsNumeric(.Value) Then
            ' /* number */
            .DataType = E_NUMBER
            .LongValue = Val(.Value)
            .BoolValue = Not .LongValue
            .IsValid = True

        Else
            Select Case LCase$(.Value)
            Case "true", "false"
                .DataType = E_BOOL
                .IsValid = True
                .BoolValue = (LCase$(.Value) = "true")
                .LongValue = IIf(.BoolValue, 1, 0)

            End Select

        End If

    End With

End Function

Public Function IsValid() As Boolean

    IsValid = mContent.IsValid

End Function

Public Function Name() As String

    Name = mContent.Name

End Function

Public Function AsString() As String

    AsString = mContent.Value

End Function

Public Function AsLong() As Long

    AsLong = Val(mContent.LongValue)

End Function

Public Function AsBool() As Boolean

    AsBool = Val(mContent.BoolValue)

End Function

Public Function Value() As String

    Value = mContent.Value

End Function

Public Function DataType() As E_JSON_DATATYPES

    DataType = mContent.DataType

End Function

Public Function AsObject() As CJSONObject

    Set AsObject = mContent.Object

End Function
