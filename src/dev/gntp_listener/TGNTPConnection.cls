VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGNTPConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Public Sub Accept(ByVal Request As Long)

    Set theSocket = New CSocket
    theSocket.Accept Request

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim ppd As BPackedData
Dim szReply As String
Dim sz As String
Dim b As Boolean

    theSocket.PeekData sz

    If g_SafeRightStr(sz, 4) = vbCrLf & vbCrLf Then
'        Debug.Print "end of packet"

        theSocket.GetData sz
        If gntp_Process(sz, szReply, theSocket) Then
            Form1.Output "success! response=" & Replace$(szReply, vbCrLf, "¶")
            theSocket.SendData szReply

        Else
            Form1.Output "failed: response=" & Replace$(szReply, vbCrLf, "¶")
            theSocket.SendData szReply

        End If

        theSocket.CloseSocket

    End If

End Sub

