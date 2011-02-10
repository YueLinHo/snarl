VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   4395
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":000C
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mToken As Long
Dim mMsg As Long

Private Sub Command1_Click()

'    If mMsg Then _
        sn41EZUpdate mMsg, , Text2.Text

End Sub

Private Sub Command4_Click()

    If mToken Then _
        mMsg = snDoRequest("notify?token=" & CStr(mToken) & "&title=Lorem ipsum dolor sit amet" & _
                            "&text=" & "Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua." & _
                            "&icon-base64=" & uGetIconData(True))

    Label3.Caption = "result: " & CStr(mMsg)

End Sub

Private Sub Form_Load()
Dim hr As Long

    If Not snIsSnarlRunning() Then
        MsgBox "Snarl isn't running - launch Snarl, then run this demo.", vbExclamation Or vbOKOnly, App.Title
        Unload Me

    Else
        hr = snDoRequest("register?app-sig=" & App.ProductName & "&title=" & App.Title & "&icon-base64=" & uGetIconData())
        If hr < 0 Then
            Me.Caption = "Error registering with Snarl: " & Abs(hr)

        Else
            Me.Caption = "Registered with Snarl V" & CStr(snDoRequest("version")) & " (" & Hex$(hr) & ")"
            mToken = hr

        End If

'        Combo1.ListIndex = 1

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = snDoRequest("unregister?token=" & CStr(mToken))
    If hr < 0 Then
        Debug.Print "FAILED: " & Abs(hr)

    Else
        Debug.Print "OK"

    End If

End Sub

Private Function uGetIconData(Optional ByVal PickAtRandom As Boolean) As String
Dim sz As String
Dim i As Integer
Dim n As Integer
Dim s As String

    Randomize Timer

    s = "icon"

    If PickAtRandom Then _
        s = s & "-" & CStr(CInt(Rnd * 3) + 1)

    s = s & ".png"

    i = FreeFile()
    Open App.Path & "\" & s For Binary Access Read Lock Write As #i
    sz = String$(LOF(i), Chr$(0))
    Get #i, , sz
    Close #i

    uGetIconData = Replace$(Encode64(sz), "=", "%")

End Function
