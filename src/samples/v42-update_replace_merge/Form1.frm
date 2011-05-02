VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   5055
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
   ScaleHeight     =   5055
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Merge"
      Height          =   255
      Index           =   2
      Left            =   2820
      TabIndex        =   9
      Top             =   2520
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Update"
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Replace"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   2520
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include icon"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":000C
      Top             =   1080
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Lorem ipsum dolor sit amet"
      Top             =   360
      Width           =   4395
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":008F
      Height          =   1215
      Left            =   60
      TabIndex        =   10
      Top             =   3420
      Width           =   4395
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   4740
      Width           =   4395
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1155
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
Dim sz As String

    If mToken Then
        sz = "notify?app-sig=" & App.ProductName & "&title=" & Text1.Text & "&text=" & Text2.Text & _
             IIf(Check1.Value = vbChecked, "&icon=" & App.Path & "\icon.png", "") & "&timeout=0&uid=12345"

        If Option1(0).Value = True Then
            sz = sz & "&replace-uid=12345"

        ElseIf Option1(1).Value = True Then
            sz = sz & "&update-uid=12345"

        ElseIf Option1(2).Value = True Then
            sz = sz & "&merge-uid=12345"

        End If

        mMsg = snDoRequest(sz)

    End If

    Label3.Caption = "result: " & CStr(mMsg)

End Sub

'Private Sub Command4_Click()
'Dim pri As Long
'
'    Select Case Combo1.ListIndex
'    Case 0
'        pri = -1
'
'    Case 1
'        pri = 0
'
'    Case 2
'        pri = 1
'
'    End Select
'
'    If mToken Then _
'        mMsg = snDoRequest("notify?app-sig=" & App.ProductName & "&title=" & Text1.Text & "&text=" & Text2.Text & _
'                            IIf(Check1.Value = vbChecked, "&icon=" & App.Path & "\icon.png", "") & _
'                            "&timeout=0&priority=" & CStr(pri) & "&reply-to=" & CStr(Me.hWnd) & "&reply=" & CStr(&H401) & _
'                            "&uid=12345")
'
'    Label3.Caption = "result: " & CStr(mMsg)
'
'End Sub

Private Sub Form_Load()
Dim hr As Long

    If Not snIsSnarlRunning() Then
        MsgBox "Snarl isn't running - launch Snarl, then run this demo.", vbExclamation Or vbOKOnly, App.Title
        Unload Me

    Else
        hr = snarl_register(App.ProductName, App.Title, App.Path & "\icon.png")
        If hr >= 0 Then
            Me.Caption = "Registered with Snarl V" & CStr(snDoRequest("version")) & " (" & Hex$(hr) & ")"
            mToken = hr

        Else
            Me.Caption = "Error registering with Snarl: " & CStr(hr)

        End If

        Text2.Text = "Sample text"

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = snarl_unregister(mToken)
    If hr = 0 Then
        Debug.Print "Unregistered"

    Else
        Debug.Print "Unregister failed: " & Abs(hr)

    End If

End Sub
