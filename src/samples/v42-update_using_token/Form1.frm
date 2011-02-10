VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   5205
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
   ScaleHeight     =   5205
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Random Icon"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   1620
      TabIndex        =   7
      Top             =   3660
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Text"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Value           =   1  'Checked
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   3660
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
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
      Height          =   615
      Left            =   60
      TabIndex        =   4
      Top             =   4200
      Width           =   4395
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   4860
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
Dim sz As String
Dim hr As Long

    mMsg = 0

    ' /* show */

    If mToken Then
        sz = "notify?app-sig=" & App.ProductName & _
             IIf(Check1(0).Value = vbChecked, "&title=" & Text1.Text, "") & _
             IIf(Check1(1).Value = vbChecked, "&text=" & Text2.Text, "") & _
             "&icon=" & App.Path & "\icon.png&timeout=0"

        hr = snDoRequest(sz)
        If hr > 0 Then _
            mMsg = hr

    End If

    Label3.Caption = "show: " & CStr(hr)

End Sub

Private Sub Command2_Click()
Dim sz As String
Dim hr As Long
Dim i As Integer

    Randomize Timer

    If mMsg = 0 Then
        Label3.Caption = "Click 'Show' to create a notification"

    Else
        sz = "update?token=" & CStr(mMsg) & _
             IIf(Check1(0).Value = vbChecked, "&title=" & Text1.Text, "") & _
             IIf(Check1(1).Value = vbChecked, "&text=" & Text2.Text, "")

        If Check1(2).Value = vbChecked Then
            i = (Rnd * 5) + 1
            sz = sz & "&icon=" & App.Path & "\" & CStr(i) & ".png"

        End If

        hr = snDoRequest(sz)

    End If

    Label3.Caption = "update: " & CStr(hr)

End Sub

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

    If mToken Then
        hr = snarl_unregister(mToken)
        If hr = 0 Then
            Debug.Print "Unregistered"

        Else
            Debug.Print "Unregister failed: " & Abs(hr)

        End If
    End If

End Sub
