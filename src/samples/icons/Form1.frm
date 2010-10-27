VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   3570
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
   ScaleHeight     =   3570
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Icon"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Image"
      Height          =   495
      Left            =   2940
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   2460
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":628A
      Top             =   1080
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Notification Title"
      Top             =   360
      Width           =   4395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Icon"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Left            =   60
      TabIndex        =   2
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

Private Declare Function SHChangeIconDialog Lib "shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal Reserved As Long, lpIconIndex As Long) As Long

Dim mAppToken As Long

Private Sub Command1_Click()

    With New CFileDialog
        .DialogType = E_DIALOG_OPEN
        .Title = "Pick image file..."

        If .Go(False, E_FILE_DIALOG_CENTRE_SCREEN) Then
            Text3.Text = .SelectedFile


        End If

    End With

End Sub

Private Sub Command2_Click()
Dim sz As String
Dim dw As Long

    sz = String$(260, 0)
    If SHChangeIconDialog(Me.hWnd, sz, 260, dw) <> 0 Then _
        Text3.Text = g_TrimStr(StrConv(sz, vbFromUnicode)) & ",-" & CStr(dw + 1)

End Sub

Private Sub Command4_Click()
Dim sz As String

    If Text3.Text = "" Then
        sz = "%" & Me.Icon.Handle

    Else
        sz = Text3.Text

    End If

    If mAppToken Then _
        sn41EZNotify mAppToken, "", Text1.Text, Text2.Text, , sz

End Sub

Private Sub Form_Load()
Dim hr As Long

    If Not sn41IsSnarlRunning() Then
        MsgBox "Snarl isn't running - launch Snarl, then run this demo.", vbExclamation Or vbOKOnly, App.Title
        Unload Me

    Else
        hr = sn41RegisterApp(App.ProductName, App.Title, App.Path & "\icon.png")
        If hr = 0 Then
            Me.Caption = "Error registering with Snarl: " & sn41GetLastError()

        Else
            Me.Caption = "Registered with Snarl V" & CStr(sn41GetVersion()) & " (" & Hex$(hr) & ")"
            mAppToken = hr

        End If

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = sn41UnregisterApp(mAppToken)
    If hr = 0 Then
        Debug.Print "FAILED: " & sn41GetLastError()

    Else
        Debug.Print "OK: " & hr

    End If

End Sub

