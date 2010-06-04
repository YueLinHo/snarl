VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snarl time format sample"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   315
      Left            =   3780
      TabIndex        =   5
      Top             =   1500
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   1500
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1500
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Make sure the Meter/Clock style is selected in Snarl for this application before clicking ""Go""."
      Height          =   435
      Left            =   60
      TabIndex        =   4
      Top             =   1980
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":0000
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   4515
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":008E
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mToken As Long

Private Sub Command1_Click()

    sn41EZNotify mToken, "", Me.Caption, Text1.Text

End Sub

Private Sub Command2_Click()

    sn41EZNotify mToken, "", "", Text1.Text

End Sub

Private Sub Form_Load()

    mToken = sn41RegisterApp(App.FileDescription, Me.Caption, "")
    If mToken = 0 Then
        MsgBox "Unable to register with Snarl.  Ensure Snarl is running and then re-run this application.", vbCritical Or vbOKOnly, Me.Caption
        Unload Me

    Else
        Text1.Text = Format$(Now(), "yyyymmddhhnnss")

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    sn41UnregisterApp mToken

End Sub
