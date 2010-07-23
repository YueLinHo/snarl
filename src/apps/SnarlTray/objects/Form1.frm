VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SnarlTray Log"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mTrayIcon As BNotifyIcon
Dim mClassId As String

Private Sub Command1_Click()

    If Not (mTrayIcon Is Nothing) Then _
        mTrayIcon.ShowInfo "1", "SnarlTray", "Test Message", B_NOTIFY_ICON_INFO

End Sub

Private Sub Command3_Click()

    PostMessage Val(Me.Tag), sn41AppMsg(), SNARL41_APP_PREFS, ByVal 0&

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then _
        PostQuitMessage 0

End Sub

Public Sub Add(ByVal Text As String)

    With Form1.List1
        .AddItem Text
        .ListIndex = .ListCount - 1

    End With

End Sub

Public Sub InstallIcon()

    Set mTrayIcon = New BNotifyIcon
    With mTrayIcon
        .SetTo Me.hWnd, &H401
        .Add "1", Me.Icon.Handle, "SnarlTray"

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not (mTrayIcon Is Nothing) Then
        mTrayIcon.Remove "1"
        Set mTrayIcon = Nothing

    End If

End Sub
