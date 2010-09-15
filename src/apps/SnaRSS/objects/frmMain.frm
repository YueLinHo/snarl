VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "snaRSS Log"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6870
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
      Height          =   2790
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then _
        PostQuitMessage 0

End Sub

Public Sub Add(ByVal Text As String)

    List1.AddItem Text
    List1.ListIndex = List1.ListCount - 1
    g_Debug Text

End Sub
