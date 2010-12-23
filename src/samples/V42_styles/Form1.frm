VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   5850
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
   ScaleHeight     =   5850
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   60
      TabIndex        =   14
      Text            =   "primary::36°"
      Top             =   4740
      Width           =   4395
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Custom background colour (web format #rrggbb)"
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   2880
      Width           =   3915
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   60
      TabIndex        =   12
      Top             =   3240
      Width           =   4395
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   4020
      Width           =   4395
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show Icon"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   3660
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show meter"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   10
      Left            =   60
      Max             =   100
      TabIndex        =   7
      Top             =   2460
      Value           =   23
      Width           =   4395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   5220
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":000C
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
      Top             =   5220
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Extra Data"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   4440
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "@"
      Height          =   195
      Left            =   3720
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Percent"
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

Dim mToken As Long
Dim mMsg As Long

Private Sub Command1_Click()

    If mMsg Then _
        sn41EZUpdate mMsg, , Text2.Text, , , IIf(Check1.Value = vbChecked, "percent::" & CStr(HScroll1.Value), "")

End Sub

Private Sub Command4_Click()
Dim szIcon As String

    szIcon = Text3.Text
    If g_SafeLeftStr(szIcon, 2) = ".\" Then _
        szIcon = g_MakePath(App.Path) & g_SafeRightStr(szIcon, Len(szIcon) - 2)


    If mToken Then _
        mMsg = sn41EZNotify(mToken, "", _
                            Text1.Text, Text2.Text, 0, _
                            IIf(Check2.Value = vbChecked, szIcon, ""), , , , , _
                            IIf(Check1.Value = vbChecked, "percent::" & CStr(HScroll1.Value), "") & _
                                IIf(Check3.Value = vbChecked, "#?col::" & Text4.Text, "") & _
                                IIf(Text5.Text <> "", "#?" & Text5.Text, ""))

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
            mToken = hr

        End If

        Text3.Text = ".\icon.png"
        Label4.Caption = HScroll1.Value & "%"

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = sn41UnregisterApp(mToken)
    If hr = 0 Then
        Debug.Print "FAILED: " & sn41GetLastError()

    Else
        Debug.Print "OK: " & hr

    End If

End Sub

Private Sub HScroll1_Change()

    Label4.Caption = HScroll1.Value & "%"

End Sub

Private Sub HScroll1_Scroll()

    Label4.Caption = HScroll1.Value & "%"

End Sub

