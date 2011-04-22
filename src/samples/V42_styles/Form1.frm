VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@"
   ClientHeight    =   6015
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
   ScaleHeight     =   6015
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Text            =   "mynotification"
      Top             =   1920
      Width           =   4395
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Text            =   "label-subtext=36°"
      Top             =   4980
      Width           =   4395
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Custom background colour (web format #rrggbb)"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   3300
      Width           =   3915
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Text            =   "#789abc"
      Top             =   3600
      Width           =   4395
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   4320
      Width           =   4395
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show Icon"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   4020
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show meter"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   2340
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   10
      Left            =   60
      Max             =   100
      TabIndex        =   3
      Top             =   2880
      Value           =   23
      Width           =   4395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":000C
      Top             =   960
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Text            =   "Notification Title"
      Top             =   300
      Width           =   4395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show"
      Default         =   -1  'True
      Height          =   495
      Left            =   60
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "UID"
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "Extra Data"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   4740
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "@"
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Percent"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   195
      Left            =   60
      TabIndex        =   11
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

'    If mMsg Then _
        sn41EZUpdate mMsg, , Text2.Text, , , IIf(Check1.Value = vbChecked, "percent::" & CStr(HScroll1.Value), "")

End Sub

Private Sub Command4_Click()
Dim szIcon As String
Dim sz As String

    If mToken = 0 Then _
        Exit Sub

    szIcon = Text3.Text
    If g_SafeLeftStr(szIcon, 2) = ".\" Then _
        szIcon = g_MakePath(App.Path) & g_SafeRightStr(szIcon, Len(szIcon) - 2)

    ' /* base request */
    sz = "notify?app-sig=" & App.ProductName & "&title=" & Text1.Text & "&text=" & Text2.Text

    ' /* add icon? */
    If Check2.Value = vbChecked Then
        sz = sz & "&icon=" & szIcon

    Else
        ' /* must specify a blank icon */
        sz = sz & "&icon="

    End If

    ' /* add value-percent? */
    If Check1.Value = vbChecked Then _
        sz = sz & "&value-percent=" & CStr(HScroll1.Value)

    ' /* colour */
    If Check3.Value = vbChecked Then _
        sz = sz & "&colour-background=" & Text4.Text

    ' /* uid */
    If Text6.Text <> "" Then _
        sz = sz & "&uid=" & Text6.Text

    ' /* free-form */
    If Text5.Text <> "" Then _
        sz = sz & "&" & Text5.Text

'                            IIf(Check2.Value = vbChecked, szIcon, ""), , , , , _
'                            IIf(Check1.Value = vbChecked, "percent::" & CStr(HScroll1.Value), "") & _
'                                IIf(Check3.Value = vbChecked, "#?col::" & Text4.Text, "") & _
'                                IIf(Text5.Text <> "", "#?" & Text5.Text, ""))

    mMsg = snDoRequest(sz)
    Debug.Print mMsg

End Sub

Private Sub Form_Load()
Dim hr As Long

    If Not snIsSnarlRunning() Then
        MsgBox "Snarl isn't running - launch Snarl, then run this demo.", vbExclamation Or vbOKOnly, App.Title
        Unload Me

    Else
        hr = snarl_register(App.ProductName, App.Title, App.Path & "\icon.png")
        If hr > 0 Then
            Me.Caption = "Registered with Snarl V" & CStr(snarl_version()) & " (" & Hex$(hr) & ")"
            mToken = hr
            
        Else
            Me.Caption = "Error registering with Snarl: " & Abs(hr)

        End If

        Text3.Text = ".\icon.png"
        Label4.Caption = HScroll1.Value & "%"

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim hr As Long

    hr = snarl_unregister(mToken)
    If hr = 0 Then
        Debug.Print "unregistered ok"

    Else
        Debug.Print "failed to unregister: " & Abs(hr)

    End If

End Sub

Private Sub HScroll1_Change()

    Label4.Caption = HScroll1.Value & "%"

End Sub

Private Sub HScroll1_Scroll()

    Label4.Caption = HScroll1.Value & "%"

End Sub

