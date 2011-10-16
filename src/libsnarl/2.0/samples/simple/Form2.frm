VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
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
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Use Win32 API"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   1500
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   555
      Left            =   300
      TabIndex        =   0
      Top             =   540
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents myApp As SnarlApp
Attribute myApp.VB_VarHelpID = -1

Private Sub Command1_Click()

    ' /* create actions */

Dim pa As Actions

    Set pa = New Actions
    pa.Add "Action 1", "act1"
    pa.Add "Forward", "fwd"
    pa.Add "Quit", "x"

    ' /* create notification */

Dim pn As Notification

    Set pn = New Notification
    With pn
        .Title = "Hello, world!"
        .Text = "This is a test..."
        .UID = "123456"
        .Actions = pa

    End With

    If Check1.Value = vbUnchecked Then _
        myApp.RemoteComputer = "127.0.0.1"

    ' /* go */
    
    myApp.Show pn

End Sub

Private Sub Command2_Click()

'    myApp.TidyUp

End Sub

Private Sub Form_Load()

    ' /* create app */

    Set myApp = New SnarlApp

    Debug.Print myApp.IsSnarlInstalled()


    With myApp
        .Signature = "test/libmsnarl2"
        .Title = "snarl.library 2 test"
        .Icon = .MakePath(App.Path) & "icon.png"
        .Hint = "Acme Products Present"
        .IsDaemon = True

    End With

    Me.Caption = "Snarl: " & myApp.IsSnarlRunning()

End Sub

Private Sub myApp_Activated()

    Me.Show

End Sub

Private Sub myApp_NotificationActionSelected(ByVal UID As String, ByVal Identifier As String)

    Debug.Print "## action '" & Identifier & "' from " & UID & " selected ##"

    If Identifier = "x" Then _
        myApp.Hide UID

End Sub

Private Sub myApp_NotificationInvoked(ByVal UID As String)

    Debug.Print "## " & UID & " invoked ##"

End Sub

Private Sub myApp_Quit()

    Unload Me

End Sub

