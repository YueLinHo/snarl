VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "THandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const SOS_BUILD = 1

Dim mDebugMode As Boolean

Public ErrN As String
Public ErrD As String
Public ErrS As String

'Public Obj As SOS

Implements BDoForEach

Private Function BDoForEach_DoForEach(Item As melon4.BObject, Optional ByVal Arg2 As Long = 0&) As Boolean

    Select Case Arg2
    
    Case SOS_COLD_BOOT
        ' /* cold boot (prompt) */
    
    Case SOS_WARM_BOOT
        ' /* warm boot (prompt) */
        SOS_printline "**** SOS System Error " & CStr(ErrN) & "!"
        SOS_printline ErrD & " in " & "[" & ErrS & "]"
        SOS_printline ""
        SOS_printline "type <h> for help, <c> to continue or <x> to exit"
        SOS_printline ""
        SOS_printline ""

    Case SOS_ENDING


    Case SOS_PROCESS

        Select Case Item.Name

        Case "/"
            SOS_printline "Inspired by the Sanity Operating System developed by Dierk (Chaos) Ohlerich"
            SOS_printline ""

        Case "a"
            SOS_printline "SOS Debugger 1.0 (Build " & CStr(SOS_BUILD) & ")"
            SOS_printline "(c) 1992-2014 k23 productions, The KKI"
            SOS_printline ""
            SOS_printline "System Information:"
            SOS_printline "  Windows " & g_GetOSVersionString(True)
            SOS_printline "  Physical RAM: " & Format$(g_GetPhysMem(True), "#,#") & " bytes"
            SOS_printline "  Attached to: #" & GetCurrentProcessId() & " (" & App.ProductName & "; " & App.FileDescription & ")"
            SOS_printline ""
            SOS_printline ""

        Case "h", "?"
            SOS_printline ""
            SOS_printline "SOS Help Page"
            SOS_printline "-------------"
            SOS_printline "c - continue application"
            SOS_printline "x - exit application"
            SOS_printline ""

        Case "x"
            ExitProcess 0

        Case "c"
            SOS_quit

'        Case ""
'            ' /* special case */
'            Debug.Print "CR ONLY"
'            Exit Function

        Case Else
            Debug.Print Item.Name
            SOS_warmboot

           

        End Select

    End Select

End Function
