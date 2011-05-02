Attribute VB_Name = "mMain"
Option Explicit

Public Const SCHEME_1 = "1996"
Public Const SCHEME_2 = "2004"

'Public Const SM_CXSCREEN = 0
'Public Const SM_CYSCREEN = 1
'Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public gSettings As CConfFile

Public gTinyDefaults As BPackedData
