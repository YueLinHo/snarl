Attribute VB_Name = "mMain"
Option Explicit

Public gID As Long

Dim mList As BTagList

'Dim mInstance() As TToastyInstance
'Dim mInstances As Long

Public Sub g_AddThis(ByRef It As TToastyInstance)

    If (mList Is Nothing) Then _
        Set mList = new_BTagList()

    mList.Add It

'    mInstances = mInstances + 1
'    ReDim Preserve mInstance(mInstances)
'    Set mInstance(mInstances) = It

End Sub

Public Sub g_RemoveThis(ByRef It As BTagItem)

    mList.Remove mList.IndexOf(It.Name)

End Sub
