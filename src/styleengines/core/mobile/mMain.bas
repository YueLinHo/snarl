Attribute VB_Name = "mMain"
Option Explicit

Public gID As Long

Dim mList As BTagList

Public Sub g_AddThis(ByRef It As TGenericPhoneInstance)

    If (mList Is Nothing) Then _
        Set mList = new_BTagList()

    mList.Add It

End Sub

Public Sub g_RemoveThis(ByRef It As BTagItem)

    mList.Remove mList.IndexOf(It.Name)

End Sub
