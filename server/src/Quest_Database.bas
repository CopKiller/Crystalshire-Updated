Attribute VB_Name = "Quest_Database"
' **********
' ** Quests **
' **********
Public Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\quests\Quest" & QuestNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Quest(QuestNum)
    Close #f
End Sub

Public Sub SaveQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next

End Sub

Public Sub CheckQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If Not FileExist(App.Path & "\data\quests\Quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If
    Next

End Sub

Public Sub LoadQuests()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Quest(i)
        Close #f
    Next

End Sub

Public Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    Quest(index).name = vbNullString
End Sub

Public Sub ClearQuests()
    Dim i As Long

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next
End Sub


