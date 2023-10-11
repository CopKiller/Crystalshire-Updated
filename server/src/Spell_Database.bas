Attribute VB_Name = "Spell_Database"
' ************
' ** Spells **
' ************
Public Sub SaveSpell(ByVal spellNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\spells\spells" & spellNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Spell(spellNum)
    Close #f
End Sub

Public Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Public Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist(App.Path & "\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Public Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Public Sub ClearSpell(ByVal index As Long)
    Spell(index) = EmptySpell
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1    'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Public Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub
