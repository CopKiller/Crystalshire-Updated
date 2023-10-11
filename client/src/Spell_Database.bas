Attribute VB_Name = "Spell_Database"
Public Sub ClearSpell(ByVal Index As Long)
    Spell(Index) = EmptySpell
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "None."
End Sub

Public Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

