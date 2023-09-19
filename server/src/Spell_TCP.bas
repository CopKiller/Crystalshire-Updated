Attribute VB_Name = "Spell_TCP"
Public Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellNum
    Buffer.WriteBytes SpellData
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If

    Next

End Sub

Public Sub SendUpdateSpellToAll(ByVal spellNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong spellNum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
