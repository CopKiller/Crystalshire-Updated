Attribute VB_Name = "Account_Logic"
Option Explicit

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal charNum As Long)
    Dim f As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(index).Name)) = 0 Then

        spritecheck = False

        If charNum < 1 Or charNum > MAX_CHARS Then Exit Sub
        TempPlayer(index).charNum = charNum

        Player(index).Name = Name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum

        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).dir = DIR_DOWN
        Player(index).Map = START_MAP
        Player(index).x = START_X
        Player(index).y = START_Y
        Player(index).dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)

        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If

        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Spell(n).Spell = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).Slot = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).sType = 2    ' spells
                    End If
                End If
            Next
        End If

        ' Append name to file
        f = FreeFile
        Open App.Path & "\data\accounts\_charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub
