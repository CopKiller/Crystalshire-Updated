Attribute VB_Name = "Player_Database"
' *************
' ** Players **
' *************
Public Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next
End Sub

Public Sub SavePlayer(ByVal index As Long)
    Dim filename As String, i As Long, charHeader As String, f As Long

    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    If TempPlayer(index).charNum <= 0 Or TempPlayer(index).charNum > MAX_CHARS Then Exit Sub

    ' the file
    filename = App.Path & "\data\accounts\" & Trim$(Account(index).Login) & "\CharNum_" & TempPlayer(index).charNum & ".bin"
    ' Save Player archive
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Player(index)
    Close #f
End Sub

Public Sub LoadPlayer(ByVal index As Long, ByVal charNum As Long)
    Dim filename As String, i As Long, charHeader As String, f As Long

    '//Verify player have account
    If Trim$(Account(index).Login) = vbNullString Then Exit Sub
    ' clear player
    Call ClearPlayer(index)

    ' the file
    filename = App.Path & "\data\accounts\" & Trim$(Account(index).Login) & "\CharNum_" & charNum & ".bin"

    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Player(index)
    Close #f
End Sub

Public Sub DeleteCharacter(Login As String, charNum As Long)
    Dim filename As String, charHeader As String, i As Long

    Login = Trim$(Login)
    If Login = vbNullString Then Exit Sub

    ' the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Login) & ".ini"

    ' exit out early if invalid char
    If charNum < 1 Or charNum > MAX_CHARS Then Exit Sub

    ' the char header
    charHeader = "CHAR" & charNum

    ' character
    PutVar filename, charHeader, "Name", vbNullString
    PutVar filename, charHeader, "Sex", 0
    PutVar filename, charHeader, "Class", 0
    PutVar filename, charHeader, "Sprite", 0
    PutVar filename, charHeader, "Level", 0
    PutVar filename, charHeader, "exp", 0
    PutVar filename, charHeader, "Access", 0
    PutVar filename, charHeader, "PK", 0

    ' Vitals
    For i = 1 To Vitals.Vital_Count - 1
        PutVar filename, charHeader, "Vital" & i, 0
    Next

    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        PutVar filename, charHeader, "Stat" & i, 0
    Next
    PutVar filename, charHeader, "Points", 0

    ' Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        PutVar filename, charHeader, "Equipment" & i, 0
    Next

    ' Inventory
    For i = 1 To MAX_INV
        PutVar filename, charHeader, "InvNum" & i, 0
        PutVar filename, charHeader, "InvValue" & i, 0
        PutVar filename, charHeader, "InvBound" & i, 0
    Next

    ' Spells
    For i = 1 To MAX_PLAYER_SPELLS
        PutVar filename, charHeader, "Spell" & i, 0
        PutVar filename, charHeader, "SpellUses" & i, 0
    Next

    ' Hotbar
    For i = 1 To MAX_HOTBAR
        PutVar filename, charHeader, "HotbarSlot" & i, 0
        PutVar filename, charHeader, "HotbarType" & i, 0
    Next

    ' Position
    PutVar filename, charHeader, "Map", 0
    PutVar filename, charHeader, "X", 0
    PutVar filename, charHeader, "Y", 0
    PutVar filename, charHeader, "Dir", 0

    ' Tutorial
    PutVar filename, charHeader, "TutorialState", 0

    ' Bank
    For i = 1 To MAX_BANK
        PutVar filename, charHeader, "BankNum" & i, 0
        PutVar filename, charHeader, "BankValue" & i, 0
        PutVar filename, charHeader, "BankBound" & i, 0
    Next
End Sub

Public Sub ClearPlayer(ByVal index As Long)
    Dim i As Long

    TempPlayer(index) = EmptyTempPlayer
    Set TempPlayer(index).Buffer = New clsBuffer

    Player(index) = EmptyPlayer
    Player(index).Name = vbNullString
    Player(index).Class = 1

    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

Public Sub ClearChar(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Public Sub LoadClasses()
    Dim filename As String
    Dim i As Long, N As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")

        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For N = 0 To UBound(tmpArray)
            Class(i).MaleSprite(N) = Val(tmpArray(N))
        Next

        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For N = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(N) = Val(tmpArray(N))
        Next

        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))

        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)

        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If

        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)

        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Public Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long

    filename = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Public Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.Path & "\data\classes.ini"

    If Not FileExist(filename) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Public Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Class(i) = EmptyClass
        Class(i).Name = vbNullString
    Next

End Sub
