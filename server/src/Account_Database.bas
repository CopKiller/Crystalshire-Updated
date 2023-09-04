Attribute VB_Name = "Account_Database"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' *************
' ** Account **
' *************

Sub SaveAccount(ByVal index As Long)
    Dim filename As String, i As Long, charHeader As String, f As Long

    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub

    ' the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Account(index).Login)) & "\account.bin"

    ' Save Player archive
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Account(index)
    Close #f
End Sub
Sub LoadAccount(ByVal index As Long, ByVal Name As String)
    Dim filename As String, i As Long, charHeader As String, f As Long

    If Trim$(Name) = vbNullString Then Exit Sub
    ' clear Account
    Call ClearAccount(index)

    ' the file
    filename = App.Path & "\data\accounts\" & SanitiseString(Name) & "\account.bin"
    
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Account(index)
    Close #f
End Sub

Sub ClearAccount(ByVal index As Long)
    Dim i As Long

    Call ZeroMemory(ByVal VarPtr(Account(index)), LenB(Account(index)))
    
    Account(index).Login = vbNullString
    Account(index).Password = vbNullString
    Account(index).Mail = vbNullString
    
    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

Function OldAccount_Exist(ByVal username As String) As Boolean
    Dim filename As String

    filename = App.Path & "\data\accounts\old\" & SanitiseString(username) & ".ini"
    If FileExist(filename) Then
        If LenB(Trim$(GetVar(filename, "ACCOUNT", "Name"))) > 0 Then
            OldAccount_Exist = True
        End If
    End If
End Function

Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim(Name)) & "\" & "account.bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * ACCOUNT_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & SanitiseString(Trim(Name)) & "\" & "account.bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String, ByVal Code As String)
    Dim i As Long

    ClearAccount index

    '//Create the file destination folder
    ChkDir App.Path & "\data\accounts\", Trim$(Name)

    Account(index).Login = Name
    Account(index).Password = Password
    Account(index).Mail = Code
    
    Call SaveAccount(index)

    For i = 1 To MAX_CHARS
        Call ClearPlayer(index)
        TempPlayer(index).charNum = i
        Call SavePlayer(index)
    Next i
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\_charlist.txt", App.Path & "\data\accounts\_chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\_chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\_chartemp.txt")
End Sub

Public Sub MergeAccount(ByVal index As Long, ByVal charNum As Long, ByVal oldAccount As String)
    Dim tempChar As PlayerRec, charHeader As String, filename As String, i As Long

    ' set the filename
    filename = App.Path & "\data\accounts\old\" & SanitiseString(oldAccount) & ".ini"
    charHeader = "ACCOUNT"

    ' load the old account shit
    With tempChar
        .Name = Trim$(GetVar(filename, charHeader, "Name"))
        .Sex = Val(GetVar(filename, charHeader, "Sex"))
        .Class = Val(GetVar(filename, charHeader, "Class"))
        .Sprite = Val(GetVar(filename, charHeader, "Sprite"))
        .Level = Val(GetVar(filename, charHeader, "Level"))
        .exp = Val(GetVar(filename, charHeader, "Exp"))
        .Access = Val(GetVar(filename, charHeader, "Access"))
        .PK = Val(GetVar(filename, charHeader, "PK"))

        ' Vitals
        For i = 1 To Vitals.Vital_Count - 1
            .Vital(i) = Val(GetVar(filename, charHeader, "Vital" & i))
        Next

        ' Stats
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = Val(GetVar(filename, charHeader, "Stat" & i))
        Next
        .POINTS = Val(GetVar(filename, charHeader, "Points"))

        ' Equipment
        For i = 1 To Equipment.Equipment_Count - 1
            .Equipment(i) = Val(GetVar(filename, charHeader, "Equipment" & i))
        Next

        ' Inventory
        For i = 1 To MAX_INV
            .Inv(i).Num = Val(GetVar(filename, charHeader, "InvNum" & i))
            .Inv(i).Value = Val(GetVar(filename, charHeader, "InvValue" & i))
            .Inv(i).Bound = Val(GetVar(filename, charHeader, "InvBound" & i))
        Next

        ' Spells
        For i = 1 To MAX_PLAYER_SPELLS
            .Spell(i).Spell = Val(GetVar(filename, charHeader, "Spell" & i))
            .Spell(i).Uses = Val(GetVar(filename, charHeader, "SpellUses" & i))
        Next

        ' Hotbar
        For i = 1 To MAX_HOTBAR
            .Hotbar(i).Slot = Val(GetVar(filename, charHeader, "HotbarSlot" & i))
            .Hotbar(i).sType = Val(GetVar(filename, charHeader, "HotbarType" & i))
        Next

        ' Position
        .Map = Val(GetVar(filename, charHeader, "Map"))
        .x = Val(GetVar(filename, charHeader, "X"))
        .y = Val(GetVar(filename, charHeader, "Y"))
        .dir = Val(GetVar(filename, charHeader, "Dir"))

        ' Tutorial
        .TutorialState = Val(GetVar(filename, charHeader, "TutorialState"))
    End With

    ' set the filename
    filename = App.Path & "\data\accounts\" & SanitiseString(Trim$(Player(index).Login)) & ".ini"
    charHeader = "CHAR" & charNum

    ' save it in the new account's character slot
    With tempChar
        PutVar filename, charHeader, "Name", Trim$(.Name)
        PutVar filename, charHeader, "Sex", Val(.Sex)
        PutVar filename, charHeader, "Class", Val(.Class)
        PutVar filename, charHeader, "Sprite", Val(.Sprite)
        PutVar filename, charHeader, "Level", Val(.Level)
        PutVar filename, charHeader, "exp", Val(.exp)
        PutVar filename, charHeader, "Access", Val(.Access)
        PutVar filename, charHeader, "PK", Val(.PK)

        ' Vitals
        For i = 1 To Vitals.Vital_Count - 1
            PutVar filename, charHeader, "Vital" & i, Val(.Vital(i))
        Next

        ' Stats
        For i = 1 To Stats.Stat_Count - 1
            PutVar filename, charHeader, "Stat" & i, Val(.Stat(i))
        Next
        PutVar filename, charHeader, "Points", Val(.POINTS)

        ' Equipment
        For i = 1 To Equipment.Equipment_Count - 1
            PutVar filename, charHeader, "Equipment" & i, Val(.Equipment(i))
        Next

        ' Inventory
        For i = 1 To MAX_INV
            PutVar filename, charHeader, "InvNum" & i, Val(.Inv(i).Num)
            PutVar filename, charHeader, "InvValue" & i, Val(.Inv(i).Value)
            PutVar filename, charHeader, "InvBound" & i, Val(.Inv(i).Bound)
        Next

        ' Spells
        For i = 1 To MAX_PLAYER_SPELLS
            PutVar filename, charHeader, "Spell" & i, Val(.Spell(i).Spell)
            PutVar filename, charHeader, "SpellUses" & i, Val(.Spell(i).Uses)
        Next

        ' Hotbar
        For i = 1 To MAX_HOTBAR
            PutVar filename, charHeader, "HotbarSlot" & i, Val(.Hotbar(i).Slot)
            PutVar filename, charHeader, "HotbarType" & i, Val(.Hotbar(i).sType)
        Next

        ' Position
        PutVar filename, charHeader, "Map", Val(.Map)
        PutVar filename, charHeader, "X", Val(.x)
        PutVar filename, charHeader, "Y", Val(.y)
        PutVar filename, charHeader, "Dir", Val(.dir)

        ' Tutorial
        PutVar filename, charHeader, "TutorialState", Val(.TutorialState)
    End With

    ' kill the old account - permanently
    Kill App.Path & "\data\accounts\old\" & SanitiseString(oldAccount) & ".ini"

    ' send to portal again
    SendPlayerChars index

    ' confirmation message
    AlertMsg index, DIALOGUE_MSG_MERGE, MENU_CHARS, False
End Sub
