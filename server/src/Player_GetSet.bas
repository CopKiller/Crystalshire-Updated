Attribute VB_Name = "Player_GetSet"
' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Public Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Account(index).Login)
End Function

Public Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Account(index).Login = Login
End Sub

Public Function GetPlayerName(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(index).Name)
End Function

Public Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Name = Name
End Sub

Public Function GetPlayerClass(ByVal index As Long) As Long
    If index <= 0 Or index > Player_HighIndex Then Exit Function
    GetPlayerClass = Player(index).Class
End Function

Public Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Public Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Public Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
    Case HP
        With Class(ClassNum)
            GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
        End With
    Case MP
        With Class(ClassNum)
            GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
        End With
    End Select
End Function

Public Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Public Function GetPlayerSprite(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(index).Sprite
End Function

Public Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Public Function GetPlayerLevel(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(index).Level
End Function

Public Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    SetPlayerLevel = False
    If Level > MAX_LEVELS Then
        Player(index).Level = MAX_LEVELS
        Exit Function
    End If
    Player(index).Level = Level
    SetPlayerLevel = True
End Function

Public Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = 100 + (((GetPlayerLevel(index) ^ 2) * 10) * 2)
End Function

Public Function GetPlayerExp(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(index).exp
End Function

Public Sub SetPlayerExp(ByVal index As Long, ByVal exp As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).exp = exp
End Sub

Public Function GetPlayerAccess(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(index).Access
End Function

Public Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Public Function GetPlayerPK(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(index).PK
End Function

Public Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Public Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(index).Vital(Vital)
End Function

Public Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(index).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Player(index).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Player(index).Vital(Vital) = 0
    End If

End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    If index > MAX_PLAYERS Then Exit Function
    
    x = Player(index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Equipment(i) > 0 Then
            If Item(Player(index).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + Item(Player(index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    If index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = Player(index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Player(index).Stat(Stat) = Value
End Sub

Public Function GetPlayerPOINTS(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(index).POINTS
End Function

Public Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    If POINTS <= 0 Then POINTS = 0
    Player(index).POINTS = POINTS
End Sub

Public Function GetPlayerMap(ByVal index As Long) As Long

    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(index).Map
End Function

Public Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)

    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Map = mapnum
    End If

End Sub

Public Function GetPlayerX(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(index).x
End Function

Public Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).x = x
End Sub

Public Function GetPlayerY(ByVal index As Long) As Long
    If index <= 0 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(index).y
End Function

Public Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    Player(index).y = y
End Sub

Public Function GetPlayerDir(ByVal index As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(index).Dir
End Function

Public Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Public Function GetPlayerIP(ByVal index As Long) As String

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Public Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    If index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Inv(invSlot).Num
End Function

Public Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    Player(index).Inv(invSlot).Num = ItemNum
End Sub

Public Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Player(index).Inv(invSlot).Value
End Function

Public Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(invSlot).Value = ItemValue
End Sub

Public Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long

    If index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    
    GetPlayerEquipment = Player(index).Equipment(EquipmentSlot)
End Function

Public Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    Player(index).Equipment(EquipmentSlot) = invNum
End Sub

Public Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemNum = Player(index).Bank(BankSlot).Num
End Function

Public Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    If BankSlot = 0 Then Exit Sub
    Player(index).Bank(BankSlot).Num = ItemNum
End Sub

Public Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot = 0 Then Exit Function
    GetPlayerBankItemValue = Player(index).Bank(BankSlot).Value
End Function

Public Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    If BankSlot = 0 Then Exit Sub
    Player(index).Bank(BankSlot).Value = ItemValue
End Sub

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(index, Armor)
    Helm = GetPlayerEquipment(index, Helmet)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
    End If

End Function
