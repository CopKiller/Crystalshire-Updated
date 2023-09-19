Attribute VB_Name = "Player_TCP"
Public Function PlayerData(ByVal index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If index > MAX_PLAYERS Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerLevel(index)
    Buffer.WriteLong GetPlayerPOINTS(index)
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteLong GetPlayerClass(index)
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    PlayerData = Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Function

Public Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Public Sub SendPlayerData_Party(partynum As Long)
    Dim i As Long, x As Long
    ' loop through all the party members
    For i = 1 To Party(partynum).MemberCount
        For x = 1 To Party(partynum).MemberCount
            SendDataTo Party(partynum).Member(x), PlayerData(Party(partynum).Member(i))
        Next
    Next
End Sub

Public Sub SendPlayerVariables(ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerVariables
    For i = 1 To MAX_BYTE
        Buffer.WriteLong Player(index).Variable(i)
    Next
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendAttack(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTarget(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    Buffer.WriteLong TempPlayer(index).target
    Buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendHotbar(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong Player(index).Hotbar(i).Slot
        Buffer.WriteByte Player(index).Hotbar(i).sType
    Next
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerMove(ByVal index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    End If
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXY
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
    
    ' check if they're in a party
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
End Sub

Public Sub SendEXP(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendStats(ByVal index As Long)
    Dim i As Long
    Dim packet As String
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, i)
        Buffer.WriteLong GetPlayerInvItemValue(index, i)
        Buffer.WriteByte Player(index).Inv(i).Bound
    Next

    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    Buffer.WriteByte Player(index).Inv(invSlot).Bound
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong Player(index).Spell(i).Spell
        Buffer.WriteLong Player(index).Spell(i).Uses
    Next

    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    Buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim tradeTarget As Long
    Dim totalWorth As Long, multiplier As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    Buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    Buffer.WriteByte Status
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    Buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyUpdate(ByVal partynum As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partynum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partynum).Member(i)
    Next
    Buffer.WriteLong Party(partynum).MemberCount
    SendDataToParty partynum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyUpdateTo(ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long, partynum As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partynum = TempPlayer(index).inParty
    If partynum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partynum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partynum).Member(i)
        Next
        Buffer.WriteLong Party(partynum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendPartyVitals(ByVal partynum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong Player(index).Vital(i)
    Next
    SendDataToParty partynum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendDataToParty(ByVal partynum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Party(partynum).MemberCount
        If Party(partynum).Member(i) > 0 Then
            Call SendDataTo(Party(partynum).Member(i), Data)
        End If
    Next
End Sub

Public Sub SendBlood(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    Buffer.WriteLong shopNum
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Player(index).Bank(i).Num
        Buffer.WriteLong Player(index).Bank(i).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendWelcome(ByVal index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Public Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendStartTutorial(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStartTutorial
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
