Attribute VB_Name = "Player_TCP"
Public Function PlayerData(ByVal Index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

    If Index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong Index
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerLevel(Index)
    buffer.WriteLong GetPlayerPOINTS(Index)
    buffer.WriteLong GetPlayerSprite(Index)
    buffer.WriteLong GetPlayerMap(Index)
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteLong GetPlayerClass(Index)
    
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    For i = 1 To MAX_PLAYER_MISSIONS
        buffer.WriteLong Player(Index).Mission(i).ID
        buffer.WriteLong Player(Index).Mission(i).Count
    Next i
    
    For i = 1 To MAX_MISSIONS
        buffer.WriteLong Player(Index).CompletedMission(i)
    Next i
    
    PlayerData = buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Function

Public Sub SendPlayerData(ByVal Index As Long)
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Public Sub SendPlayerMission(ByVal Index As Long, ByVal MissionIndex As Byte)
    Dim buffer As clsBuffer, i As Long

    If Index <= 0 Or Index > MAX_PLAYERS Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMission
    buffer.WriteLong Index
    buffer.WriteLong MissionIndex
    buffer.WriteLong Player(Index).Mission(MissionIndex).Count
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub OfferMission(ByVal Index As Long, ByVal MissionID As Long)
    Dim i As Long
    Dim FreeSlot As Boolean
    Dim buffer As clsBuffer
    Dim MissionSlot As Long
    Dim CompletedPreviousQuest As Boolean
    Dim PreviousMissionID As Long
    
    FreeSlot = False
    CompletedPreviousQuest = False
    
    If TempPlayer(Index).MissionRequest <> 0 Then Exit Sub
    
    If Mission(MissionID).PreviousMissionComplete > 0 Then
        For i = 1 To MAX_MISSIONS
            If Player(Index).CompletedMission(i) = Mission(MissionID).PreviousMissionComplete Then
                CompletedPreviousQuest = True
            End If
        Next i
    End If
    
    If Mission(MissionID).PreviousMissionComplete > 0 And CompletedPreviousQuest = False Then
        PreviousMissionID = Mission(MissionID).PreviousMissionComplete
        Call PlayerMsg(Index, "You do not meet the requirements for this quest.", Yellow)
        Call PlayerMsg(Index, "You are required to complete the quest: " & Mission(PreviousMissionID).Name, Yellow)
        Exit Sub
    End If
    
    For i = 1 To MAX_PLAYER_MISSIONS
        If FreeSlot = False Then
            If Player(Index).Mission(i).ID = 0 Then
                FreeSlot = True
            End If
        End If
    Next i
    
    If FreeSlot = False Then
        Call PlayerMsg(Index, "Your quest log is full!", BrightRed)
        Exit Sub
    End If
    
    If FreeSlot = True Then
        TempPlayer(Index).MissionRequest = MissionID
        Set buffer = New clsBuffer
        buffer.WriteLong SOfferMission
        buffer.WriteLong MissionID
        SendDataTo Index, buffer.ToArray()
        buffer.Flush: Set buffer = Nothing
    End If
End Sub

Public Sub CompleteMission(ByVal Index As Long, ByVal MissionSlot As Long)
    Dim i As Long
    Dim MissionID As Long
    Dim ItemNum As Long
    Dim Count As Long
    Dim n As Long
    
    Count = 0
    MissionID = Player(Index).Mission(MissionSlot).ID
    Call PlayerMsg(Index, Mission(MissionID).Completed, Yellow)
    
    ' Zerando informações sobre missões ativas
    Player(Index).Mission(MissionSlot).ID = 0
    Player(Index).Mission(MissionSlot).Count = 0
    
    For i = MissionSlot To MAX_PLAYER_MISSIONS
        If i < MAX_PLAYER_MISSIONS Then
            Player(Index).Mission(i).ID = Player(Index).Mission(i + 1).ID
            Player(Index).Mission(i).Count = Player(Index).Mission(i + 1).Count
        End If
    Next
    
    For i = 1 To MAX_MISSIONS
        If Mission(MissionID).Repeatable = 1 Then
            If Player(Index).CompletedMission(i) = MissionID Or Player(Index).CompletedMission(i) = 0 Then
                Player(Index).CompletedMission(i) = MissionID
                Call GivePlayerEXP(Index, Mission(MissionID).RewardExperience)
                
                For n = 1 To 5
                    If Mission(MissionID).RewardItem(n).ItemNum > 1 Then
                        Call GiveInvItem(Index, Mission(MissionID).RewardItem(n).ItemNum, Mission(MissionID).RewardItem(n).ItemAmount, True)
                    End If
                Next n
            
                Call SendPlayerData(Index)
                Exit Sub
            End If
        Else ' If Mission(MissionID).Repeatable = 1 Then
            If Player(Index).CompletedMission(i) = 0 Then
                Player(Index).CompletedMission(i) = MissionID
                Call GivePlayerEXP(Index, Mission(MissionID).RewardExperience)
                
                If Mission(MissionID).Type = QUEST_TYPE_COLLECT Then
                    ItemNum = Mission(MissionID).CollectItem
                    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem Index, Mission(MissionID).CollectItem, Mission(MissionID).CollectItemAmount
                    Else
                        For n = 1 To Mission(MissionID).CollectItemAmount
                            TakeInvItem Index, Mission(MissionID).CollectItem, 1
                        Next n
                    End If
                End If
            
                For n = 1 To 5
                    If Mission(MissionID).RewardItem(n).ItemNum > 1 Then
                        Call GiveInvItem(Index, Mission(MissionID).RewardItem(n).ItemNum, Mission(MissionID).RewardItem(n).ItemAmount, True)
                    End If
                Next n
                Call SendPlayerData(Index)
                Exit Sub
            End If
        End If
    Next i
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

Public Sub SendPlayerVariables(ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerVariables
    For i = 1 To MAX_BYTE
        buffer.WriteLong Player(Index).Variable(i)
    Next
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendAttack(ByVal Index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong Index
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendTarget(ByVal Index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    buffer.WriteLong TempPlayer(Index).Target
    buffer.WriteLong TempPlayer(Index).targetType
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendHotbar(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    For i = 1 To MAX_HOTBAR
        buffer.WriteLong Player(Index).Hotbar(i).Slot
        buffer.WriteByte Player(Index).Hotbar(i).sType
    Next
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    End If
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerXY(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerXYToMap(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXYMap
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
    
    ' check if they're in a party
    If TempPlayer(Index).inParty > 0 Then SendPartyVitals TempPlayer(Index).inParty, Index
End Sub

Public Sub SendEXP(ByVal Index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong GetPlayerExp(Index)
    buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendStats(ByVal Index As Long)
    Dim i As Long
    Dim packet As String
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(Index, i)
        buffer.WriteLong GetPlayerInvItemValue(Index, i)
        buffer.WriteByte Player(Index).Inv(i).Bound
    Next

    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendInventoryUpdate(ByVal Index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteLong invSlot
    buffer.WriteLong GetPlayerInvItemNum(Index, invSlot)
    buffer.WriteLong GetPlayerInvItemValue(Index, invSlot)
    buffer.WriteByte Player(Index).Inv(invSlot).Bound
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerWornEq
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    buffer.WriteLong GetPlayerEquipment(Index, Pants)
    buffer.WriteLong GetPlayerEquipment(Index, Feet)
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerSpells(ByVal Index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong Player(Index).Spell(i).Spell
        buffer.WriteLong Player(Index).Spell(i).Uses
    Next

    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteString Trim$(Player(TradeRequest).Name)
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendTrade(ByVal Index As Long, ByVal tradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    buffer.WriteLong tradeTarget
    buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendCloseTrade(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendTradeUpdate(ByVal Index As Long, ByVal dataType As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim tradeTarget As Long
    Dim totalWorth As Long, multiplier As Long
    
    tradeTarget = TempPlayer(Index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
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
    buffer.WriteLong totalWorth
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    buffer.WriteByte Status
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPartyInvite(ByVal Index As Long, ByVal targetPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    buffer.WriteString Trim$(Player(targetPlayer).Name)
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPartyUpdate(ByVal partynum As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    buffer.WriteByte 1
    buffer.WriteLong Party(partynum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(partynum).Member(i)
    Next
    buffer.WriteLong Party(partynum).MemberCount
    
    SendDataToParty partynum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPartyUpdateTo(ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long, partynum As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partynum = TempPlayer(Index).inParty
    If partynum > 0 Then
        ' send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(partynum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(partynum).Member(i)
        Next
        buffer.WriteLong Party(partynum).MemberCount
    Else
        ' send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPartyVitals(ByVal partynum As Long, ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(Index, i)
        buffer.WriteLong Player(Index).Vital(i)
    Next
    
    SendDataToParty partynum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
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
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapnum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendStunned(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCooldown
    buffer.WriteLong Slot
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub ResetShopAction(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataToAll buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendOpenShop(ByVal Index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong shopNum
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendBank(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Player(Index).Bank(i).Num
        buffer.WriteLong Player(Index).Bank(i).Value
    Next
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendWelcome(ByVal Index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(Index)
End Sub

Public Sub SendPlayerSound(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendStartTutorial(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStartTutorial
    
    SendDataTo Index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendProjectile(ByVal mapnum As Long, ByVal ProjectileSlot As Long, Optional ByVal IsDirectional As Boolean = False)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SProjectileAttack
    buffer.WriteLong ProjectileSlot
    buffer.WriteLong MapProjectile_HighIndex
    With MapProjectile(ProjectileSlot)
        buffer.WriteLong .Owner
        buffer.WriteLong .OwnerType
        buffer.WriteLong .direction
        buffer.WriteLong .Graphic
        buffer.WriteByte Spell(.spellNum).IsAoE
        buffer.WriteLong .Rotate
        buffer.WriteLong .RotateSpeed
        buffer.WriteLong .Speed
        buffer.WriteLong .Duration
        buffer.WriteLong .x
        buffer.WriteLong .y
        buffer.WriteLong .tX
        buffer.WriteLong .tY
        buffer.WriteByte IsDirectional
        For i = 1 To 4
            buffer.WriteLong .ProjectileOffset(i).x
            buffer.WriteLong .ProjectileOffset(i).y
        Next
    End With
    
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub
