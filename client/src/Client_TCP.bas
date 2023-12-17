Attribute VB_Name = "Client_TCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit(ByVal IP As String, ByVal Port As Long)
    Set PlayerBuffer = Nothing
    Set PlayerBuffer = New clsBuffer
    ' connect
    frmMain.Socket.Close
    frmMain.Socket.RemoteHost = IP
    frmMain.Socket.RemotePort = Port
End Sub

Sub DestroyTCP()
    frmMain.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
    Dim buffer() As Byte
    Dim pLength As Long
    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    PlayerBuffer.WriteBytes buffer()

    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)

    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4

        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0

        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop

    PlayerBuffer.Trim

    DoEvents
End Sub

Public Function ConnectToServer() As Boolean
    Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Wait = getTime
    frmMain.Socket.Close
    frmMain.Socket.Connect
    SetStatus "Connecting to server."

    ' Wait until connected or 3 seconds have passed and report the server being down
    Do While (Not IsConnected) And (getTime <= Wait + 3000)
        DoEvents
    Loop

    ConnectToServer = IsConnected
    SetStatus vbNullString
End Function

Function IsConnected() As Boolean

    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

End Function

Sub SendData(ByRef Data() As Byte)
    Dim buffer As clsBuffer

    If IsConnected Then
        Set buffer = New clsBuffer
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
        frmMain.Socket.SendData buffer.ToArray()
    End If

End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************

Public Sub SendLogin(ByVal Name As String, ByVal password As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong clogin
    buffer.WriteString Name
    buffer.WriteString password
    buffer.WriteLong CLIENT_MAJOR
    buffer.WriteLong CLIENT_MINOR
    buffer.WriteLong CLIENT_REVISION
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
Public Sub SendNewAccount(ByVal AName As String, ByVal APass As String, ByVal ACode As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong CNewAccount
    buffer.WriteString AName
    buffer.WriteString APass
    buffer.WriteString ACode
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal sprite As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAddChar
    buffer.WriteString Name
    buffer.WriteLong Sex
    buffer.WriteLong ClassNum
    buffer.WriteLong sprite
    buffer.WriteLong CharNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUseChar
    buffer.WriteLong CharSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendDelChar(ByVal CharSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDelChar
    buffer.WriteLong CharSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SayMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal text As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString MsgTo
    buffer.WriteString text
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerMove()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong Player(MyIndex).Moving
    buffer.WriteLong Player(MyIndex).X
    buffer.WriteLong Player(MyIndex).Y
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerDir()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendMap()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    CanMoveNow = False
    
    buffer.WriteLong CMapData

    buffer.WriteString Trim$(Map.MapData.Name)
    buffer.WriteString Trim$(Map.MapData.Music)
    buffer.WriteByte Map.MapData.Moral
    buffer.WriteLong Map.MapData.Up
    buffer.WriteLong Map.MapData.Down
    buffer.WriteLong Map.MapData.Left
    buffer.WriteLong Map.MapData.Right
    buffer.WriteLong Map.MapData.BootMap
    buffer.WriteByte Map.MapData.BootX
    buffer.WriteByte Map.MapData.BootY
    buffer.WriteByte Map.MapData.MaxX
    buffer.WriteByte Map.MapData.MaxY
    buffer.WriteLong Map.MapData.Weather
    buffer.WriteLong Map.MapData.WeatherIntensity
    buffer.WriteLong Map.MapData.Fog
    buffer.WriteLong Map.MapData.FogSpeed
    buffer.WriteLong Map.MapData.FogOpacity
    buffer.WriteLong Map.MapData.Red
    buffer.WriteLong Map.MapData.Green
    buffer.WriteLong Map.MapData.Blue
    buffer.WriteLong Map.MapData.alpha
    buffer.WriteLong Map.MapData.BossNpc
    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map.MapData.Npc(i)
    Next

    For X = 0 To Map.MapData.MaxX
        For Y = 0 To Map.MapData.MaxY
            With Map.TileData.Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).tileSet
                    buffer.WriteByte .Autotile(i)
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteLong .Data4
                buffer.WriteLong .Data5
                buffer.WriteByte .DirBlock
            End With
        Next
    Next

    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString Name
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString Name
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub WarpTo(ByVal mapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong mapNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString Name
    buffer.WriteLong Access
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSetSprite
    buffer.WriteLong SpriteNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendKick(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendBan(ByVal Name As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString Name
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendBanList()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBanList
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendMapRespawn()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
    Dim buffer As clsBuffer

    If InBank Or InShop Then Exit Sub

    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).Num < 1 Or PlayerInv(invNum).Num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong invNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendWhosOnline()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendBanDestroy()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBanDestroy
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong oldSlot
    buffer.WriteLong newSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
    ' buffer it
    PlayerSwitchInvSlots oldSlot, newSlot
End Sub

Sub SendChangeSpellSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong oldSlot
    buffer.WriteLong newSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
    ' buffer it
    PlayerSwitchSpellSlots oldSlot, newSlot
End Sub

Sub GetPing()
    Dim buffer As clsBuffer
    PingStart = getTime
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendUnequip(ByVal eqNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong eqNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendRequestPlayerData()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPlayerData
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendRequestItems()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItems
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendTrainStat(ByVal statNum As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte statNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLevelUp
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong shopSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SellItem(ByVal invSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub WithdrawItem(ByVal BankSlot As Long, ByVal Amount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItem
    buffer.WriteLong BankSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub CloseBank()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CCloseBank
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
    InBank = False
End Sub

Public Sub ChangeBankSlots(ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CChangeBankSlots
    buffer.WriteLong oldSlot
    buffer.WriteLong newSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
    If X < 0 Or Y < 0 Or X > Map.MapData.MaxX Or Y > Map.MapData.MaxY Then Exit Sub
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong X
    buffer.WriteLong Y
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub AcceptTrade()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub DeclineTrade()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong Slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
    Dim buffer As clsBuffer, X As Long

    ' check if spell
    If Hotbar(Slot).sType = 1 Then ' Item
        For X = 1 To MAX_INV
            ' Is the item matching the hotbar
            If GetPlayerInvItemNum(MyIndex, X) = Hotbar(Slot).Slot Then
                SendUseItem X
                Exit Sub
            End If
        Next
        
        For X = 1 To Equipment.Equipment_Count - 1
            If Player(MyIndex).Equipment(X) = Hotbar(Slot).Slot Then
                SendUnequip X
                Exit Sub
            End If
        Next
        
        If Hotbar(Slot).Slot > 0 Then
            AddText "Você não tem este item!", 12
        End If
    ElseIf Hotbar(Slot).sType = 2 Then ' spell

        For X = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(X).Spell = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell X
                Exit Sub
            End If
        Next

        ' can't find the spell, exit out
        Exit Sub
    End If
End Sub

Public Sub SendMapReport()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub PlayerTarget(ByVal target As Long, ByVal TargetType As Long)
    Dim buffer As clsBuffer

    If myTargetType = TargetType And myTarget = target Then
        myTargetType = 0
        myTarget = 0
    Else
        myTarget = target
        myTargetType = TargetType
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CTarget
    buffer.WriteLong target
    buffer.WriteLong TargetType
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendTradeRequest(playerIndex As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    buffer.WriteLong playerIndex
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendAcceptTradeRequest()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendDeclineTradeRequest()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendPartyLeave()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendPartyRequest(Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    buffer.WriteLong Index
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendAcceptParty()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendDeclineParty()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendChatOption(ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CChatOption
    buffer.WriteLong Index
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendFinishTutorial()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CFinishTutorial
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendCloseShop()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CCloseShop
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendTarget(ByVal target As Long, ByVal TargetType As Long)
    Dim buffer As clsBuffer

    If myTargetType = TargetType And myTarget = target Then
        Exit Sub
    Else
        myTarget = target
        myTargetType = TargetType
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CTarget
    buffer.WriteLong target
    buffer.WriteLong TargetType
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
