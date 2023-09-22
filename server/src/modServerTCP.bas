Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    frmServer.Caption = GAME_NAME & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Public Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim TempData() As Byte

    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData buffer.ToArray()
    End If
End Sub

Public Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Public Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong color
    SendDataToAll buffer.ToArray
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteLong color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, buffer.ToArray
        End If
    Next
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong color
    SendDataTo index, buffer.ToArray
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteLong color
    SendDataToMap MapNum, buffer.ToArray
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal MessageNo As Long, Optional ByVal MenuReset As Long = 0, Optional ByVal kick As Boolean = True)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteLong MessageNo
    buffer.WriteLong MenuReset
    If kick Then buffer.WriteLong 1 Else buffer.WriteLong 0
    SendDataTo index, buffer.ToArray
    
    If kick Then
        DoEvents
        Call CloseSocket(index)
    End If
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partynum As Long, ByVal Msg As String, ByVal color As Byte)
Dim i As Long
    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partynum).Member(i) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partynum).Member(i)) And IsPlaying(Party(partynum).Member(i)) Then
                PlayerMsg Party(partynum).Member(i), Msg, color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long)
    Call AlertMsg(index, DIALOGUE_MSG_CONNECTION)
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If index <> 0 Then
        ' make sure they're not banned
        If Not isBanned_IP(GetPlayerIP(index)) Then
            If GetPlayerIP(index) <> "69.163.139.25" Then Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, DIALOGUE_MSG_BANNED)
        End If
        ' re-set the high index
        SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(index).buffer.WriteBytes buffer()
    
    If TempPlayer(index).buffer.Length >= 4 Then
        pLength = TempPlayer(index).buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).buffer.Length - 4
        If pLength <= TempPlayer(index).buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).buffer.ReadLong
            HandleData index, TempPlayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).buffer.Length >= 4 Then
            pLength = TempPlayer(index).buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
    End If

End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong MapNum
    buffer.WriteString Trim$(Map(MapNum).MapData.name)
    buffer.WriteString Trim$(Map(MapNum).MapData.Music)
    buffer.WriteByte Map(MapNum).MapData.Moral
    buffer.WriteLong Map(MapNum).MapData.Up
    buffer.WriteLong Map(MapNum).MapData.Down
    buffer.WriteLong Map(MapNum).MapData.left
    buffer.WriteLong Map(MapNum).MapData.Right
    buffer.WriteLong Map(MapNum).MapData.BootMap
    buffer.WriteByte Map(MapNum).MapData.BootX
    buffer.WriteByte Map(MapNum).MapData.BootY
    buffer.WriteByte Map(MapNum).MapData.MaxX
    buffer.WriteByte Map(MapNum).MapData.MaxY
    
    buffer.WriteLong Map(MapNum).MapData.Weather
    buffer.WriteLong Map(MapNum).MapData.WeatherIntensity
    
    buffer.WriteLong Map(MapNum).MapData.Fog
    buffer.WriteLong Map(MapNum).MapData.FogSpeed
    buffer.WriteLong Map(MapNum).MapData.FogOpacity
    
    buffer.WriteLong Map(MapNum).MapData.Red
    buffer.WriteLong Map(MapNum).MapData.Green
    buffer.WriteLong Map(MapNum).MapData.Blue
    buffer.WriteLong Map(MapNum).MapData.Alpha
    
    buffer.WriteLong Map(MapNum).MapData.BossNpc
    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map(MapNum).MapData.Npc(i)
    Next
    
    For x = 0 To Map(MapNum).MapData.MaxX
        For y = 0 To Map(MapNum).MapData.MaxY
            With Map(MapNum).TileData.Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).x
                    buffer.WriteLong .Layer(i).y
                    buffer.WriteLong .Layer(i).Tileset
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

    MapCache(MapNum).Data = buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    buffer.Flush: Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong index
    SendDataToMapBut index, MapNum, buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SClassesData
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendDoorAnimation(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SDoorAnimation
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap MapNum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendActionMsg(ByVal MapNum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString message
    buffer.WriteLong color
    buffer.WriteLong MsgType
    buffer.WriteLong x
    buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendClearSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SayMsg_Map(ByVal MapNum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Map] "
    buffer.WriteLong saycolour
    
    SendDataToMap MapNum, buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Global] "
    buffer.WriteLong saycolour
    
    SendDataToAll buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte Value
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendMapKeyToMap(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte Value
    SendDataToMap MapNum, buffer.ToArray()
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendLoginOk(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    buffer.WriteLong index
    buffer.WriteLong Player_HighIndex
    SendDataTo index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendInGame(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendHighIndex()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    SendDataToAll buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnItem
    buffer.WriteLong index
    buffer.WriteString MapItem(MapNum, index).playerName
    buffer.WriteLong MapItem(MapNum, index).Num
    buffer.WriteLong MapItem(MapNum, index).Value
    buffer.WriteLong MapItem(MapNum, index).x
    buffer.WriteLong MapItem(MapNum, index).y
    If MapItem(MapNum, index).Bound Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    SendDataToMap MapNum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendChatUpdate(ByVal index As Long, ByVal npcNum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SChatUpdate
    buffer.WriteLong npcNum
    buffer.WriteString mT
    buffer.WriteString o1
    buffer.WriteString o2
    buffer.WriteString o3
    buffer.WriteString o4
    SendDataTo index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendNpcDeath(ByVal MapNum As Long, ByVal mapNpcNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDead
    buffer.WriteLong mapNpcNum
    SendDataToMap MapNum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendChatBubble(ByVal MapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal colour As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    buffer.WriteLong target
    buffer.WriteLong targetType
    buffer.WriteString message
    buffer.WriteLong colour
    SendDataToMap MapNum, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Function SanitiseString(ByVal theString As String) As String
    Dim i As Long, tmpString As String
    tmpString = vbNullString
    If Len(theString) <= 0 Then Exit Function
    For i = 1 To Len(theString)
        Select Case Mid$(theString, i, 1)
            Case "*"
                tmpString = tmpString + "[s]"
            Case ":"
                tmpString = tmpString + "[c]"
            Case Else
                tmpString = tmpString + Mid$(theString, i, 1)
        End Select
    Next
    SanitiseString = tmpString
End Function

Public Sub SendCancelAnimation(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCancelAnimation
    buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendCheckForMap(index As Long, MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SCheckForMap
    buffer.WriteLong MapNum
    buffer.WriteLong MapCRC32(MapNum).MapDataCRC
    buffer.WriteLong MapCRC32(MapNum).MapTileCRC
    
    SendDataTo index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendEvent(index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SEvent
    If TempPlayer(index).inEvent Then
        buffer.WriteLong 1
    Else
        buffer.WriteLong 0
    End If
    buffer.WriteLong TempPlayer(index).eventNum
    buffer.WriteLong TempPlayer(index).pageNum
    buffer.WriteLong TempPlayer(index).commandNum
    
    SendDataTo index, buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
