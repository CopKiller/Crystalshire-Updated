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
    Dim Buffer As clsBuffer
    Dim TempData() As Byte

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
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

Sub SendDataToMap(ByVal mapnum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapnum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToAll Buffer.ToArray
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataTo index, Buffer.ToArray
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToMap mapnum, Buffer.ToArray
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal MessageNo As Long, Optional ByVal MenuReset As Long = 0, Optional ByVal kick As Boolean = True)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteLong MessageNo
    Buffer.WriteLong MenuReset
    If kick Then Buffer.WriteLong 1 Else Buffer.WriteLong 0
    SendDataTo index, Buffer.ToArray
    
    If kick Then
        DoEvents
        Call CloseSocket(index)
    End If
    
    Buffer.Flush: Set Buffer = Nothing
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
        Call SetHighIndex
        Call SendHighIndex
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer() As Byte
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
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        
        Call ClearPlayer(index)
        
        ' Set The High Index
        Call SetHighIndex
        Call SendHighIndex
    End If

End Sub

Public Sub MapCache_Create(ByVal mapnum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong mapnum
    Buffer.WriteString Trim$(Map(mapnum).MapData.Name)
    Buffer.WriteString Trim$(Map(mapnum).MapData.Music)
    Buffer.WriteByte Map(mapnum).MapData.Moral
    Buffer.WriteLong Map(mapnum).MapData.Up
    Buffer.WriteLong Map(mapnum).MapData.Down
    Buffer.WriteLong Map(mapnum).MapData.left
    Buffer.WriteLong Map(mapnum).MapData.Right
    Buffer.WriteLong Map(mapnum).MapData.BootMap
    Buffer.WriteByte Map(mapnum).MapData.BootX
    Buffer.WriteByte Map(mapnum).MapData.BootY
    Buffer.WriteByte Map(mapnum).MapData.MaxX
    Buffer.WriteByte Map(mapnum).MapData.MaxY
    
    Buffer.WriteLong Map(mapnum).MapData.Weather
    Buffer.WriteLong Map(mapnum).MapData.WeatherIntensity
    
    Buffer.WriteLong Map(mapnum).MapData.Fog
    Buffer.WriteLong Map(mapnum).MapData.FogSpeed
    Buffer.WriteLong Map(mapnum).MapData.FogOpacity
    
    Buffer.WriteLong Map(mapnum).MapData.Red
    Buffer.WriteLong Map(mapnum).MapData.Green
    Buffer.WriteLong Map(mapnum).MapData.Blue
    Buffer.WriteLong Map(mapnum).MapData.Alpha
    
    Buffer.WriteLong Map(mapnum).MapData.BossNpc
    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(mapnum).MapData.Npc(i)
    Next
    
    For x = 0 To Map(mapnum).MapData.MaxX
        For y = 0 To Map(mapnum).MapData.MaxY
            With Map(mapnum).TileData.Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).x
                    Buffer.WriteLong .Layer(i).y
                    Buffer.WriteLong .Layer(i).Tileset
                    Buffer.WriteByte .Autotile(i)
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteLong .Data4
                Buffer.WriteLong .Data5
                Buffer.WriteByte .DirBlock
            End With
        Next
    Next
    
    'zlib
    SetStatus "Map " & mapnum & " size before compression = " & Buffer.Count
    Buffer.CompressBuffer
    SetStatus "Map " & mapnum & " size after compression = " & Buffer.Count
    MapCache(mapnum).Data = Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim N As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                s = s & GetPlayerName(i) & ", "
                N = N + 1
            End If
        End If

    Next

    If N = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & N & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

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
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong index
    
    SendDataToMapBut index, mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, N As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        N = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        N = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    
    SendDataToAllBut index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendDoorAnimation(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendActionMsg(ByVal mapnum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString message
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap mapnum, Buffer.ToArray()
    End If
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendClearSpellBuffer(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SayMsg_Map(ByVal mapnum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapKeyToMap(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapKey
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte Value
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendLoginOk(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendInGame(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendHighIndex()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    Buffer.WriteLong Player_HighIndex
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendSpawnItemToMap(ByVal mapnum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong index
    Buffer.WriteString MapItem(mapnum, index).playerName
    Buffer.WriteLong MapItem(mapnum, index).Num
    Buffer.WriteLong MapItem(mapnum, index).Value
    Buffer.WriteLong MapItem(mapnum, index).x
    Buffer.WriteLong MapItem(mapnum, index).y
    If MapItem(mapnum, index).Bound Then
        Buffer.WriteLong 1
    Else
        Buffer.WriteLong 0
    End If
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendChatUpdate(ByVal index As Long, ByVal npcNum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatUpdate
    Buffer.WriteLong npcNum
    Buffer.WriteString mT
    Buffer.WriteString o1
    Buffer.WriteString o2
    Buffer.WriteString o3
    Buffer.WriteString o4
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendNpcDeath(ByVal mapnum As Long, ByVal mapNpcNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDead
    Buffer.WriteLong mapNpcNum
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendChatBubble(ByVal mapnum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal colour As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    Buffer.WriteLong target
    Buffer.WriteLong targetType
    Buffer.WriteString message
    Buffer.WriteLong colour
    
    SendDataToMap mapnum, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
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
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCancelAnimation
    Buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendCheckForMap(index As Long, mapnum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong mapnum
    Buffer.WriteLong MapCRC32(mapnum).MapDataCRC
    Buffer.WriteLong MapCRC32(mapnum).MapTileCRC
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
