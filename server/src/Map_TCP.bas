Attribute VB_Name = "Map_TCP"
Public Sub SendMap(ByVal index As Long, ByVal mapnum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    'Buffer.PreAllocate (UBound(MapCache(mapNum).Data) - LBound(MapCache(mapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(mapnum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    Buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapItemsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapnum, i).playerName
        Buffer.WriteLong MapItem(mapnum, i).Num
        Buffer.WriteLong MapItem(mapnum, i).Value
        Buffer.WriteLong MapItem(mapnum, i).x
        Buffer.WriteLong MapItem(mapnum, i).y
        If MapItem(mapnum, i).Bound Then
            Buffer.WriteLong 1
        Else
            Buffer.WriteLong 0
        End If
    Next

    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapItemsToAll(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(mapnum, i).playerName
        Buffer.WriteLong MapItem(mapnum, i).Num
        Buffer.WriteLong MapItem(mapnum, i).Value
        Buffer.WriteLong MapItem(mapnum, i).x
        Buffer.WriteLong MapItem(mapnum, i).y
        If MapItem(mapnum, i).Bound Then
            Buffer.WriteLong 1
        Else
            Buffer.WriteLong 0
        End If
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapNpcVitals(ByVal mapnum As Long, ByVal mapNpcNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcVitals
    Buffer.WriteLong mapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(mapnum).Npc(mapNpcNum).Vital(i)
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapNpcsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(mapnum).Npc(i).Num
        Buffer.WriteLong MapNpc(mapnum).Npc(i).x
        Buffer.WriteLong MapNpc(mapnum).Npc(i).y
        Buffer.WriteLong MapNpc(mapnum).Npc(i).dir
        Buffer.WriteLong MapNpc(mapnum).Npc(i).Vital(HP)
    Next

    SendDataTo index, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapNpcsToMap(ByVal mapnum As Long)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(mapnum).Npc(i).Num
        Buffer.WriteLong MapNpc(mapnum).Npc(i).x
        Buffer.WriteLong MapNpc(mapnum).Npc(i).y
        Buffer.WriteLong MapNpc(mapnum).Npc(i).dir
        Buffer.WriteLong MapNpc(mapnum).Npc(i).Vital(HP)
    Next

    SendDataToMap mapnum, Buffer.ToArray()
    
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
