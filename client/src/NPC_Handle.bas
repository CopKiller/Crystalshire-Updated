Attribute VB_Name = "NPC_Handle"
Public Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS

        With MapNpc(i)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadLong
            .Vital(HP) = buffer.ReadLong
        End With

    Next

End Sub

Public Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long
    Dim Movement As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNpcNum = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    Movement = buffer.ReadLong

    With MapNpc(MapNpcNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir

            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
            Case DIR_UP_LEFT
                .yOffset = PIC_Y
                .xOffset = PIC_X
            Case DIR_UP_RIGHT
                .yOffset = PIC_Y
                .xOffset = PIC_X * -1
            Case DIR_DOWN_LEFT
                .yOffset = PIC_Y * -1
                .xOffset = PIC_X
            Case DIR_DOWN_RIGHT
                .yOffset = PIC_Y * -1
                .xOffset = PIC_X * -1
        End Select

    End With

End Sub

Public Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong

    With MapNpc(i)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Public Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = getTime
End Sub

Public Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    N = buffer.ReadLong

    With MapNpc(N)
        .Num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadLong
        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With

End Sub

Public Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNpcNum = buffer.ReadLong

    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = buffer.ReadLong
    Next

    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    N = buffer.ReadLong
    Call ClearMapNpc(N)
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' NPC EDITORES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Sub HandleNpcEditor()
    Dim i As Long

    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With

End Sub

Public Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    N = buffer.ReadLong
    NpcSize = LenB(Npc(N))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(Npc(N)), ByVal VarPtr(NpcData(0)), NpcSize
    buffer.Flush: Set buffer = Nothing
End Sub
