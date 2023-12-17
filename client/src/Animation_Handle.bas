Attribute VB_Name = "Animation_Handle"
Public Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, X As Long, Y As Long, isCasting As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    AnimationIndex = AnimationIndex + 1

    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1

    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockindex = buffer.ReadLong
        .isCasting = buffer.ReadByte
        .Used(0) = True
        .Used(1) = True
    End With

    buffer.Flush: Set buffer = Nothing

    ' play the sound if we've got one
    With AnimInstance(AnimationIndex)

        If .LockType = 0 Then
            X = AnimInstance(AnimationIndex).X
            Y = AnimInstance(AnimationIndex).Y
        ElseIf .LockType = TARGET_TYPE_PLAYER Then
            X = GetPlayerX(.lockindex)
            Y = GetPlayerY(.lockindex)
        ElseIf .LockType = TARGET_TYPE_NPC Then
            X = MapNpc(.lockindex).X
            Y = MapNpc(.lockindex).Y
        End If

    End With

    PlayMapSound X, Y, SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation
End Sub

Public Sub HandleCancelAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim theIndex As Long, buffer As clsBuffer, i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    theIndex = buffer.ReadLong
    buffer.Flush: Set buffer = Nothing
    ' find the casting animation
    For i = 1 To MAX_BYTE
        If AnimInstance(i).LockType = TARGET_TYPE_PLAYER Then
            If AnimInstance(i).lockindex = theIndex Then
                If AnimInstance(i).isCasting = 1 Then
                    ' clear it
                    ClearAnimInstance i
                End If
            End If
        End If
    Next
End Sub

Public Sub HandleDoorAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    X = buffer.ReadLong
    Y = buffer.ReadLong

    With TempTile(X, Y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = getTime
    End With

    buffer.Flush: Set buffer = Nothing
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ANIMATION EDITORES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Sub HandleAnimationEditor()
    Dim i As Long

    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

End Sub

Public Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    N = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(N))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(N)), ByVal VarPtr(AnimationData(0)), AnimationSize
    buffer.Flush: Set buffer = Nothing
End Sub
