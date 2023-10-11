Attribute VB_Name = "Animation_Database"
Public Sub ClearAnimInstance(ByVal Index As Long)
    AnimInstance(Index) = EmptyAnimInstance
End Sub

Public Sub ClearAnimation(ByVal Index As Long)
    Animation(Index) = EmptyAnimation
    Animation(Index).Name = vbNullString
    Animation(Index).sound = "None."
End Sub

Public Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

End Sub
