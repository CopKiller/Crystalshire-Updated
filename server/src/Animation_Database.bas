Attribute VB_Name = "Animation_Database"
' ****************
' ** animations **
' ****************

Public Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Animation(AnimationNum)
    Close #f
End Sub

Public Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Public Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist(App.Path & "\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Public Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long

    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Animation(i)
        Close #f
    Next

End Sub

Public Sub ClearAnimation(ByVal index As Long)
    Animation(index) = EmptyAnimation
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
End Sub

Public Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub
