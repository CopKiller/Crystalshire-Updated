Attribute VB_Name = "Animation_Editor"
Option Explicit

Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean, tmpNum As Long

    If frmEditor_Animation.visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    tmpNum = UBound(soundCache)

    For i = 1 To tmpNum
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With Animation(EditorIndex)
        frmEditor_Animation.txtName.text = Trim$(.Name)

        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            tmpNum = frmEditor_Animation.cmbSound.ListCount

            For i = 0 To tmpNum

                If frmEditor_Animation.cmbSound.list(i) = Trim$(.sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If

        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).value = .sprite(i)
            frmEditor_Animation.scrlFrameCount(i).value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).value = .LoopCount(i)

            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).value = 45
            End If

        Next

        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Animation_Changed(EditorIndex) = True
End Sub

Public Sub AnimationEditorOk()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If

    Next

    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
End Sub

Public Sub AnimationEditorCancel()
    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
End Sub

Public Sub ClearChanged_Animation()
    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
End Sub
