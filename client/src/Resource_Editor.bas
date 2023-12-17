Attribute VB_Name = "Resource_Editor"
Option Explicit

Public Resource_Changed(1 To MAX_RESOURCES) As Boolean

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
    Dim i As Long
    Dim SoundSet As Boolean

    If frmEditor_Resource.visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1

    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If

    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."

    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next

    ' finished populating
    With frmEditor_Resource
        .scrlExhaustedPic.max = CountResource
        .scrlNormalPic.max = CountResource
        .scrlAnimation.max = MAX_ANIMATIONS
        .txtName.text = Trim$(Resource(EditorIndex).name)
        .txtMessage.text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .scrlHealth.Value = Resource(EditorIndex).health
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation

        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then

            For i = 0 To .cmbSound.ListCount

                If .cmbSound.list(i) = Trim$(Resource(EditorIndex).sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If

            Next

            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If

    End With

    Resource_Changed(EditorIndex) = True
End Sub

Public Sub ResourceEditorOk()
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If

    Next

    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
End Sub

Public Sub ResourceEditorCancel()
    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
End Sub

Public Sub ClearChanged_Resource()
    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
End Sub
