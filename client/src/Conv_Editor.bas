Attribute VB_Name = "Conv_Editor"
Option Explicit

Public Conv_Changed(1 To MAX_CONVS) As Boolean

' /////////////////
' // Conv Editor //
' /////////////////
Public Sub ConvEditorInit()
    Dim i As Long, N As Long

    If frmEditor_Conv.visible = False Then Exit Sub
    EditorIndex = frmEditor_Conv.lstIndex.ListIndex + 1

    With frmEditor_Conv
        .txtName.text = Trim$(Conv(EditorIndex).Name)

        If Conv(EditorIndex).chatCount = 0 Then
            Conv(EditorIndex).chatCount = 1
            ReDim Conv(EditorIndex).Conv(1 To Conv(EditorIndex).chatCount)
        End If

        For N = 1 To 4
            .cmbReply(N).Clear
            .cmbReply(N).AddItem "None"

            For i = 1 To Conv(EditorIndex).chatCount
                .cmbReply(N).AddItem i
            Next
        Next

        .scrlChatCount = Conv(EditorIndex).chatCount
        .scrlConv.max = Conv(EditorIndex).chatCount
        .scrlConv.value = 1
        .txtConv = Conv(EditorIndex).Conv(.scrlConv.value).Conv

        For i = 1 To 4
            .txtReply(i).text = Conv(EditorIndex).Conv(.scrlConv.value).rText(i)
            .cmbReply(i).ListIndex = Conv(EditorIndex).Conv(.scrlConv.value).rTarget(i)
        Next
        
        .cmbEvent.ListIndex = Conv(EditorIndex).Conv(.scrlConv.value).EventType
        
        If .cmbEvent.ListIndex = EventType.Event_OpenShop Then
            ' build EventNum combo
            .cmbEventNum.Clear
            .cmbEventNum.AddItem "None"
    
            For i = 1 To MAX_SHOPS
                .cmbEventNum.AddItem Trim$(Shop(i).Name)
            Next
    
            .cmbEventNum.ListIndex = Conv(EditorIndex).Conv(.scrlConv.value).EventNum
        ElseIf .cmbEvent.ListIndex = EventType.Event_OpenQuest Then
            ' build EventNum combo
            .cmbEventNum.Clear
            .cmbEventNum.AddItem "None"
    
            For i = 1 To MAX_MISSIONS
                .cmbEventNum.AddItem Trim$(Mission(i).Name)
            Next
    
            .cmbEventNum.ListIndex = Conv(EditorIndex).Conv(.scrlConv.value).EventNum
        End If
    End With

    Conv_Changed(EditorIndex) = True
End Sub

Public Sub ConvEditorOk()
    Dim i As Long

    For i = 1 To MAX_CONVS

        If Conv_Changed(i) Then
            Call SendSaveConv(i)
        End If

    Next

    Unload frmEditor_Conv
    Editor = 0
    ClearChanged_Conv
End Sub

Public Sub ConvEditorCancel()
    Editor = 0
    Unload frmEditor_Conv
    ClearChanged_Conv
    ClearConvs
    SendRequestConvs
End Sub

Public Sub ClearChanged_Conv()
    ZeroMemory Conv_Changed(1), MAX_CONVS * 2 ' 2 = boolean length
End Sub
