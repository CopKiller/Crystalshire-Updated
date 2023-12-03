Attribute VB_Name = "modEvents"
Option Explicit

Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

' temporary event
Public cpEvent As EventRec

Sub CopyEvent_Map(X As Long, y As Long)
Dim count As Long, i As Long
    count = Map.TileData.EventCount
    If count = 0 Then Exit Sub
    
    For i = 1 To count
        If Map.TileData.Events(i).X = X And Map.TileData.Events(i).y = y Then
            ' copy it
            CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(Map.TileData.Events(i)), LenB(Map.TileData.Events(i))
            ' exit
            Exit Sub
        End If
    Next
End Sub




Sub ClearEvent(EventNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map.TileData.Events(EventNum)), LenB(Map.TileData.Events(EventNum)))
End Sub

Sub CopyEvent(original As Long, newone As Long)
    CopyMemory ByVal VarPtr(Map.TileData.Events(newone)), ByVal VarPtr(Map.TileData.Events(original)), LenB(Map.TileData.Events(original))
End Sub

Sub EventEditorInit(EventNum As Long)
Dim i As Long
    EditorEvent = EventNum
    ' copy the event data to the temp event
    CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(Map.TileData.Events(EventNum)), LenB(Map.TileData.Events(EventNum))
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For i = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , Str(i)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "None"
        For i = 1 To MAX_ITEMS
            .cmbHasItem.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "None"
        For i = 1 To MAX_BYTE
            .cmbPlayerVar.AddItem i
        Next
        ' name
        .txtName.text = tmpEvent.Name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.enabled = True
        Else
            .cmdDeletePage.enabled = False
        End If
        .cmdPastePage.enabled = False
        ' set the commands frame
        .fraCommands.Width = 417
        .fraCommands.Height = 497
        ' set the dialogue frame
        .fraDialogue.Width = 417
        .fraDialogue.Height = 497
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show
End Sub

Sub AddCommand(theType As EventType)
Dim count As Long
    ' update the array
    With tmpEvent.EventPage(curPageNum)
        count = .CommandCount + 1
        ReDim Preserve .Commands(1 To count)
        .CommandCount = count
        ' set the shit
        Select Case theType
            Case EventType.evAddText
                ' set the values
                .Commands(count).Type = EventType.evAddText
                .Commands(count).text = frmEditor_Events.txtAddText_Text.text
                .Commands(count).Colour = frmEditor_Events.scrlAddText_Colour.value
                If frmEditor_Events.optAddText_Game.value Then
                    .Commands(count).Channel = 0
                ElseIf frmEditor_Events.optAddText_Map.value Then
                    .Commands(count).Channel = 1
                ElseIf frmEditor_Events.optAddText_Global.value Then
                    .Commands(count).Channel = 2
                End If
            Case EventType.evShowChatBubble
                .Commands(count).Type = EventType.evShowChatBubble
                .Commands(count).text = frmEditor_Events.txtChatBubble.text
                .Commands(count).Colour = frmEditor_Events.scrlChatBubble.value
                .Commands(count).TargetType = frmEditor_Events.cmbChatBubbleType.ListIndex
                .Commands(count).target = frmEditor_Events.cmbChatBubble.ListIndex
            Case EventType.evPlayerVar
                .Commands(count).Type = EventType.evPlayerVar
                .Commands(count).target = frmEditor_Events.cmbVariable.ListIndex
                .Commands(count).Colour = Val(frmEditor_Events.txtVariable.text)
            Case EventType.evWarpPlayer
                .Commands(count).Type = EventType.evWarpPlayer
                .Commands(count).X = frmEditor_Events.scrlWPX.value
                .Commands(count).y = frmEditor_Events.scrlWPY.value
                .Commands(count).target = frmEditor_Events.scrlWPMap.value
        End Select
    End With
    ' re-list the commands
    EventListCommands
End Sub

Sub EditCommand()
    With tmpEvent.EventPage(curPageNum).Commands(curCommand)
        Select Case .Type
            Case EventType.evAddText
                .text = frmEditor_Events.txtAddText_Text.text
                .Colour = frmEditor_Events.scrlAddText_Colour.value
                If frmEditor_Events.optAddText_Game.value Then
                    .Channel = 0
                ElseIf frmEditor_Events.optAddText_Map.value Then
                    .Channel = 1
                ElseIf frmEditor_Events.optAddText_Global.value Then
                    .Channel = 2
                End If
            Case EventType.evShowChatBubble
                .text = frmEditor_Events.txtChatBubble.text
                .Colour = frmEditor_Events.scrlChatBubble.value
                .TargetType = frmEditor_Events.cmbChatBubbleType.ListIndex
                .target = frmEditor_Events.cmbChatBubble.ListIndex
            Case EventType.evPlayerVar
                .target = frmEditor_Events.cmbVariable.ListIndex
                .Colour = Val(frmEditor_Events.txtVariable.text)
            Case EventType.evWarpPlayer
                .X = frmEditor_Events.scrlWPX.value
                .y = frmEditor_Events.scrlWPY.value
        End Select
    End With
    ' re-list the commands
    EventListCommands
End Sub

Sub EventListCommands()
Dim i As Long, count As Long
    frmEditor_Events.lstCommands.Clear
    ' check if there are any
    count = tmpEvent.EventPage(curPageNum).CommandCount
    If count > 0 Then
        ' list them
        For i = 1 To count
            With tmpEvent.EventPage(curPageNum).Commands(i)
                Select Case .Type
                    Case EventType.evAddText
                        ListCommandAdd "@>Add Text: " & .text & " - Colour: " & GetColourString(.Colour) & " - Channel: " & .Channel
                    Case EventType.evShowChatBubble
                        ListCommandAdd "@>Show Chat Bubble: " & .text & " - Colour: " & GetColourString(.Colour) & " - Target Type: " & .TargetType & " - Target: " & .target
                    Case EventType.evPlayerVar
                        ListCommandAdd "@>Change variable #" & .target & " to " & .Colour
                    Case EventType.evWarpPlayer
                        ListCommandAdd "@>Warp Player to Map #" & .target & ", X: " & .X & ", Y: " & .y
                    Case Else
                        ListCommandAdd "@>Unknown"
                End Select
            End With
        Next
    Else
        frmEditor_Events.lstCommands.AddItem "@>"
    End If
    frmEditor_Events.lstCommands.ListIndex = 0
    curCommand = 1
End Sub

Sub ListCommandAdd(s As String)
Static X As Long
    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If X < frmEditor_Events.TextWidth(s & "  ") Then
       X = frmEditor_Events.TextWidth(s & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX ' if twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    End If
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' populate form
    With tmpEvent.EventPage(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemNum
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        frmEditor_Events.cmbPlayerVar.ListIndex = .PlayerVarNum
        frmEditor_Events.cmbPriority.ListIndex = .Priority
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchNum
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.value = .DirFix
        frmEditor_Events.chkHasItem.value = .chkHasItem
        frmEditor_Events.chkPlayerVar.value = .chkPlayerVar
        frmEditor_Events.chkSelfSwitch.value = .chkSelfSwitch
        frmEditor_Events.chkStepAnim.value = .StepAnim
        frmEditor_Events.chkWalkAnim.value = .WalkAnim
        frmEditor_Events.chkWalkThrough.value = .WalkThrough
        frmEditor_Events.txtPlayerVariable = .PlayerVariable
        frmEditor_Events.scrlGraphic.value = .Graphic
        If .chkHasItem = 0 Then frmEditor_Events.cmbHasItem.enabled = False Else frmEditor_Events.cmbHasItem.enabled = True
        If .chkSelfSwitch = 0 Then frmEditor_Events.cmbSelfSwitch.enabled = False Else frmEditor_Events.cmbSelfSwitch.enabled = True
        If .chkPlayerVar = 0 Then
            frmEditor_Events.cmbPlayerVar.enabled = False
            frmEditor_Events.txtPlayerVariable.enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.enabled = True
            frmEditor_Events.txtPlayerVariable.enabled = True
        End If
        ' show the commands
        EventListCommands
    End With
End Sub

Sub EventEditorOK()
    ' copy the event data from the temp event
    CopyMemory ByVal VarPtr(Map.TileData.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)
    ' unload the form
    Unload frmEditor_Events
End Sub
