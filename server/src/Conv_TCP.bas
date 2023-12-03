Attribute VB_Name = "Conv_TCP"
Public Sub SendUpdateConvTo(ByVal index As Long, ByVal convNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim x As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateConv
    Buffer.WriteLong convNum
    With Conv(convNum)
        Buffer.WriteString .Name
        Buffer.WriteLong .chatCount
        For i = 1 To .chatCount
            Buffer.WriteString .Conv(i).Conv
            For x = 1 To 4
                Buffer.WriteString .Conv(i).rText(x)
                Buffer.WriteLong .Conv(i).rTarget(x)
            Next
            Buffer.WriteLong .Conv(i).EventType
            Buffer.WriteLong .Conv(i).eventNum
        Next
    End With
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Public Sub SendConvs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_CONVS
        If LenB(Trim$(Conv(i).Name)) > 0 Then
            Call SendUpdateConvTo(index, i)
        End If
    Next
End Sub

Public Sub SendUpdateConvToAll(ByVal convNum As Long)
    Dim packet As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim x As Long

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SUpdateConv
    Buffer.WriteLong convNum
    With Conv(convNum)
        Buffer.WriteString .Name
        Buffer.WriteLong .chatCount
        For i = 1 To .chatCount
            Buffer.WriteString .Conv(i).Conv
            For x = 1 To 4
                Buffer.WriteString .Conv(i).rText(x)
                Buffer.WriteLong .Conv(i).rTarget(x)
            Next
            Buffer.WriteLong .Conv(i).EventType
            Buffer.WriteLong .Conv(i).eventNum
        Next
    End With
    
    SendDataToAll Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub


