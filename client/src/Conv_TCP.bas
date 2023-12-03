Attribute VB_Name = "Conv_TCP"
Public Sub SendRequestConvs()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestConvs
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestEditConv()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditConv
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSaveConv(ByVal Convnum As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim X As Long
    Set buffer = New clsBuffer
    buffer.WriteLong CSaveConv
    buffer.WriteLong Convnum

    With Conv(Convnum)
        buffer.WriteString .Name
        buffer.WriteLong .chatCount

        For i = 1 To .chatCount
            buffer.WriteString .Conv(i).Conv

            For X = 1 To 4
                buffer.WriteString .Conv(i).rText(X)
                buffer.WriteLong .Conv(i).rTarget(X)
            Next

            buffer.WriteLong .Conv(i).EventType
            buffer.WriteLong .Conv(i).EventNum
        Next

    End With

    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
