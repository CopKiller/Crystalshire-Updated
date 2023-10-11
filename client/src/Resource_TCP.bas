Attribute VB_Name = "Resource_TCP"
Public Sub SendRequestEditResource()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditResource
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal N As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set buffer = New clsBuffer
    ResourceSize = LenB(Resource(N))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(N)), ResourceSize
    buffer.WriteLong CSaveResource
    buffer.WriteLong N
    buffer.WriteBytes ResourceData
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub SendRequestResources()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestResources
    SendData buffer.ToArray()
    buffer.Flush: Set buffer = Nothing
End Sub
