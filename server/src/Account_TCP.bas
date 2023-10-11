Attribute VB_Name = "Account_TCP"
Option Explicit

Sub SendNewCharClasses(ByVal index As Long)
    Dim packet As String
    Dim i As Long, N As Long, q As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.HP)
        Buffer.WriteLong GetClassMaxVital(i, Vitals.MP)
        
        ' set sprite array size
        N = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        N = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub SendPlayerChars(ByVal index As Long)
Dim Buffer As clsBuffer, tmpName As String, i As Long, tmpSprite As Long, tmpAccess As Long, tmpClass As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerChars
    
    ' loop through each character. clear, load, add. repeat.
    For i = 1 To MAX_CHARS
    
        Call LoadPlayer(index, i)
        
        Buffer.WriteString Trim$(Player(index).Name)
        Buffer.WriteLong Player(index).Sprite
        Buffer.WriteLong Player(index).Access
        Buffer.WriteLong Player(index).Class
        
        Call ClearPlayer(index)
    Next
    
    SendDataTo index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub
