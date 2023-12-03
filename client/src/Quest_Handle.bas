Attribute VB_Name = "Quest_Handle"
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' MISSION EDITORES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Sub HandleMissionEditor()
    Dim i As Long

    With frmEditor_Quest
        Editor = EDITOR_Mission
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_MISSIONS
            .lstIndex.AddItem i & ": " & Trim$(Mission(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        MissionEditorInit
    End With

End Sub

Public Sub HandleUpdateMission(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim MissionSize As Long
    Dim MissionData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    N = buffer.ReadLong
    MissionSize = LenB(Mission(N))
    
    ReDim MissionData(MissionSize - 1)
    MissionData = buffer.ReadBytes(MissionSize)
    
    ClearMission N
    CopyMemory ByVal VarPtr(Mission(N)), ByVal VarPtr(MissionData(0)), MissionSize
    
    buffer.Flush: Set buffer = Nothing
End Sub

Public Sub HandleOfferMission(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Index_Offer As Integer
    Set buffer = New clsBuffer

    buffer.WriteBytes Data()
    Index_Offer = FindOpenOfferSlot
    If Index_Offer <> 0 Then
        inOffer(Index_Offer) = buffer.ReadLong
        inOfferType(Index_Offer) = Offers.Offer_Type_Mission
    End If
    buffer.Flush: Set buffer = Nothing
    
    Call UpdateWindowOffer(Index_Offer)
End Sub

Public Sub UpdateOffers(Index_Offer)
    Dim i As Long
    
    If Index_Offer <> Offer_HighIndex Then
        For i = Index_Offer To MAX_OFFER
            If i <> Offer_HighIndex And i < MAX_OFFER Then
                inOffer(i) = inOffer(i + 1)
                inOfferType(i) = inOfferType(i + 1)
                inOfferInvite(i) = inOfferInvite(i + 1)
            Else
                inOffer(i) = 0
                inOfferType(i) = 0
                inOfferInvite(i) = 0
            End If
        Next
    Else
        inOffer(Offer_HighIndex) = 0
        inOfferType(Offer_HighIndex) = 0
        inOfferInvite(Offer_HighIndex) = 0
    End If
    
    Call SetOfferHighIndex
    If Offer_HighIndex > 0 Then
        For i = 1 To Offer_HighIndex
            Call UpdateWindowOffer(i)
        Next
    Else
        Call UpdateWindowOffer(0)
    End If
End Sub

Function FindOpenOfferSlot() As Long
    Dim i As Long
    FindOpenOfferSlot = 0

    For i = 1 To MAX_OFFER
        If inOffer(i) = 0 Then
            FindOpenOfferSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub SetOfferHighIndex()
    Dim i As Integer
    Dim X As Integer
    
    For i = 0 To MAX_OFFER
        X = MAX_OFFER - i
        If X > 0 Then
            If inOffer(X) <> 0 Then
                Offer_HighIndex = X
            Exit Sub
            End If
        End If

    Next i

    Offer_HighIndex = 0
End Sub

