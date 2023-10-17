Attribute VB_Name = "Quest_Handle"
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' MISSION EDITORES
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Public Sub HandleMissionEditor()
    Dim I As Long

    With frmEditor_Quest
        Editor = EDITOR_Mission
        .lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_MISSIONS
            .lstIndex.AddItem I & ": " & Trim$(Mission(I).Name)
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

Public Sub UpdateWindowOffer(ByVal Index_Offer As Long)
    Dim I As Long
    ' gui stuff
    With Windows(GetWindowIndex("winOffer"))
        ' set main text
        If Index_Offer <> 0 Then
            .Controls(GetControlIndex("winOffer", "picBGOffer" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "picOfferBG" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "btnAccept" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "btnRecuse" & Index_Offer)).visible = True
            Select Case inOfferType(Index_Offer)
                Case Offers.Offer_Type_Mission
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = "Missão: " & Mission(inOffer(Index_Offer)).Name & "?"
                Case Offers.Offer_Type_Party
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = inOfferInvite(Index_Offer) & " has invited you to a party."
                Case Offers.Offer_Type_Trade
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = inOfferInvite(Index_Offer) & "  has invited you to trade."
            End Select
            ShowWindow GetWindowIndex("winOffer")
        Else
            For I = 1 To MAX_OFFER
                .Controls(GetControlIndex("winOffer", "picBGOffer" & I)).visible = False
                .Controls(GetControlIndex("winOffer", "picOfferBG" & I)).visible = False
                .Controls(GetControlIndex("winOffer", "lblTitleOffer" & I)).visible = False
                .Controls(GetControlIndex("winOffer", "btnAccept" & I)).visible = False
                .Controls(GetControlIndex("winOffer", "btnRecuse" & I)).visible = False
            Next
            HideWindow GetWindowIndex("winOffer")
        End If
    End With
    
    
End Sub

Public Sub UpdateOffers(Index_Offer)
    Dim I As Long
    
    If Index_Offer <> Offer_HighIndex Then
        For I = Index_Offer To MAX_OFFER
            If I <> Offer_HighIndex Then
                inOffer(I) = inOffer(I + 1)
                inOfferType(I) = inOfferType(I + 1)
                inOfferInvite(I) = inOfferInvite(I + 1)
            Else
                inOffer(I) = 0
                inOfferType(I) = 0
                inOfferInvite(I) = 0
            End If
        Next
    Else
        inOffer(Offer_HighIndex) = 0
        inOfferType(Offer_HighIndex) = 0
        inOfferInvite(Offer_HighIndex) = 0
    End If
    
    Call SetOfferHighIndex
    If Offer_HighIndex > 0 Then
        For I = 1 To Offer_HighIndex
            Call UpdateWindowOffer(I)
        Next
    Else
        Call UpdateWindowOffer(0)
    End If
End Sub

Function FindOpenOfferSlot() As Long
    Dim I As Long
    FindOpenOfferSlot = 0

    For I = 1 To MAX_OFFER
        If inOffer(I) = 0 Then
            FindOpenOfferSlot = I
            Exit Function
        End If
    Next
End Function

Public Sub SetOfferHighIndex()
    Dim I As Integer
    Dim x As Integer
    
    For I = 0 To MAX_OFFER
        x = MAX_OFFER - I

        If inOffer(x) <> 0 Then
            Offer_HighIndex = x
        Exit Sub
        End If

    Next I

    Offer_HighIndex = 0
End Sub

