Attribute VB_Name = "Shop_Editor"
Option Explicit

Public Shop_Changed(1 To MAX_SHOPS) As Boolean
' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
    Dim i As Long

    If frmEditor_Shop.visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    frmEditor_Shop.txtName.text = Trim$(Shop(EditorIndex).Name)

    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.value = 100
    End If

    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    UpdateShopTrade
    Shop_Changed(EditorIndex) = True
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
    Dim i As Long
    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES

        With Shop(EditorIndex).TradeItem(i)

            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).Name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).Name)
            End If

        End With

    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
End Sub

Public Sub ShopEditorOk()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If

    Next

    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
End Sub

Public Sub ShopEditorCancel()
    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
End Sub

Public Sub ClearChanged_Shop()
    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
End Sub
