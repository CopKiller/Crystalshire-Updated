Attribute VB_Name = "Shop_UDT"
Option Explicit

Public Const MAX_TRADES As Long = 35

Public Shop(1 To MAX_SHOPS) As ShopRec
Public EmptyShop As ShopRec

Public Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type
