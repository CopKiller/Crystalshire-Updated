Attribute VB_Name = "Quest_UDT"
Option Explicit

Public Quest(1 To MAX_QUESTS) As QuestRec

Private Type QuestRec
    name As String
    Repeatable As Byte
    Dialogue As String
    Type As Long
    KillNPC As Long
    KillNPCAmount As Long
    CollectItem As Long
    CollectItemAmount As Long
    Incomplete As String
    Completed As String
    RewardCurrency As Long
    RewardItem As Long
    RewardItemAmount As Long
    RewardExperience As Long
    TalkNPC As Long
    PreviousQuestComplete As Long
    SwitchID As Long
    SwitchValue As Long
End Type
