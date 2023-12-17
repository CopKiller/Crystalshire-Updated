Attribute VB_Name = "Quest_UDT"
Option Explicit

Public Const MAX_MISSIONS As Long = 255

' Mission Type Const
Public Enum MissionType
    Mission_TypeKill = 0
    Mission_TypeCollect
    Mission_TypeTalk
    ' MAX MissionType
    MissionType_Count
End Enum

Public Mission(1 To MAX_MISSIONS) As MissionRec
Public EmptyMission As MissionRec

Private Type ItemReward
    ItemNum As Long
    ItemAmount As Long
End Type

Private Type MissionRec
    ' General
    Name As String * NAME_LENGTH
    Type As Long
    Repeatable As Byte
    Description As String * DESC_LENGTH
    ' Mission Type Kill
    KillNPC As Long
    KillNPCAmount As Long
    ' Mission Type Collect
    CollectItem As Long
    CollectItemAmount As Long
    ' Mission Type Talk
    TalkNPC As Long
    ' Next Mission
    PreviousMissionComplete As Long
    ' Message
    Incomplete As String * DESC_LENGTH
    Completed As String * DESC_LENGTH
    ' Reward
    RewardItem(1 To 5) As ItemReward
    RewardExperience As Long
End Type
