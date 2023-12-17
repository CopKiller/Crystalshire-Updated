Attribute VB_Name = "Server_Enum"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
    SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame ' = 5
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp ' = 10
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir ' = 15
    SNpcDir
    SPlayerXY
    SPlayerXYMap
    SAttack
    SNpcAttack ' = 20
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone ' = 25
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem ' = 30
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead ' = 35
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor ' = 40
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft ' = 45
    SResourceCache
    SResourceEditor
    SUpdateResource
    SSendPing
    SDoorAnimation ' = 50
    SActionMsg
    SPlayerEXP
    SBlood
    SAnimationEditor
    SUpdateAnimation ' = 55
    SAnimation
    SMapNpcVitals
    SCooldown
    SClearSpellBuffer
    SSayMsg ' = 60
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SBank ' = 65
    STrade
    SCloseTrade
    STradeUpdate
    STradeStatus
    STarget ' = 70
    SHotbar
    SHighIndex
    SSound
    STradeRequest
    SPartyInvite ' = 75
    SPartyUpdate
    SPartyVitals
    SChatUpdate
    SConvEditor
    SUpdateConv ' = 80
    SStartTutorial
    SChatBubble
    SPlayerChars
    SCancelAnimation
    SMissionEditor ' = 85
    SUpdateMission
    SPlayerMission
    SPlayerVariables
    SOfferMission ' = 90
    SProjectileAttack
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelChar
    CLogin
    CAddChar
    CUseChar ' = 5
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove ' = 10
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest ' = 15
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats ' = 20
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem ' = 25
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy ' = 30
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc ' = 35
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell ' = 40
    CSetAccess
    CWhosOnline
    CSetMotd
    CTarget
    CSpells ' = 45
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource ' = 50
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS ' = 55
    CRequestResources
    CSpawnItem
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations ' = 60
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop ' = 65
    CBuyItem
    CSellItem
    CChangeBankSlots
    CDepositItem
    CWithdrawItem ' = 70
    CCloseBank
    CAdminWarp
    CTradeRequest
    CAcceptTrade
    CDeclineTrade ' = 75
    CTradeItem
    CUntradeItem
    CHotbarChange
    CHotbarUse
    CSwapSpellSlots ' = 80
    CAcceptTradeRequest
    CDeclineTradeRequest
    CPartyRequest
    CAcceptParty
    CDeclineParty
    CPartyLeave
    CChatOption
    CRequestEditConv
    CSaveConv
    CRequestConvs ' = 90
    CRequestEditMission
    CSaveMission
    CRequestMissions
    CAcceptMissionRequest
    CDeclineMissionRequest
    CFinishTutorial
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    Strength = 1
    Endurance
    Intelligence
    Agility
    Willpower
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    HP = 1
    MP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor
    Helmet
    Shield
    Pants
    Feet
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Usando para eventos comuns
Public Enum EventType
    Event_OpenShop = 1
    Event_OpenBank
    Event_OpenQuest
    
    Event_Count
End Enum

' Layers in a map
Public Enum MapLayer
    Ground = 1
    Mask
    Mask2
    Fringe
    Fringe2
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum

' Sound entities
Public Enum SoundEntity
    seAnimation = 1
    seItem
    seNpc
    seResource
    seSpell
    ' Make sure SoundEntity_Count is below everything else
    SoundEntity_Count
End Enum
