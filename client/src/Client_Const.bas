Attribute VB_Name = "Client_Const"
Option Explicit
' System of compressing
Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
' in development? [turn off music]
Public Const inDevelopment As Boolean = True
'Loop
Public Const TICKS_PER_SECOND As Long = 60
Public Const SKIP_TICKS = 1000 / TICKS_PER_SECOND
Public Const MAX_FRAME_SKIP = 5
' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 8
Public Const CLIENT_REVISION As Byte = 0
' Connection details
Public Const GAME_SERVER_IP As String = "127.0.0.1" ' "46.23.70.66"
Public Const GAME_SERVER_PORT As Long = 7001 ' the port used by the main game server
' Resolution count
Public Const RES_COUNT As Long = 16
' Music
Public Const MenuMusic = "_menu.mid"
' GUI
Public Const ChatBubbleWidth As Long = 200
Public Const CHAT_TIMER As Long = 20000
' Offer constants
Public Const OfferTop As Long = 0
Public Const OfferLeft As Long = 475
Public Const OfferOffsetY As Long = 37
Public Const OfferOffsetX As Long = 0
Public Const OfferColumns As Long = 1
' Bank constants
Public Const BankTop As Long = 28
Public Const BankLeft As Long = 9
Public Const BankOffsetY As Long = 6
Public Const BankOffsetX As Long = 6
Public Const BankColumns As Long = 10
' Inventory constants
Public Const InvTop As Long = 28
Public Const InvLeft As Long = 9
Public Const InvOffsetY As Long = 6
Public Const InvOffsetX As Long = 6
Public Const InvColumns As Long = 5
' Character consts
Public Const EqTop As Long = 86
Public Const EqLeft As Long = 178
Public Const EqOffsetX As Long = 6
Public Const EqColumns As Long = 6
' Inventory constants
Public Const SkillTop As Long = 28
Public Const SkillLeft As Long = 9
Public Const SkillOffsetY As Long = 6
Public Const SkillOffsetX As Long = 6
Public Const SkillColumns As Long = 5
' Hotbar constants
Public Const HotbarTop As Long = 0
Public Const HotbarLeft As Long = 8
Public Const HotbarOffsetX As Long = 41
' Shop constants
Public Const ShopTop As Long = 28
Public Const ShopLeft As Long = 9
Public Const ShopOffsetY As Long = 6
Public Const ShopOffsetX As Long = 6
Public Const ShopColumns As Long = 7
' Trade
Public Const TradeTop As Long = 0
Public Const TradeLeft As Long = 0
Public Const TradeOffsetY As Long = 6
Public Const TradeOffsetX As Long = 6
Public Const TradeColumns As Long = 5
' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
' Animation
Public Const AnimColumns As Long = 5
' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647
' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"
' Map Path and variables
Public Const MAP_PATH As String = "\Data Files\maps\"
Public Const MAP_EXT As String = ".map"
' Key constants
Public Const VK_A As Long = &H41
Public Const VK_D As Long = &H44
Public Const VK_S As Long = &H53
Public Const VK_W As Long = &H57
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11
Public Const VK_TAB As Long = &H9
Public Const VK_LEFT As Long = &H25
Public Const VK_UP As Long = &H26
Public Const VK_RIGHT As Long = &H27
Public Const VK_DOWN As Long = &H28
' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8
' Speed moving vars
Public Const WALK_SPEED As Byte = 2
Public Const RUN_SPEED As Byte = 4
' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32
' ********************************************************
' * The values below must match with the server's values *
' ********************************************************
' General constants
Public Const MAX_PLAYERS As Long = 200
Public Const MAX_PLAYER_MISSIONS As Long = 12
Public Const MAX_OFFER As Long = 3
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 25
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 10
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_CONVS As Byte = 255
Public Const MAX_NPC_DROPS As Byte = 30
Public Const MAX_NPC_SPELLS As Byte = 10
Public Const MAX_CHARS As Byte = 3
Public Const MAX_WEATHER_PARTICLES As Byte = 250
Public Const MAX_PROJECTILE_PLAYER As Byte = 25
Public Const MAX_PROJECTILE_MAP As Byte = 125

' Website
Public Const GAME_NAME As String = "Crystalshire"
Public Const GAME_WEBSITE As String = "http://www.crystalshire.com"
' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const DESC_LENGTH As Byte = 150
' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1
' Map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 24
Public Const MAX_MAPY As Byte = 18
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_BOSS As Byte = 2
' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_CHAT As Byte = 15
Public Const TILE_TYPE_APPEAR As Byte = 16
Public Const TILE_TYPE_SOUND As Byte = 17
' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_PANTS As Byte = 5
Public Const ITEM_TYPE_FEET As Byte = 6
Public Const ITEM_TYPE_CONSUME As Byte = 7
Public Const ITEM_TYPE_KEY As Byte = 8
Public Const ITEM_TYPE_CURRENCY As Byte = 9
Public Const ITEM_TYPE_SPELL As Byte = 10
Public Const ITEM_TYPE_UNIQUE As Byte = 11
Public Const ITEM_TYPE_FOOD As Byte = 12
' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UP_LEFT As Byte = 4
Public Const DIR_UP_RIGHT As Byte = 5
Public Const DIR_DOWN_LEFT As Byte = 6
Public Const DIR_DOWN_RIGHT As Byte = 7
' Constants for player movement: Tiles per Second
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2
' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4
' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4
' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_PROJECTILE As Byte = 5
' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_CONV As Byte = 7
Public Const EDITOR_Mission As Byte = 8
' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2
Public Const TARGET_TYPE_EVENT As Byte = 3
' Autotiles
Public Const AUTO_INNER As Byte = 1
Public Const AUTO_OUTER As Byte = 2
Public Const AUTO_HORIZONTAL As Byte = 3
Public Const AUTO_VERTICAL As Byte = 4
Public Const AUTO_FILL As Byte = 5
' Autotile types
Public Const AUTOTILE_NONE As Byte = 0
Public Const AUTOTILE_normal As Byte = 1
Public Const AUTOTILE_FAKE As Byte = 2
Public Const AUTOTILE_ANIM As Byte = 3
Public Const AUTOTILE_CLIFF As Byte = 4
Public Const AUTOTILE_WATERFALL As Byte = 5
' Rendering
Public Const RENDER_STATE_NONE As Long = 0
Public Const RENDER_STATE_normal As Long = 1
Public Const RENDER_STATE_AUTOTILE As Long = 2
Public Const RENDER_STATE_APPEAR As Long = 3
' Scrolling action message constants
Public Const ACTIONMsgSTATIC As Long = 0
Public Const ACTIONMsgSCROLL As Long = 1
Public Const ACTIONMsgSCREEN As Long = 2
' text color pointers
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16
Public Const Gold As Byte = 17
Public Const LightGreen As Byte = 18
' pointers
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue
'Weather Type Constants
Public Const WEATHER_TYPE_NONE As Byte = 0
Public Const WEATHER_TYPE_RAIN As Byte = 1
Public Const WEATHER_TYPE_SNOW As Byte = 2
Public Const WEATHER_TYPE_HAIL As Byte = 3
Public Const WEATHER_TYPE_SANDSTORM As Byte = 4
Public Const WEATHER_TYPE_STORM As Byte = 5
'Weather Stuff... events take precedent OVER map settings so we will keep temp map weather settings here.
Public CurrentWeather As Long
Public CurrentWeatherIntensity As Long
Public CurrentFog As Long
Public CurrentFogSpeed As Long
Public CurrentFogOpacity As Long
Public CurrentTintR As Long
Public CurrentTintG As Long
Public CurrentTintB As Long
Public CurrentTintA As Long
Public DrawThunder As Long
