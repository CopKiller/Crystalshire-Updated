Attribute VB_Name = "Server_Globals"
Option Explicit

' Used for closing key doors again
Public KeyTimer As Long

' Used for gradually giving back npcs hp
Public GiveNPCHPTimer As Long

' Used for logging
Public ServerLog As Boolean

' Text vars
Public vbQuote As String

' Maximum classes
Public Max_Classes As Long

' Used for server loop
Public ServerOnline As Boolean

' Used for outputting text
Public NumLines As Long

' Used to handle shutting down server with countdown.
Public isShuttingDown As Boolean
Public Secs As Long
Public TotalPlayersOnline As Long

' GameCPS
Public GameCPS As Long
Public ElapsedTime As Long

' high indexing
Public Player_HighIndex As Long
Public MapProjectile_HighIndex As Integer

' lock the CPS?
Public CPSUnlock As Boolean

' Timers ServerLoop
Public tick As Long
Public TickCPS As Long
Public CPS As Long
Public FrameTime As Long
Public tmr25 As Long
Public tmr500 As Long
Public tmr1000 As Long
Public LastUpdateSavePlayers
Public LastUpdateMapSpawnItems As Long
Public LastUpdatePlayerVitals As Long
