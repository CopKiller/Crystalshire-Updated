Attribute VB_Name = "Server_Database"
Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

Public Sub CheckDirs()
    ChkDir App.Path & "\data\", "accounts"
    ChkDir App.Path & "\data\", "animations"
    ChkDir App.Path & "\data\", "items"
    ChkDir App.Path & "\data\", "logs"
    ChkDir App.Path & "\data\", "maps"
    ChkDir App.Path & "\data\", "npcs"
    ChkDir App.Path & "\data\", "resources"
    ChkDir App.Path & "\data\", "shops"
    ChkDir App.Path & "\data\", "spells"
    ChkDir App.Path & "\data\", "convs"
    ChkDir App.Path & "\data\", "missions"
End Sub
Public Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Resources...")
    Call LoadResources
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading conversations...")
    Call LoadConvs
    Call SetStatus("Loading quests...")
    Call LoadMissions
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing conversations...")
    Call ClearConvs
    Call SetStatus("Clearing quests...")
    Call ClearMissions
End Sub
