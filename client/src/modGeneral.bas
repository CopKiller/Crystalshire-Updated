Attribute VB_Name = "modGeneral"
Option Explicit
' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)

Public Sub Main()
Dim i As Long
    InitCRC32
    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\data files\", "graphics"
    ChkDir App.Path & "\data files\graphics\", "animations"
    ChkDir App.Path & "\data files\graphics\", "characters"
    ChkDir App.Path & "\data files\graphics\", "items"
    ChkDir App.Path & "\data files\graphics\", "paperdolls"
    ChkDir App.Path & "\data files\graphics\", "resources"
    ChkDir App.Path & "\data files\graphics\", "spellicons"
    ChkDir App.Path & "\data files\graphics\", "tilesets"
    ChkDir App.Path & "\data files\graphics\", "faces"
    ChkDir App.Path & "\data files\graphics\", "gui"
    ChkDir App.Path & "\data files\", "logs"
    ChkDir App.Path & "\data files\", "maps"
    ChkDir App.Path & "\data files\", "music"
    ChkDir App.Path & "\data files\", "sound"
    ChkDir App.Path & "\data files\", "video"
    ' load options
    LoadOptions
    ' check the resolution
    CheckResolution
    ' load dx8
    If Options.Fullscreen Then
        frmMain.BorderStyle = 0
        frmMain.caption = frmMain.caption
    End If
    frmMain.Show
    InitDX8 frmMain.hWnd
    DoEvents
    
    Set GDIToken = New cGDIpToken
    If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
    
    LoadTextures
    LoadFonts
    ' initialise the gui
    InitGUI
    ' Resize the GUI to screen size
    ResizeGUI
    ' initialise sound & music engines
    InitFmod
    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    vbQuote = ChrW$(34)
    ' Update the form with the game's name before it's loaded
    frmMain.caption = GAME_NAME
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings.")
    Call TcpInit(GAME_SERVER_IP, GAME_SERVER_PORT)
    Call InitMessages
    ' Reset values
    Ping = -1
    ' cache the buttons then reset & render them
    Call SetStatus("Caching map CRC32 checksums...")
    ' cache map crc32s
    For i = 1 To MAX_MAPS
        GetMapCRC32 i
    Next
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    ' set the paperdoll order
    ReDim PaperdollOrder(1 To Equipment.Equipment_Count - 1) As Long
    Call SetPaperdollOrder
    ' set status
    SetStatus vbNullString
    ' show the main menu
    frmMain.Show
    inMenu = True
    ' show login window
    ShowWindow GetWindowIndex("winLogin")
    inSmallChat = True
    ' Set the loop going
    fadeAlpha = 255
    If Options.PlayIntro = 1 Then
        PlayIntro
    Else
        videoPlaying = False
        frmMain.picIntro.visible = False
        ' play the menu music
        If Len(Trim$(MenuMusic)) > 0 Then Play_Music Trim$(MenuMusic)
    End If
    Call InitTime
    
    Call MenuLoop
End Sub

Public Sub HandleDeveloperOptions()
'//-> Verifica se está no modo desenvolvimento, então deixa o botão de atualizar janelas visivel <--
    With frmMain
        If App.LogMode = 0 Then
            .cmdAttWindow.visible = True
            .cmdAttWindow.Left = (.ScaleWidth / 2) - (.cmdAttWindow.Width / 2)
        Else
            .cmdAttWindow.visible = False
        End If
    End With
End Sub

Public Sub SetPaperdollOrder(Optional ByVal Swap As Boolean = False)
    PaperdollOrder(1) = Equipment.Feet
    PaperdollOrder(2) = Equipment.Shield
    PaperdollOrder(3) = Equipment.Armor
    PaperdollOrder(4) = Equipment.Pants
    PaperdollOrder(5) = Equipment.Helmet
    PaperdollOrder(6) = Equipment.Weapon
    
    If Not Swap Then Exit Sub
    'PaperdollOrder(6) = Equipment.Glasses
    'PaperdollOrder(7) = Equipment.Facemask
    'PaperdollOrder(8) = Equipment.Hat
    'PaperdollOrder(9) = Equipment.Acessory
    'PaperdollOrder(10) = Equipment.Vest
End Sub

Public Sub AddChar(Name As String, Sex As Long, Class As Long, sprite As Long)

    If ConnectToServer Then
        Call SetStatus("Sending character information.")
        Call SendAddChar(Name, Sex, Class, sprite)
        Exit Sub
    Else
        ShowWindow GetWindowIndex("winLogin")
        Dialogue "Connection Problem", "Cannot connect to game server.", "", TypeALERT
    End If

End Sub

Public Sub Login(Name As String, password As String)
    'TcpInit GAME_SERVER_IP, GAME_SERVER_PORT

    If ConnectToServer Then
        Call SetStatus("Sending login information.")
        Call SendLogin(Name, password)
        ' save details
        If Options.SaveUser Then Options.Username = Name Else Options.Username = vbNullString
        SaveOptions
        Exit Sub
    Else
        ShowWindow GetWindowIndex("winLogin")
        Dialogue "Connection Problem", "Cannot connect to login server.", "Please try again later.", TypeALERT
    End If

End Sub
Public Sub SendRegister(Name As String, password As String, codigo As String)
    'TcpInit GAME_SERVER_IP, GAME_SERVER_PORT

    If ConnectToServer Then
        Call SetStatus("Sending Register information.")
        Call SendNewAccount(Name, password, codigo)
    Else
        ShowWindow GetWindowIndex("winregister")
        Dialogue "Connection Problem", "Cannot connect to login server.", "Please try again later.", TypeALERT
    End If

End Sub
Public Sub logoutGame()
    Dim i As Long
    isLogging = True
    InGame = False
    
    DestroyTCP

    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next

    ' destroy temp values
    DragInvSlotNum = 0
    LastItemDesc = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    ' clear chat
    For i = 1 To ChatLines
        Chat(i).text = vbNullString
    Next

    inMenu = True
    MenuLoop
End Sub

Sub GameInit()
    Dim musicFile As String
    ' hide gui
    InBank = False
    InTrade = False
    CloseShop
    ' get ping
    GetPing
    ' play music
    musicFile = Trim$(Map.MapData.Music)

    If Not musicFile = "None." Then
        Play_Music musicFile
    Else
        Stop_Music
    End If

    SetStatus vbNullString
End Sub

Public Sub DestroyGame()
    StopIntro
    Call DestroyTCP
    ' destroy music & sound engines
    Destroy_Music
    Call UnloadAllForms
    End
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next

End Sub

Public Sub SetStatus(ByVal caption As String)
    HideWindows
    If Len(Trim$(caption)) > 0 Then
        ShowWindow GetWindowIndex("winLoading")
        Windows(GetWindowIndex("winLoading")).Controls(GetControlIndex("winLoading", "lblLoading")).text = caption
    Else
        HideWindow GetWindowIndex("winLoading")
        Windows(GetWindowIndex("winLoading")).Controls(GetControlIndex("winLoading", "lblLoading")).text = vbNullString
    End If
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)

    If NewLine Then
        Txt.text = Txt.text + Msg + vbCrLf
    Else
        Txt.text = Txt.text + Msg
    End If

    Txt.SelStart = Len(Txt.text) - 1
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal password As String) As Boolean

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
    Dim i As Long, tmpNum As Long
    ' Prevent high ascii chars
    tmpNum = Len(sInput)

    For i = 1 To tmpNum

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Dialogue "Illegal Characters", "This string contains illegal characters.", "", TypeALERT
            Exit Function
        End If

    Next

    isStringLegal = True
End Function

Public Sub PopulateLists()
    Dim strLoad As String, i As Long
    ' Cache music list
    strLoad = Dir$(App.Path & MUSIC_PATH & "*.*")
    i = 1

    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop

    ' Cache sound list
    strLoad = Dir$(App.Path & SOUND_PATH & "*.*")
    i = 1

    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = Dir
        i = i + 1
    Loop

End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal targetX As Integer, ByVal targetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = targetY Then
        'Check for going right (90 degrees)
        If CenterX < targetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
        
        'Exit the function
        Exit Function
    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = targetX Then
        'Check for going up (360 degrees)
        If CenterY > targetY Then
            Engine_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function
        Exit Function
    End If

    'Calculate Side C
    SideC = Sqr(Abs(targetX - CenterX) ^ 2 + Abs(targetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(targetX - CenterX) ^ 2 + targetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If targetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    'Exit function
    Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function
End Function
