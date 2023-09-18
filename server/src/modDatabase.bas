Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private crcTable(0 To 255) As Long

Public Sub InitCRC32()
    Dim i As Long, n As Long, CRC As Long

    For i = 0 To 255
        CRC = i
        For n = 0 To 7
            If CRC And 1 Then
                CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next
        crcTable(i) = CRC
    Next
End Sub

Public Function CRC32(ByRef Data() As Byte) As Long
    Dim lCurPos As Long
    Dim lLen As Long

    lLen = AryCount(Data) - 1
    CRC32 = &HFFFFFFFF

    For lCurPos = 0 To lLen
        CRC32 = (((CRC32 And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((CRC32 And 255) Xor Data(lCurPos)))
    Next

    CRC32 = CRC32 Xor &HFFFFFFFF
End Function

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
    Dim filename As String
    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
    Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
    Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
    Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim f As Long

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If

        f = FreeFile
        Open filename For Append As #f
        Print #f, DateValue(Now) & " " & Time & ": " & Text
        Close #f
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

'//This check if the file exist
Public Function FileExist(ByVal filename As String) As Boolean
' Checking if File Exist
    If LenB(dir(filename)) > 0 Then FileExist = True
End Function

Public Sub SaveOptions()
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
End Sub

Public Sub LoadOptions()
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
End Sub

Public Sub ToggleMute(ByVal index As Long)
' exit out for rte9
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub

    ' toggle the player's mute
    If Player(index).isMuted = 1 Then
        Player(index).isMuted = 0
        ' Let them know
        PlayerMsg index, "You have been unmuted and can now talk in global.", BrightGreen
        TextAdd GetPlayerName(index) & " has been unmuted."
    Else
        Player(index).isMuted = 1
        ' Let them know
        PlayerMsg index, "You have been muted and can no longer talk in global.", BrightRed
        TextAdd GetPlayerName(index) & " has been muted."
    End If

    ' save the player
    SavePlayer index
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String, IP As String, f As Long, i As Long

    ' Add banned to the player's index
    Player(BanPlayerIndex).isBanned = 1
    SavePlayer BanPlayerIndex

    ' IP banning
    filename = App.Path & "\data\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(filename) Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)
    f = FreeFile
    Open filename For Append As #f
    Print #f, IP
    Close #f

    ' Tell them they're banned
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & ".", White)
    Call AddLog(GetPlayerName(BanPlayerIndex) & " has been banned.", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, DIALOGUE_MSG_BANNED)
End Sub

Public Function isBanned_IP(ByVal IP As String) As Boolean
    Dim filename As String, fIP As String, f As Long

    filename = App.Path & "\data\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(filename) Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    f = FreeFile
    Open filename For Input As #f

    Do While Not EOF(f)
        Input #f, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #f
            Exit Function
        End If
    Loop

    Close #f
End Function

Public Function isBanned_Account(ByVal index As Long) As Boolean
    If Player(index).isBanned = 1 Then
        isBanned_Account = True
    Else
        isBanned_Account = False
    End If
End Function

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long, ByVal charNum As Long) As Boolean
    Dim theName As String
    theName = GetVar(App.Path & "\data\accounts\CharNum_" & charNum & ".bin", "CHAR" & charNum, "Name")
    'If LenB(Trim$(Player(index).Name)) > 0 Then
    If LenB(theName) > 0 Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal charNum As Long)
    Dim f As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(index).Name)) = 0 Then

        spritecheck = False

        If charNum < 1 Or charNum > MAX_CHARS Then Exit Sub
        TempPlayer(index).charNum = charNum

        Player(index).Name = Name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum

        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).dir = DIR_DOWN
        Player(index).Map = START_MAP
        Player(index).x = START_X
        Player(index).y = START_Y
        Player(index).dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)

        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If

        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Spell(n).Spell = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).Slot = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).sType = 2    ' spells
                    End If
                End If
            Next
        End If

        ' Append name to file
        f = FreeFile
        Open App.Path & "\data\accounts\_charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Shop(shopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Shop(i)
        Close #f
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist(App.Path & "\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\spells\spells" & spellNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Spell(spellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist(App.Path & "\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1    'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Resource(ResourceNum)
    Close #f
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long

    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Resource(i)
        Close #f
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist(App.Path & "\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
    Case HP
        With Class(ClassNum)
            GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
        End With
    Case MP
        With Class(ClassNum)
            GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
        End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub
