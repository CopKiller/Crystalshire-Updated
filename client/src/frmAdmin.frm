VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Painel Adminitrativo"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Lucida Sans"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Editor"
      Height          =   4935
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   2175
      Begin VB.CommandButton CmdQuest 
         Caption         =   "Quest's"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton CmdConv 
         Caption         =   "Conversations"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton CmdResource 
         Caption         =   "Resources"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton CmdAnimation 
         Caption         =   "Animation"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton CmdShop 
         Caption         =   "Shops"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton CmdNpc 
         Caption         =   "Npc's"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton CmdSpell 
         Caption         =   "Spell"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton CmdItem 
         Caption         =   "Itens"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton CmdMap 
         Caption         =   "Map"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Administrator"
      Height          =   4935
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp to map"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Warp to me"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kickar"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "Warp me to"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdAtt 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Level UP"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtAmount 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "1"
         Top             =   3960
         Width           =   2055
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Drop"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   1
         Top             =   3360
         Value           =   1
         Width           =   2055
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawn Item: None"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdABan_Click()
    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
End Sub

Private Sub cmdAKick_Click()
    If Len(Trim$(txtAName.text)) < 1 Then Exit Sub

    SendKick Trim$(txtAName.text)
End Sub

Private Sub CmdAnimation_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditAnimation
End Sub

Private Sub cmdASpawn_Click()

    If Len(txtAmount.text) = 0 Then Exit Sub
    If txtAmount.text = 0 Then Exit Sub
    
    If scrlAItem.value > 0 Then
        SendSpawnItem scrlAItem.value, Trim$(txtAmount.text)
    End If
End Sub

Private Sub cmdASprite_Click()

    If Len(Trim$(txtASprite.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtASprite.text)) Then
        Exit Sub
    End If

    SendSetSprite CLng(Trim$(txtASprite.text))

    Exit Sub
End Sub

Private Sub cmdAtt_Click()
    SendMapRespawn
End Sub

Private Sub cmdAWarp_Click()
    Dim N As Long

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then

        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    N = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If N > 0 And N <= MAX_MAPS Then
        Call WarpTo(N)
    Else
        Call AddText("Invalid map number.", Red)
    End If

    ' Error handler
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
End Sub

Private Sub cmdAWarpMe2_Click()
    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
End Sub

Private Sub CmdConv_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditConv
End Sub

Private Sub CmdItem_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditItem
End Sub

Private Sub cmdLevel_Click()
    SendRequestLevelUp
End Sub

Private Sub CmdMap_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendRequestEditMap
End Sub

Private Sub CmdNpc_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditNpc
End Sub

Private Sub CmdQuest_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditMission
End Sub

Private Sub CmdResource_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditResource
End Sub

Private Sub cmdShop_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditShop
End Sub

Private Sub CmdSpell_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    SendRequestEditSpell
End Sub

Private Sub scrlAItem_Change()
    If scrlAItem.value > 0 Then
        lblAItem.caption = " " & Trim$(Item(scrlAItem.value).Name)
    Else
        lblAItem.caption = "Item: None"
    End If
End Sub

