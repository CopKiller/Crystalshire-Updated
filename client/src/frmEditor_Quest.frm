VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quest Editor"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8985
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   9615
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   28
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   27
      Top             =   9120
      Width           =   1455
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   9120
      Width           =   2895
   End
   Begin VB.Frame frmKillQuest 
      Caption         =   "Kill Quest"
      Height          =   1335
      Left            =   3360
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
      Begin VB.HScrollBar scrlKillAmount 
         Height          =   255
         LargeChange     =   5
         Left            =   2280
         Max             =   50
         TabIndex        =   33
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmbKillNPC 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblKillAmount 
         Caption         =   "Kill Amount:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Kill NPC:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame frmQuestVars 
      Caption         =   "Main Quest Variables"
      Height          =   2895
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   5535
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   1440
         List            =   "frmEditor_Quest.frx":000D
         TabIndex        =   11
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtDialogue 
         Height          =   1215
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1560
         Width           =   3855
      End
      Begin VB.OptionButton optRepeatableNo 
         Caption         =   "No"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optRepeatableYes 
         Caption         =   "Yes"
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Quest Type:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Repeatable:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Quest Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quest List"
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   8640
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Chain"
      Height          =   855
      Left            =   3360
      TabIndex        =   38
      Top             =   3000
      Width           =   5535
      Begin VB.ComboBox cmbPreviousQuest 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0026
         Left            =   1680
         List            =   "frmEditor_Quest.frx":0028
         TabIndex        =   40
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label9 
         Caption         =   "Previous Quest:"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frmTalkQuest 
      Caption         =   "Talk Quest"
      Height          =   1335
      Left            =   3360
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox cmbTalkNPC 
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use the 'Completed' dialogue for your finish quest chatter."
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label6 
         Caption         =   "NPC to Talk to:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmCollectQuest 
      Caption         =   "Collect Quest"
      Height          =   1335
      Left            =   3360
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   5535
      Begin VB.HScrollBar scrlCollectAmount 
         Height          =   255
         Left            =   2280
         Max             =   200
         TabIndex        =   31
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmbCollectItem 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblCollectAmount 
         Caption         =   "Collect Amount: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Collect Item:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Messages"
      Height          =   1215
      Left            =   3360
      TabIndex        =   21
      Top             =   5400
      Width           =   5535
      Begin VB.TextBox txtCompleted 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtIncomplete 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label13 
         Caption         =   "Completed:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Incomplete:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quest Reward"
      Height          =   2295
      Left            =   3360
      TabIndex        =   20
      Top             =   6720
      Width           =   5535
      Begin VB.HScrollBar scrlItemAmount 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   32000
         TabIndex        =   46
         Top             =   1320
         Width           =   5295
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   45
         Top             =   600
         Width           =   480
      End
      Begin VB.HScrollBar scrlItemNum 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   0
         TabIndex        =   43
         Top             =   840
         Width           =   4695
      End
      Begin VB.HScrollBar scrlRewardNum 
         Height          =   255
         Left            =   1200
         Max             =   5
         Min             =   1
         TabIndex        =   41
         Top             =   240
         Value           =   1
         Width           =   4215
      End
      Begin VB.HScrollBar scrlRewardExperience 
         Height          =   255
         LargeChange     =   100
         Left            =   3360
         Max             =   10000
         TabIndex        =   32
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5400
         Y1              =   1750
         Y2              =   1750
      End
      Begin VB.Label lblItemAmount 
         Caption         =   "Amount: 10"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1090
         Width           =   1935
      End
      Begin VB.Label lblItemName 
         Caption         =   "Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblRewardNum 
         Caption         =   "Reward (1):"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   250
         Width           =   1935
      End
      Begin VB.Label lblRewardExperience 
         Caption         =   "Reward Experience: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCollectItem_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).CollectItem = cmbCollectItem.ListIndex
End Sub

Private Sub cmbKillNPC_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).KillNPC = cmbKillNPC.ListIndex
End Sub

Private Sub cmbPreviousQuest_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).PreviousMissionComplete = cmbPreviousQuest.ListIndex
End Sub

Private Sub cmbTalkNPC_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).TalkNPC = cmbTalkNPC.ListIndex
End Sub

Private Sub cmbType_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).Type = cmbType.ListIndex
    If Mission(EditorIndex).Type = MissionType.Mission_TypeKill Then
        frmKillQuest.visible = True
        frmCollectQuest.visible = False
        frmTalkQuest.visible = False
    End If
    If Mission(EditorIndex).Type = MissionType.Mission_TypeCollect Then
        frmKillQuest.visible = False
        frmCollectQuest.visible = True
        frmTalkQuest.visible = False
    End If
    If Mission(EditorIndex).Type = MissionType.Mission_TypeTalk Then
        frmKillQuest.visible = False
        frmCollectQuest.visible = True
        frmTalkQuest.visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    MissionEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    
    ClearMission EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Mission(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub cmdSave_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    'If it's a kill quest, we want to set null values for a collect quest
    If cmbType.ListIndex = MissionType.Mission_TypeKill Then
        Mission(EditorIndex).CollectItem = 0
        Mission(EditorIndex).CollectItemAmount = 0
        Mission(EditorIndex).TalkNPC = 0
    ElseIf cmbType.ListIndex = MissionType.Mission_TypeCollect Then
        Mission(EditorIndex).KillNPC = 0
        Mission(EditorIndex).KillNPCAmount = 0
        Mission(EditorIndex).TalkNPC = 0
    ElseIf cmbType.ListIndex = MissionType.Mission_TypeTalk Then
        Mission(EditorIndex).KillNPC = 0
        Mission(EditorIndex).KillNPCAmount = 0
        Mission(EditorIndex).CollectItem = 0
        Mission(EditorIndex).CollectItemAmount = 0
    End If
    
   'Do Save Code here
   Call MissionEditorOk
End Sub

Private Sub lstIndex_Click()
    MissionEditorInit
End Sub

Private Sub optRepeatableNo_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    optRepeatableYes.value = False
    optRepeatableNo.value = True
    Mission(EditorIndex).Repeatable = 0
End Sub

Private Sub optRepeatableYes_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    optRepeatableYes.value = True
    optRepeatableNo.value = False
    Mission(EditorIndex).Repeatable = 1
End Sub

Private Sub txtCollectAmount_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).CollectItemAmount = txtCollectAmount.text
End Sub

Private Sub scrlCollectAmount_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    lblCollectAmount.caption = "Collect Amount: " & scrlCollectAmount.value
    Mission(EditorIndex).CollectItemAmount = scrlCollectAmount.value
End Sub

Private Sub scrlItemAmount_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    lblItemAmount.caption = "Amount: " & scrlItemAmount.value
    ' Set
    Mission(EditorIndex).RewardItem(scrlRewardNum.value).ItemAmount = scrlItemAmount.value
End Sub

Private Sub scrlItemNum_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    
    If scrlItemNum.value <> 0 Then lblItemName.caption = "Item: " & Trim$(Item(scrlItemNum.value).Name)
    
    ' Set
    Mission(EditorIndex).RewardItem(scrlRewardNum.value).ItemNum = scrlItemNum.value
End Sub

Private Sub scrlKillAmount_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    
    lblKillAmount.caption = "Kill Amount: " & scrlKillAmount.value
    Mission(EditorIndex).KillNPCAmount = scrlKillAmount.value
End Sub

Private Sub scrlRewardExperience_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    
    lblRewardExperience.caption = "Reward Experience: " & scrlRewardExperience.value
    Mission(EditorIndex).RewardExperience = scrlRewardExperience.value
End Sub

Private Sub scrlRewardNum_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    If scrlRewardNum.value <= 0 Or scrlRewardNum.value > 5 Then Exit Sub
    
    lblRewardNum.caption = "Reward (" & scrlRewardNum.value & "):"
    
    ' Set
    scrlItemNum.value = Mission(EditorIndex).RewardItem(scrlRewardNum.value).ItemNum
    scrlItemAmount.value = Mission(EditorIndex).RewardItem(scrlRewardNum.value).ItemAmount
End Sub

Private Sub txtCompleted_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).Completed = txtCompleted.text
End Sub

Private Sub txtDialogue_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).Description = txtDialogue.text
End Sub

Private Sub txtIncomplete_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).Incomplete = txtIncomplete.text
End Sub

Private Sub txtKillAmount_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    Mission(EditorIndex).KillNPCAmount = txtKillAmount.text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_MISSIONS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Mission(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Mission(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

