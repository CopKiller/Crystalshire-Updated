VERSION 5.00
Begin VB.Form frmEditor_Conv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversation Editor"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
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
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   23
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   22
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Frame fraConv 
      Caption         =   "Conversation - 1"
      Height          =   6495
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   4215
      Begin VB.ComboBox cmbEventNum 
         Height          =   315
         ItemData        =   "frmEditor_Conv.frx":0000
         Left            =   120
         List            =   "frmEditor_Conv.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   5640
         Width           =   3975
      End
      Begin VB.ComboBox cmbEvent 
         Height          =   315
         ItemData        =   "frmEditor_Conv.frx":0004
         Left            =   120
         List            =   "frmEditor_Conv.frx":0014
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   5040
         Width           =   3975
      End
      Begin VB.HScrollBar scrlConv 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   20
         Top             =   240
         Value           =   1
         Width           =   3975
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   4
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   4335
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   4350
         Width           =   2775
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   3
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3975
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   3990
         Width           =   2775
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   2
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3615
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   3630
         Width           =   2775
      End
      Begin VB.ComboBox cmbReply 
         Height          =   315
         Index           =   1
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3225
         Width           =   1095
      End
      Begin VB.TextBox txtReply 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox txtConv 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Event Num:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Event Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Replies:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   975
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.HScrollBar scrlChatCount 
         Height          =   255
         Left            =   1680
         Max             =   100
         Min             =   1
         TabIndex        =   19
         Top             =   600
         Value           =   1
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblChatCount 
         AutoSize        =   -1  'True
         Caption         =   "Chat Count: 1"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Conversation List"
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7800
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Conv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curConv As Long
Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    ClearConv EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conv(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    ConvEditorInit
End Sub

Private Sub cmdSave_Click()
    Call ConvEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ConvEditorCancel
End Sub

Private Sub Form_Load()
    cmbEvent.ListIndex = 0
End Sub

Private Sub lstIndex_Click()
    Call ConvEditorInit
End Sub

Private Sub scrlChatCountChange()
    Dim N As Long, i As Long
    
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    If curConv <= 0 Then Exit Sub
    
    lblChatCount.caption = "Chat Count: " & scrlChatCount.Value
    Conv(EditorIndex).chatCount = scrlChatCount.Value
    scrlConv.max = scrlChatCount.Value
    ReDim Preserve Conv(EditorIndex).Conv(1 To scrlChatCount.Value) As ConvRec
    
    For N = 1 To 4
        cmbReply(N).Clear
        cmbReply(N).AddItem "None"

        For i = 1 To Conv(EditorIndex).chatCount
            cmbReply(N).AddItem i
        Next
        
        With Conv(EditorIndex).Conv(curConv)
            txtConv.text = .Conv
    
            For i = 1 To 4
                txtReply(i).text = .rText(i)
                cmbReply(i).ListIndex = .rTarget(i)
            Next
        End With

    Next
End Sub

Private Sub scrlConv_Change()
    Dim X As Long
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    
    curConv = scrlConv.Value
    fraConv.caption = "Conversation - " & curConv

    With Conv(EditorIndex).Conv(curConv)
        txtConv.text = .Conv

        For X = 1 To 4
            txtReply(X).text = .rText(X)
            cmbReply(X).ListIndex = .rTarget(X)
        Next
    End With

End Sub

Private Sub txtConv_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    Conv(EditorIndex).Conv(curConv).Conv = txtConv.text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Conv(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Conv(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtReply_Change(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    Conv(EditorIndex).Conv(curConv).rText(Index) = txtReply(Index).text
End Sub

Private Sub cmbReply_Click(Index As Integer)
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    Conv(EditorIndex).Conv(curConv).rTarget(Index) = cmbReply(Index).ListIndex
End Sub

Private Sub cmbEvent_Click()
    Dim i As Long
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    If curConv = 0 Then Exit Sub
    If (cmbEvent.ListIndex = EventType.Event_OpenShop) Then
        cmbEventNum.visible = True
        ' build EventNum combo
        cmbEventNum.Clear
        cmbEventNum.AddItem "None"

        For i = 1 To MAX_SHOPS
            cmbEventNum.AddItem Trim$(Shop(i).Name)
        Next

        cmbEventNum.ListIndex = 0
    ElseIf (cmbEvent.ListIndex = EventType.Event_OpenQuest) Then
        cmbEventNum.visible = True
        ' build EventNum combo
        cmbEventNum.Clear
        cmbEventNum.AddItem "None"

        For i = 1 To MAX_MISSIONS
            cmbEventNum.AddItem Trim$(Mission(i).Name)
        Next

        cmbEventNum.ListIndex = 0
    Else
        cmbEventNum.visible = False
    End If
    
    Conv(EditorIndex).Conv(curConv).EventType = cmbEvent.ListIndex
End Sub

Private Sub cmbEventNum_Click()
    If EditorIndex = 0 Or EditorIndex > MAX_CONVS Then Exit Sub
    Conv(EditorIndex).Conv(curConv).EventNum = cmbEventNum.ListIndex
End Sub
