VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6480
      TabIndex        =   56
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   7560
      TabIndex        =   55
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   7695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spell Properties"
      Height          =   7695
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.Frame Frame2 
         Caption         =   "Basic Information"
         Height          =   5655
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   4725
            Width           =   3015
         End
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   44
            Top             =   5040
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   5280
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   32
            Top             =   4185
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   30
            Top             =   3680
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3120
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2560
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2040
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1500
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0000
            Left            =   120
            List            =   "frmEditor_Spell.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   4485
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   5040
            Width           =   3015
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   3975
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   3440
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Class Required:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            Caption         =   "Access Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2320
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level Required: None"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblMP 
            Caption         =   "MP Cost: None"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1280
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   745
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.HScrollBar scrlUses 
         Height          =   255
         LargeChange     =   10
         Left            =   3480
         TabIndex        =   54
         Top             =   7110
         Width           =   3255
      End
      Begin VB.HScrollBar scrlNext 
         Height          =   255
         Left            =   3480
         TabIndex        =   52
         Top             =   6630
         Width           =   3255
      End
      Begin VB.HScrollBar scrlIndex 
         Height          =   255
         Left            =   3480
         TabIndex        =   50
         Top             =   6165
         Width           =   3255
      End
      Begin VB.TextBox txtDesc 
         Height          =   855
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   6120
         Width           =   3255
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   7200
         Width           =   3255
      End
      Begin VB.Frame fraProjectile 
         Caption         =   "Projectile"
         Height          =   5775
         Left            =   3480
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CheckBox chkTrap 
            Caption         =   "Is Trap?"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlDurationProjectile 
            Height          =   255
            LargeChange     =   100
            Left            =   120
            SmallChange     =   50
            TabIndex        =   96
            Top             =   3240
            Width           =   3015
         End
         Begin VB.HScrollBar scrlProjectileSpeed 
            Height          =   255
            LargeChange     =   100
            Left            =   120
            SmallChange     =   50
            TabIndex        =   95
            Top             =   630
            Width           =   1695
         End
         Begin VB.HScrollBar scrlDamageProjectile 
            Height          =   255
            LargeChange     =   5
            Left            =   120
            TabIndex        =   94
            Top             =   1110
            Width           =   1695
         End
         Begin VB.CheckBox chkRecuringDamage 
            Caption         =   "Chain Damage"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1400
            Width           =   1455
         End
         Begin VB.CheckBox chkProjectileAoE 
            Caption         =   "Damage AoE"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   3990
            Width           =   1335
         End
         Begin VB.HScrollBar scrlProjectileRadiusY 
            Height          =   255
            Left            =   1680
            TabIndex        =   71
            Top             =   5400
            Width           =   1455
         End
         Begin VB.HScrollBar scrlProjectileRadiusX 
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   5400
            Width           =   1455
         End
         Begin VB.CheckBox chkDirectionalProjectile 
            Caption         =   "Directional?"
            Height          =   255
            Left            =   1920
            TabIndex        =   69
            Top             =   1800
            Width           =   1215
         End
         Begin VB.ComboBox cmbDirection 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0054
            Left            =   120
            List            =   "frmEditor_Spell.frx":0067
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   4260
            Width           =   3015
         End
         Begin VB.HScrollBar scrlOffsetProjectileX 
            Height          =   255
            Left            =   120
            Max             =   512
            Min             =   -512
            TabIndex        =   67
            Top             =   4860
            Width           =   1455
         End
         Begin VB.HScrollBar scrlOffsetProjectileY 
            Height          =   255
            Left            =   1680
            Max             =   512
            Min             =   -512
            TabIndex        =   66
            Top             =   4860
            Width           =   1455
         End
         Begin VB.HScrollBar scrlProjectileAnimOnHit 
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlProjectilePic 
            Height          =   255
            Left            =   2040
            TabIndex        =   64
            Top             =   1440
            Width           =   1095
         End
         Begin VB.HScrollBar scrlProjectileRange 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   63
            Top             =   1835
            Width           =   1695
         End
         Begin VB.HScrollBar scrlProjectileRotation 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   100
            TabIndex        =   62
            Top             =   2315
            Value           =   1
            Width           =   1695
         End
         Begin VB.HScrollBar scrlProjectileAmmo 
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   2760
            Width           =   1695
         End
         Begin VB.PictureBox picProjectile 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   2090
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   60
            Top             =   240
            Width           =   960
         End
         Begin VB.Label lblAoEProjectile 
            Caption         =   "No AoE damage settings"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblX0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X Offset: 0"
            Height          =   180
            Left            =   300
            TabIndex        =   83
            Top             =   4620
            Width           =   825
         End
         Begin VB.Label lblY0 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y Offset: 0"
            Height          =   180
            Left            =   1800
            TabIndex        =   82
            Top             =   4620
            Width           =   825
         End
         Begin VB.Label lblDamageProjectile 
            BackStyle       =   0  'Transparent
            Caption         =   "Base Damage: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   900
            Width           =   1575
         End
         Begin VB.Label lblAnimOnHit 
            Caption         =   "Anim on hit: None"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   3495
            Width           =   3015
         End
         Begin VB.Label lblProjectileDuration 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duration: 0 (seg)"
            Height          =   180
            Left            =   120
            TabIndex        =   79
            Top             =   3000
            Width           =   1290
         End
         Begin VB.Label lblProjectileSpeed 
            BackStyle       =   0  'Transparent
            Caption         =   "Speed: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label lblProjectilePic 
            Alignment       =   2  'Center
            Caption         =   "Pic: 0"
            Height          =   255
            Left            =   2040
            TabIndex        =   77
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label lblProjectileRange 
            Caption         =   "Range: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1635
            Width           =   1695
         End
         Begin VB.Label lblProjectileRotation 
            Caption         =   "Rotation Projectile: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   2100
            Width           =   1815
         End
         Begin VB.Label lblProjectileAmmo 
            Caption         =   "Item: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   2570
            Width           =   1335
         End
      End
      Begin VB.Frame fraSpellData 
         Caption         =   "Data"
         Height          =   5775
         Left            =   3480
         TabIndex        =   14
         Top             =   135
         Width           =   3255
         Begin VB.HScrollBar scrlAnim 
            Height          =   270
            Left            =   120
            TabIndex        =   92
            Top             =   5400
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   4800
            Width           =   3015
         End
         Begin VB.ComboBox cmbAoEDirection 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":0088
            Left            =   120
            List            =   "frmEditor_Spell.frx":009B
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   3600
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRadiusY 
            Height          =   255
            Left            =   1680
            TabIndex        =   87
            Top             =   4200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlRadiusX 
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   4200
            Width           =   1455
         End
         Begin VB.CheckBox chkDirectional 
            Caption         =   "Directional"
            Height          =   255
            Left            =   1920
            TabIndex        =   85
            Top             =   3240
            Width           =   1215
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "AoE spell?"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3240
            Width           =   1215
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   38
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   36
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlVital 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   1000
            TabIndex        =   34
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   4560
            Width           =   3015
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   3960
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            Caption         =   "Interval: 0s"
            Height          =   255
            Left            =   1680
            TabIndex        =   37
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            Caption         =   "Duration: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVital 
            Caption         =   "Vital: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label lblUses 
         Caption         =   "Uses: 0"
         Height          =   255
         Left            =   3480
         TabIndex        =   53
         Top             =   6915
         Width           =   3135
      End
      Begin VB.Label lblNext 
         Caption         =   "Next: None"
         Height          =   255
         Left            =   3480
         TabIndex        =   51
         Top             =   6435
         Width           =   3255
      End
      Begin VB.Label lblIndex 
         Caption         =   "Unique Index: 0"
         Height          =   255
         Left            =   3480
         TabIndex        =   49
         Top             =   5955
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   6975
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   5880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()

    If chkAOE.Value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If

End Sub

Private Sub chkDirectional_Click()
    If chkDirectional.Value = 0 Then
        Spell(EditorIndex).IsDirectional = False
        Me.cmbAoEDirection.enabled = False
    Else
        Spell(EditorIndex).IsDirectional = True
        Me.cmbAoEDirection.enabled = True
    End If
End Sub

Private Sub chkDirectionalProjectile_Click()
    If chkDirectionalProjectile.Value = 0 Then
        Spell(EditorIndex).IsDirectional = False
    Else
        Spell(EditorIndex).IsDirectional = True
    End If
End Sub

Private Sub chkProjectileAoE_Click()
    If chkProjectileAoE.Value = 0 Then
        Spell(EditorIndex).IsAoE = False
        scrlProjectileRadiusX.enabled = False
        scrlProjectileRadiusY.enabled = False
    Else
        scrlProjectileRadiusX.enabled = True
        scrlProjectileRadiusY.enabled = True
        Spell(EditorIndex).IsAoE = True
    End If
End Sub

Private Sub chkRecuringDamage_Click()
    If chkRecuringDamage.Value = 0 Then
        Spell(EditorIndex).Projectile.RecuringDamage = False
    Else
        Spell(EditorIndex).Projectile.RecuringDamage = True
    End If
End Sub

Private Sub chkTrap_Click()
    Call IsTrap
End Sub

Private Sub cmbAoEDirection_Click()
    If cmbAoEDirection.ListIndex < 1 Then
        scrlRadiusX.Value = Spell(EditorIndex).RadiusX
        scrlRadiusY.Value = Spell(EditorIndex).RadiusY
        Exit Sub
    End If

    scrlRadiusX.Value = Spell(EditorIndex).DirectionAoE(cmbAoEDirection.ListIndex).X
    scrlRadiusY.Value = Spell(EditorIndex).DirectionAoE(cmbAoEDirection.ListIndex).Y

    scrlRadiusX_Change
    scrlRadiusY_Change
End Sub

Private Sub cmbClass_Click()
    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
End Sub

Private Sub cmbDirection_Click()
    If cmbDirection.ListIndex < 1 Then
        scrlOffsetProjectileX.Value = 0
        scrlOffsetProjectileY.Value = 0
        lblX0.caption = "X Offset: 0"
        lblY0.caption = "Y Offset: 0"
        scrlOffsetProjectileX.enabled = False
        scrlOffsetProjectileY.enabled = False
        
        scrlProjectileRadiusX.Value = 0
        scrlProjectileRadiusY.Value = 0
        lblAoEProjectile.caption = "Sem configurações de dano AoE"
        scrlProjectileRadiusX.enabled = False
        scrlProjectileRadiusY.enabled = False

        Exit Sub

    End If

    If fraProjectile.visible = False Then Exit Sub

    scrlOffsetProjectileX.enabled = True
    scrlOffsetProjectileY.enabled = True
    scrlOffsetProjectileX.Value = Spell(EditorIndex).Projectile.ProjectileOffset(cmbDirection.ListIndex).X
    scrlOffsetProjectileY.Value = Spell(EditorIndex).Projectile.ProjectileOffset(cmbDirection.ListIndex).Y
    
    scrlOffsetProjectileX_Change
    scrlOffsetProjectileY_Change
    
    If chkProjectileAoE.Value = 1 Then
        scrlProjectileRadiusX.enabled = True
        scrlProjectileRadiusY.enabled = True
        scrlProjectileRadiusX.Value = Spell(EditorIndex).DirectionAoE(cmbDirection.ListIndex).X
        scrlProjectileRadiusY.Value = Spell(EditorIndex).DirectionAoE(cmbDirection.ListIndex).Y
        
        scrlProjectileRadiusX_Change
        scrlProjectileRadiusY_Change
    End If
End Sub

Private Sub cmbType_Click()
    Spell(EditorIndex).Type = cmbType.ListIndex
    If Spell(EditorIndex).Type = SPELL_TYPE_PROJECTILE Then
        fraProjectile.visible = True
        fraSpellData.visible = False
        scrlProjectilePic.max = CountProjectile
    Else
        fraProjectile.visible = False
        fraSpellData.visible = True
    End If
End Sub

Private Sub cmdCopy_Click()
    SpellEditorCopy
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    ClearSpell EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    SpellEditorInit
End Sub

Private Sub cmdPaste_Click()
    SpellEditorPaste
End Sub

Private Sub cmdSave_Click()
    SpellEditorOk
End Sub


Private Sub lstIndex_Click()
    SpellEditorInit
End Sub

Private Sub cmdCancel_Click()
    SpellEditorCancel
End Sub

Private Sub scrlAccess_Change()

    If scrlAccess.Value > 0 Then
        lblAccess.caption = "Access Required: " & scrlAccess.Value
    Else
        lblAccess.caption = "Access Required: None"
    End If

    Spell(EditorIndex).AccessReq = scrlAccess.Value
End Sub

Private Sub scrlAnim_Change()

    If scrlAnim.Value > 0 Then
        lblAnim.caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.caption = "Animation: None"
    End If

    Spell(EditorIndex).SpellAnim = scrlAnim.Value
End Sub

Private Sub scrlAnimCast_Change()

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.caption = "Cast Anim: None"
    End If

    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
End Sub

Private Sub scrlAOE_Click()
    If cmbAoEDirection.ListIndex < 1 Then
        scrlRadiusX.Value = Spell(EditorIndex).RadiusX
        scrlRadiusY.Value = Spell(EditorIndex).RadiusY

        Exit Sub

    End If

    scrlRadiusX.Value = Spell(EditorIndex).DirectionAoE(cmbAoEDirection.ListIndex).X
    scrlRadiusY.Value = Spell(EditorIndex).DirectionAoE(cmbAoEDirection.ListIndex).Y

    scrlRadiusX_Change
    scrlRadiusY_Change
End Sub

Private Sub scrlCast_Change()
    lblCast.caption = "Casting Time: " & scrlCast.Value & "s"
    Spell(EditorIndex).CastTime = scrlCast.Value
End Sub

Private Sub scrlCool_Change()
    lblCool.caption = "Cooldown Time: " & scrlCool.Value & "s"
    Spell(EditorIndex).CDTime = scrlCool.Value
End Sub

Private Sub scrlDamageProjectile_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    
    If scrlDamageProjectile.Value > 0 Then
        lblDamageProjectile.caption = "Base Damage: " & scrlDamageProjectile.Value
    Else
        lblDamageProjectile.caption = "Base Damage: 0"
    End If
    
    Spell(EditorIndex).Vital = scrlDamageProjectile.Value
End Sub

Private Sub scrlDir_Change()
    Dim sDir As String

    Select Case scrlDir.Value

        Case DIR_UP
            sDir = "Up"

        Case DIR_DOWN
            sDir = "Down"

        Case DIR_RIGHT
            sDir = "Right"

        Case DIR_LEFT
            sDir = "Left"
    End Select

    lblDir.caption = "Dir: " & sDir
    Spell(EditorIndex).Dir = scrlDir.Value
End Sub

Private Sub scrlDuration_Change()
    lblDuration.caption = "Duration: " & scrlDuration.Value & "s"
    Spell(EditorIndex).Duration = scrlDuration.Value
End Sub

Private Sub scrlDurationProjectile_Change()
    Dim DurationText As String, Duration As Long
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    Duration = scrlDurationProjectile.Value
    If Duration > 0 Then
        If Int(Duration * 100) / 1000 < 60 Then
            lblProjectileDuration.caption = "Duration: " & (Duration * 100) / 1000 & " (seg)"
        Else
            lblProjectileDuration.caption = "Duration: " & (Duration * 100) / 60000 & " (min)"
        End If
    Else
        lblProjectileDuration.caption = "Duration: 0 seg"
    End If
    
    Spell(EditorIndex).Projectile.Despawn = (Duration * 100)
End Sub

Private Sub scrlIcon_Change()

    If scrlIcon.Value > 0 Then
        lblIcon.caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.caption = "Icon: None"
    End If

    Spell(EditorIndex).icon = scrlIcon.Value
End Sub

Private Sub scrlIndex_Change()
    lblIndex.caption = "Unique Index: " & scrlIndex.Value
    Spell(EditorIndex).UniqueIndex = scrlIndex.Value
End Sub

Private Sub scrlInterval_Change()
    lblInterval.caption = "Interval: " & scrlInterval.Value & "s"
    Spell(EditorIndex).Interval = scrlInterval.Value
End Sub

Private Sub scrlLevel_Change()

    If scrlLevel.Value > 0 Then
        lblLevel.caption = "Level Required: " & scrlLevel.Value
    Else
        lblLevel.caption = "Level Required: None"
    End If

    Spell(EditorIndex).LevelReq = scrlLevel.Value
End Sub

Private Sub scrlMap_Change()
    lblMap.caption = "Map: " & scrlMap.Value
    Spell(EditorIndex).Map = scrlMap.Value
End Sub

Private Sub scrlMP_Change()

    If scrlMP.Value > 0 Then
        lblMP.caption = "MP Cost: " & scrlMP.Value
    Else
        lblMP.caption = "MP Cost: None"
    End If

    Spell(EditorIndex).MPCost = scrlMP.Value
End Sub

Private Sub scrlNext_Change()

    If scrlNext.Value > 0 Then
        lblNext.caption = "Next: " & scrlNext.Value & " - " & Trim$(Spell(scrlNext.Value).Name)
    Else
        lblNext.caption = "Next: None"
    End If

    Spell(EditorIndex).NextRank = scrlNext.Value
End Sub

Private Sub scrlOffsetProjectileX_Change()
    If cmbDirection.ListIndex < 1 Then Exit Sub
    If fraProjectile.visible = False Then Exit Sub

    lblX0.caption = "X Offset: " & scrlOffsetProjectileX.Value
    Spell(EditorIndex).Projectile.ProjectileOffset(cmbDirection.ListIndex).X = scrlOffsetProjectileX.Value
End Sub

Private Sub scrlOffsetProjectileY_Change()
    If cmbDirection.ListIndex < 1 Then Exit Sub
    If fraProjectile.visible = False Then Exit Sub

    lblY0.caption = "Y Offset: " & scrlOffsetProjectileY.Value
    Spell(EditorIndex).Projectile.ProjectileOffset(cmbDirection.ListIndex).Y = scrlOffsetProjectileY.Value
End Sub

Private Sub scrlProjectileAmmo_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    lblProjectileAmmo.caption = "Item: " & scrlProjectileAmmo.Value
    Spell(EditorIndex).Projectile.Ammo = scrlProjectileAmmo.Value
End Sub

Private Sub scrlProjectileAnimOnHit_Change()
    If scrlProjectileAnimOnHit.Value > 0 Then
        lblAnimOnHit.caption = "Anim on hit: " & Trim$(Animation(scrlProjectileAnimOnHit.Value).Name)
    Else
        lblAnimOnHit.caption = "Anim on hit: None"
    End If

    Spell(EditorIndex).Projectile.AnimOnHit = scrlProjectileAnimOnHit.Value
End Sub

Private Sub scrlProjectilePic_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    lblProjectilePic.caption = "Pic: " & scrlProjectilePic.Value
    Spell(EditorIndex).Projectile.Graphic = scrlProjectilePic.Value
End Sub

Private Sub scrlProjectileRadiusX_Change()
    If cmbDirection.ListIndex = 0 Then
        If scrlProjectileRadiusX.Value > 0 Or scrlRadiusY.Value > 0 Then
            lblAoEProjectile.caption = "Radius X: " & scrlProjectileRadiusX.Value & " Radius Y: " & scrlProjectileRadiusY.Value & " tiles."
        Else
            lblAoEProjectile.caption = "Damage Default"
        End If
    Else
        lblAoEProjectile.caption = cmbDirection.list(cmbDirection.ListIndex) & " Radius X: " & scrlProjectileRadiusX.Value & " and Radius Y: " & scrlProjectileRadiusY.Value
        Spell(EditorIndex).DirectionAoE(cmbDirection.ListIndex).X = scrlProjectileRadiusX.Value
    End If
End Sub

Private Sub scrlProjectileRadiusY_Change()
    If cmbDirection.ListIndex = 0 Then
        If scrlProjectileRadiusX.Value > 0 Or scrlRadiusY.Value > 0 Then
            lblAoEProjectile.caption = "Radius X: " & scrlProjectileRadiusX.Value & " Radius Y: " & scrlProjectileRadiusY.Value & " tiles."
        Else
            lblAoEProjectile.caption = "Damage Default"
        End If
    Else
        lblAoEProjectile.caption = cmbDirection.list(cmbDirection.ListIndex) & " Radius X: " & scrlProjectileRadiusX.Value & " and Radius Y: " & scrlProjectileRadiusY.Value
        Spell(EditorIndex).DirectionAoE(cmbDirection.ListIndex).Y = scrlProjectileRadiusY.Value
    End If
End Sub

Private Sub scrlProjectileRange_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    
    If scrlProjectileRange.Value > 0 Then
        lblProjectileRange.caption = "Range: " & scrlProjectileRange.Value & " tiles."
    Else
        lblProjectileRange.caption = "Range: 0"
    End If
    
    Spell(EditorIndex).Range = scrlProjectileRange.Value
End Sub

Private Sub scrlProjectileRotation_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    lblProjectileRotation.caption = "Rotation Projectile: " & scrlProjectileRotation.Value / 2
    Spell(EditorIndex).Projectile.Rotation = scrlProjectileRotation.Value
End Sub

Private Sub scrlProjectileSpeed_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    
    If scrlProjectileSpeed.Value > 0 Then
        lblProjectileSpeed.caption = "Speed: " & scrlProjectileSpeed.Value
    Else
        lblProjectileSpeed.caption = "Speed: 0"
    End If
    
    Spell(EditorIndex).Projectile.Speed = scrlProjectileSpeed.Value
End Sub

Private Sub scrlRadiusX_Change()
    If cmbAoEDirection.ListIndex = 0 Then
        If scrlRadiusX.Value > 0 Or scrlRadiusY.Value > 0 Then
            lblAOE.caption = "Radius X: " & scrlRadiusX.Value & " Radius Y: " & scrlRadiusY.Value & " tiles."
        Else
            lblAOE.caption = "AoE: Self-cast"
        End If

        Spell(EditorIndex).RadiusX = scrlRadiusX.Value
    Else
        lblAOE.caption = cmbAoEDirection.list(cmbAoEDirection.ListIndex) & " Radius X: " & scrlRadiusX.Value & " and Radius Y: " & scrlRadiusY.Value
        Spell(EditorIndex).DirectionAoE(cmbAoEDirection.ListIndex).X = scrlRadiusX.Value
    End If
End Sub

Private Sub scrlRadiusY_Change()
    If cmbAoEDirection.ListIndex = 0 Then
        If scrlRadiusX.Value > 0 Or scrlRadiusY.Value > 0 Then
            lblAOE.caption = "Radius X: " & scrlRadiusX.Value & " Radius Y: " & scrlRadiusY.Value & " tiles."
        Else
            lblAOE.caption = "AoE: Self-cast"
        End If

        Spell(EditorIndex).RadiusY = scrlRadiusY.Value
    Else
        lblAOE.caption = cmbAoEDirection.list(cmbAoEDirection.ListIndex) & " Radius X: " & scrlRadiusX.Value & " and Radius Y: " & scrlRadiusY.Value
        Spell(EditorIndex).DirectionAoE(cmbAoEDirection.ListIndex).Y = scrlRadiusY.Value
    End If
End Sub

Private Sub scrlRange_Change()

    If scrlRange.Value > 0 Then
        lblRange.caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.caption = "Range: Self-cast"
    End If

    Spell(EditorIndex).Range = scrlRange.Value
End Sub

Private Sub scrlStun_Change()

    If scrlStun.Value > 0 Then
        lblStun.caption = "Stun Duration: " & scrlStun.Value & "s"
    Else
        lblStun.caption = "Stun Duration: None"
    End If

    Spell(EditorIndex).StunDuration = scrlStun.Value
End Sub

Private Sub scrlUses_Change()
    lblUses.caption = "Uses: " & scrlUses.Value
    Spell(EditorIndex).NextUses = scrlUses.Value
End Sub

Private Sub scrlVital_Change()
    lblVital.caption = "Vital: " & scrlVital.Value
    Spell(EditorIndex).Vital = scrlVital.Value
End Sub

Private Sub scrlX_Change()
    lblX.caption = "X: " & scrlX.Value
    Spell(EditorIndex).X = scrlX.Value
End Sub

Private Sub scrlY_Change()
    lblY.caption = "Y: " & scrlY.Value
    Spell(EditorIndex).Y = scrlY.Value
End Sub

Private Sub txtDesc_Change()
    Spell(EditorIndex).Desc = txtDesc.text
End Sub

Public Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub cmbSound_Click()
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).sound = cmbSound.list(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).sound = "None."
    End If

End Sub
