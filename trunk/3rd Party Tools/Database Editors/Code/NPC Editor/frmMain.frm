VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "NPC Editor"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345.001
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   429
   ScaleMode       =   0  'User
   ScaleWidth      =   623
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   26
      Left            =   8760.001
      TabIndex        =   134
      ToolTipText     =   "Name of the object"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   27
      Left            =   8760.001
      TabIndex        =   133
      ToolTipText     =   "Name of the object"
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   28
      Left            =   8760.001
      TabIndex        =   132
      ToolTipText     =   "Name of the object"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   29
      Left            =   8760.001
      TabIndex        =   131
      ToolTipText     =   "Name of the object"
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   8760.001
      TabIndex        =   130
      ToolTipText     =   "Name of the object"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   31
      Left            =   8760.001
      TabIndex        =   129
      ToolTipText     =   "Name of the object"
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   32
      Left            =   8760.001
      TabIndex        =   128
      ToolTipText     =   "Name of the object"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   33
      Left            =   8760.001
      TabIndex        =   127
      ToolTipText     =   "Name of the object"
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   34
      Left            =   8760.001
      TabIndex        =   126
      ToolTipText     =   "Name of the object"
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   35
      Left            =   8760.001
      TabIndex        =   125
      ToolTipText     =   "Name of the object"
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   8760.001
      TabIndex        =   124
      ToolTipText     =   "Name of the object"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   24
      Left            =   8760.001
      TabIndex        =   123
      ToolTipText     =   "Name of the object"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   7560
      TabIndex        =   120
      ToolTipText     =   "Name of the object"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   7560
      TabIndex        =   119
      ToolTipText     =   "Name of the object"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox IDDropTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3240
      TabIndex        =   118
      ToolTipText     =   "Name of the object"
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox AmountDropTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      TabIndex        =   117
      ToolTipText     =   "Name of the object"
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox DropChanceTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4200
      TabIndex        =   116
      ToolTipText     =   "Name of the object"
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox ProRotSpeedTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8760.001
      TabIndex        =   114
      ToolTipText     =   "Name of the object"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox AtkSfxTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8760.001
      TabIndex        =   112
      ToolTipText     =   "Name of the object"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox AtkRngTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8760.001
      TabIndex        =   110
      ToolTipText     =   "Name of the object"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox AtkGrhTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8760.001
      TabIndex        =   108
      ToolTipText     =   "Name of the object"
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox ChatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   106
      ToolTipText     =   "Name of the object"
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox IDTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1800
      TabIndex        =   105
      ToolTipText     =   "Name of the object"
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   960
      TabIndex        =   104
      ToolTipText     =   "Name of the object"
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox RespTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   103
      ToolTipText     =   "Name of the object"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox GGoldTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   102
      ToolTipText     =   "Name of the object"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox QuestTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7320
      TabIndex        =   101
      ToolTipText     =   "Name of the object"
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox GExpTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   100
      ToolTipText     =   "Name of the object"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   23
      Left            =   7560
      TabIndex        =   98
      ToolTipText     =   "Name of the object"
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   7560
      TabIndex        =   96
      ToolTipText     =   "Name of the object"
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   7560
      TabIndex        =   94
      ToolTipText     =   "Name of the object"
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   7560
      TabIndex        =   92
      ToolTipText     =   "Name of the object"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   7560
      TabIndex        =   90
      ToolTipText     =   "Name of the object"
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   7560
      TabIndex        =   88
      ToolTipText     =   "Name of the object"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   7560
      TabIndex        =   86
      ToolTipText     =   "Name of the object"
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   7560
      TabIndex        =   84
      ToolTipText     =   "Name of the object"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   7560
      TabIndex        =   82
      ToolTipText     =   "Name of the object"
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   7560
      TabIndex        =   80
      ToolTipText     =   "Name of the object"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   6360
      TabIndex        =   78
      ToolTipText     =   "Name of the object"
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   6360
      TabIndex        =   76
      ToolTipText     =   "Name of the object"
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   6360
      TabIndex        =   74
      ToolTipText     =   "Name of the object"
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   6360
      TabIndex        =   72
      ToolTipText     =   "Name of the object"
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   6360
      TabIndex        =   70
      ToolTipText     =   "Name of the object"
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   6360
      TabIndex        =   68
      ToolTipText     =   "Name of the object"
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   6360
      TabIndex        =   66
      ToolTipText     =   "Name of the object"
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   64
      ToolTipText     =   "Name of the object"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   62
      ToolTipText     =   "Name of the object"
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   60
      ToolTipText     =   "Name of the object"
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   58
      ToolTipText     =   "Name of the object"
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6360
      TabIndex        =   56
      ToolTipText     =   "Name of the object"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   3960
      TabIndex        =   54
      ToolTipText     =   "Name of the object"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   3960
      TabIndex        =   52
      ToolTipText     =   "Name of the object"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   3960
      TabIndex        =   50
      ToolTipText     =   "Name of the object"
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   3960
      TabIndex        =   48
      ToolTipText     =   "Name of the object"
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3960
      TabIndex        =   46
      ToolTipText     =   "Name of the object"
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3960
      TabIndex        =   44
      ToolTipText     =   "Name of the object"
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   2280
      TabIndex        =   42
      ToolTipText     =   "Name of the object"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   2280
      TabIndex        =   40
      ToolTipText     =   "Name of the object"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2280
      TabIndex        =   38
      ToolTipText     =   "Name of the object"
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   36
      ToolTipText     =   "Name of the object"
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   34
      ToolTipText     =   "Name of the object"
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox CharTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2280
      TabIndex        =   32
      ToolTipText     =   "Name of the object"
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox AITxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   31
      ToolTipText     =   "Name of the object"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox DescTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5640
      TabIndex        =   30
      ToolTipText     =   "Name of the object"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5640
      TabIndex        =   28
      ToolTipText     =   "Name of the object"
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox SelectNpcCombo 
      Height          =   315
      Left            =   120
      TabIndex        =   26
      Text            =   "Select a NPC"
      Top             =   360
      Width           =   4575
   End
   Begin VB.ListBox OBJDropList 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   2175
      ItemData        =   "frmMain.frx":17D70
      Left            =   2880
      List            =   "frmMain.frx":17D72
      TabIndex        =   20
      Top             =   3480
      Width           =   2535
   End
   Begin VB.ListBox OBJList 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   2175
      ItemData        =   "frmMain.frx":17D74
      Left            =   240
      List            =   "frmMain.frx":17D76
      TabIndex        =   16
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Timer RenderTimer 
      Interval        =   50
      Left            =   8640.001
      Top             =   2520
   End
   Begin VB.CheckBox HostileChk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hostile"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.CheckBox AtkChk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Attackable"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   240
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8640.001
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label DeleteLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8520.001
      TabIndex        =   147
      Top             =   120
      Width           =   705
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   35
      Left            =   8160
      TabIndex        =   146
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   34
      Left            =   8160
      TabIndex        =   145
      Top             =   4200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   33
      Left            =   8160
      TabIndex        =   144
      Top             =   4440
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   32
      Left            =   8160
      TabIndex        =   143
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   31
      Left            =   8160
      TabIndex        =   142
      Top             =   4920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   8160
      TabIndex        =   141
      Top             =   5160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   29
      Left            =   8160
      TabIndex        =   140
      Top             =   5400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   28
      Left            =   8160
      TabIndex        =   139
      Top             =   5640
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   27
      Left            =   8160
      TabIndex        =   138
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   26
      Left            =   8160
      TabIndex        =   137
      Top             =   6120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   8160
      TabIndex        =   136
      Top             =   3720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   24
      Left            =   8160
      TabIndex        =   135
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   6960
      TabIndex        =   122
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   6960
      TabIndex        =   121
      Top             =   3720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proj. Rotate Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   7080
      TabIndex        =   115
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atk SFX:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   8040
      TabIndex        =   113
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atk Range:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   7800
      TabIndex        =   111
      Top             =   840
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atk Grh:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   8040
      TabIndex        =   109
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chat:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5040
      TabIndex        =   107
      Top             =   1800
      Width           =   465
   End
   Begin VB.Shape Shape4 
      Height          =   2415
      Left            =   4560
      Top             =   720
      Width           =   4695
   End
   Begin VB.Shape Shape3 
      Height          =   2415
      Left            =   120
      Top             =   720
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      Height          =   3135
      Left            =   120
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      Height          =   3135
      Left            =   5640
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   23
      Left            =   6960
      TabIndex        =   99
      Top             =   6120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   6960
      TabIndex        =   97
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   6960
      TabIndex        =   95
      Top             =   5640
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   6960
      TabIndex        =   93
      Top             =   5400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   6960
      TabIndex        =   91
      Top             =   5160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   6960
      TabIndex        =   89
      Top             =   4920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   6960
      TabIndex        =   87
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   6960
      TabIndex        =   85
      Top             =   4440
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   6960
      TabIndex        =   83
      Top             =   4200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   6960
      TabIndex        =   81
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   5760
      TabIndex        =   79
      Top             =   6120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   5760
      TabIndex        =   77
      Top             =   5880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   5760
      TabIndex        =   75
      Top             =   5640
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   5760
      TabIndex        =   73
      Top             =   5400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   5760
      TabIndex        =   71
      Top             =   5160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   5760
      TabIndex        =   69
      Top             =   4920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   5760
      TabIndex        =   67
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   5760
      TabIndex        =   65
      Top             =   4440
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5760
      TabIndex        =   63
      Top             =   4200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   5760
      TabIndex        =   61
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   5760
      TabIndex        =   59
      Top             =   3720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   5760
      TabIndex        =   57
      Top             =   3480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   3360
      TabIndex        =   55
      Top             =   2400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   3360
      TabIndex        =   53
      Top             =   2160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   3360
      TabIndex        =   51
      Top             =   1920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   3360
      TabIndex        =   49
      Top             =   1680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3360
      TabIndex        =   47
      Top             =   1440
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3360
      TabIndex        =   45
      Top             =   1200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   43
      Top             =   2400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   1680
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label CharLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   29
      Top             =   840
      Width           =   555
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a NPC to view."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   72
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drop %:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   3480
      TabIndex        =   25
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drop Items:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   24
      Left            =   2880
      TabIndex        =   24
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   23
      Left            =   3720
      TabIndex        =   23
      Top             =   5760
      Width           =   435
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   3000
      TabIndex        =   22
      Top             =   5760
      Width           =   270
   End
   Begin VB.Label AddDropLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4800
      TabIndex        =   21
      Top             =   5760
      Width           =   345
   End
   Begin VB.Label AddLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2280
      TabIndex        =   19
      Top             =   5760
      Width           =   345
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   1440
      TabIndex        =   18
      Top             =   5760
      Width           =   270
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   240
      TabIndex        =   17
      Top             =   5760
      Width           =   705
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vending Items:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   17
      Left            =   240
      TabIndex        =   15
      Top             =   3240
      Width           =   1545
   End
   Begin VB.Label LoadLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   540
   End
   Begin VB.Label SaveAsLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save As"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7440
      TabIndex        =   13
      Top             =   120
      Width           =   885
   End
   Begin VB.Label SaveLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6600
      TabIndex        =   12
      Top             =   120
      Width           =   555
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AI:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   5400
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respawn:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   4800
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quest:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   6720
      TabIndex        =   7
      Top             =   2280
      Width           =   570
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Give Gold:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   4680
      TabIndex        =   6
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Give EXP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   5
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   5760
      TabIndex        =   4
      Top             =   3240
      Width           =   600
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Doll:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   750
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   5040
      TabIndex        =   1
      Top             =   1080
      Width           =   510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RotateCount As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

Private Sub AddLbl_Click()
Dim I As Integer

    'Check for a valid object index
    If Val(frmMain.IDTxt.Text) > UBound(ObjData) Or Val(IDTxt.Text) < LBound(ObjData) Then
        MsgBox "That is not a valid object!", vbOKOnly
            Exit Sub
    End If
    
    'Make sure item doesn't already exist in the list
    If UBound(ShopObjs) >= 0 Then
        For I = 0 To UBound(ShopObjs)
            If ShopObjs(I).OBJIndex = Val(frmMain.IDTxt.Text) Then
                MsgBox "Item already listed!", vbOKOnly
                Exit Sub
            End If
        Next I
    End If
    
    'Make sure an item isn't in the array at 0 already
    If UBound(ShopObjs) = 0 Then
        If ShopObjs(0).Amount > 0 Then ReDim Preserve ShopObjs(1)
    End If
    
    'Add item to the venditems list
    ShopObjs(UBound(ShopObjs)).Amount = Val(frmMain.AmountTxt.Text)
    ShopObjs(UBound(ShopObjs)).OBJIndex = Val(frmMain.IDTxt.Text)
    
    'Update the visual list
    With frmMain.OBJList
        .Clear
        If UBound(ShopObjs) >= 0 Then
            For I = 0 To UBound(ShopObjs)
                 .AddItem ObjData(ShopObjs(I).OBJIndex).Name & " / " & ShopObjs(I).Amount, I
            Next I
        End If
    End With
    
    'Resize the array to prepare for the next item
    ReDim Preserve ShopObjs(UBound(ShopObjs) + 1)
End Sub

Private Sub AddDropLbl_Click()
Dim I As Integer

    'Check for a valid object index
    If Val(frmMain.IDDropTxt.Text) > UBound(ObjData) Or Val(IDDropTxt.Text) < LBound(ObjData) Then
        MsgBox "That is not a valid object!", vbOKOnly
            Exit Sub
    End If
    
    'Make sure item doesn't already exist in the list
    If UBound(DropObjs) >= 0 Then
        For I = 0 To UBound(DropObjs)
            If DropObjs(I).OBJIndex = Val(frmMain.IDDropTxt.Text) Then
                MsgBox "Item already listed!", vbOKOnly
                Exit Sub
            End If
        Next I
    End If
    
    'Make sure an item isn't in the array at 0 already
    If UBound(DropObjs) = 0 Then
        If DropObjs(0).Amount > 0 Then ReDim Preserve DropObjs(1)
    End If
    
    'Add item to the dropitems list
    DropObjs(UBound(DropObjs)).Amount = Val(frmMain.AmountDropTxt.Text)
    DropObjs(UBound(DropObjs)).OBJIndex = Val(frmMain.IDDropTxt.Text)
    DropObjs(UBound(DropObjs)).DropC = Val(frmMain.DropChanceTxt.Text)
    
    'Update the visual list
    With frmMain.OBJDropList
        .Clear
        If UBound(DropObjs) >= 0 Then
            For I = 0 To UBound(DropObjs)
                 .AddItem ObjData(DropObjs(I).OBJIndex).Name & " / " & DropObjs(I).Amount & " / " & DropObjs(I).DropC & "%", I
            Next I
        End If
    End With
    
    'Resize the array to prepare for the next item
    ReDim Preserve DropObjs(UBound(DropObjs) + 1)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
    
    'Close form
    If Button = vbLeftButton Then
        If X >= Me.ScaleWidth - 23 Then
            If X <= Me.ScaleWidth - 10 Then
                If Y <= 26 Then
                    If Y >= 11 Then
                        Engine_Init_UnloadTileEngine
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub LoadLbl_Click()
Dim NpcNum As String
Dim TempStr() As String
    
    'Confirm
    If MsgBox("Are you sure you wish to load another NPC?" & vbCrLf & "Any changes made to the current NPC will be lost!", vbYesNo) = vbNo Then Exit Sub
    
    'loading a new npc, erase the old npc data
    ReDim ShopObjs(0)
    ReDim DropObjs(0)
    frmMain.OBJList.Clear
    frmMain.OBJDropList.Clear
    If frmMain.SelectNpcCombo.Text = "" Then Exit Sub
    'Load npc
        TempStr = Split(frmMain.SelectNpcCombo.Text, "-")
    
    Editor_OpenNPC Val(TempStr(0))
    
'ErrOut:

End Sub

Private Sub OBJList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As Integer
Dim Index As Integer
Dim TempArray() As Obj
Dim Size As Integer
'Delete an object from the list
If KeyCode = vbKeyDelete Then
    
    Index = OBJList.ListIndex
            
        'remove the object by setting it's amount as 0, reread the array and resize it.
        ShopObjs(Index).Amount = 0
        For I = 0 To UBound(ShopObjs)
            If ShopObjs(I).Amount <> 0 Then
                ReDim Preserve TempArray(Size)
                TempArray(Size).Amount = ShopObjs(I).Amount
                TempArray(Size).OBJIndex = ShopObjs(I).OBJIndex
                Size = Size + 1
            End If
        Next I
        
        'redim the array if it's bigger then 0
        If UBound(ShopObjs) > 0 Then ReDim ShopObjs(UBound(ShopObjs) - 1)
                
        'if the size is zero just redim, otherwise it winds up null... not good.
        If Size > 0 Then
            ShopObjs = TempArray
        Else
            ReDim ShopObjs(0)
        End If
        
    'Update the visual list
    With frmMain.OBJList
        .Clear
        If UBound(ShopObjs) >= 0 Then
            For I = 0 To UBound(ShopObjs)
                'exit the sub if theres no object info
                If ShopObjs(I).OBJIndex = 0 Then Exit For
                 .AddItem ObjData(ShopObjs(I).OBJIndex).Name & " / " & ShopObjs(I).Amount, I
            Next I
        End If
    End With

End If

End Sub

Private Sub OBJdropList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim I As Integer
Dim Index As Integer
Dim TempArray() As Obj
Dim Size As Integer
'Delete an object from the list
If KeyCode = vbKeyDelete Then

        Index = OBJDropList.ListIndex
        
        'remove the object by setting it's amount as 0, reread the array and resize it.
        DropObjs(Index).Amount = 0
        For I = 0 To UBound(DropObjs)
            If DropObjs(I).Amount <> 0 Then
                ReDim Preserve TempArray(Size)
                TempArray(Size).Amount = DropObjs(I).Amount
                TempArray(Size).OBJIndex = DropObjs(I).OBJIndex
                TempArray(Size).DropC = DropObjs(I).DropC
                Size = Size + 1
            End If
        Next I
        
        'redim the array if it's bigger then 0
        If UBound(DropObjs) > 0 Then ReDim DropObjs(UBound(DropObjs) - 1)
        
        'if the size is zero just redim, otherwise it winds up null... not good.
        If Size > 0 Then
            DropObjs = TempArray
        Else
            ReDim DropObjs(0)
        End If
    
    'Update the visual list
    With frmMain.OBJDropList
        .Clear
        If UBound(DropObjs) >= 0 Then
            For I = 0 To UBound(DropObjs)
                'exit the loop if theres nothing to display
                If DropObjs(I).OBJIndex = 0 Then Exit For
                 .AddItem ObjData(DropObjs(I).OBJIndex).Name & " / " & DropObjs(I).Amount & " / " & DropObjs(I).DropC & "%", I
            Next I
        End If
    End With

End If

End Sub

Private Sub RenderTimer_Timer()

    'Usually I don't use timers, but we dont want to create a loop since we dont need intensive drawing
    ElapsedTime = 50
    
    RotateCount = RotateCount + 1
    If RotateCount >= 30 Then
        CharList(1).Heading = CharList(1).Heading + 1
        If CharList(1).Heading > 4 Then CharList(1).Heading = 1
        CharList(1).Heading = CharList(1).Heading
        RotateCount = 0
    End If
    CharList(1).HeadHeading = CharList(1).Heading
    'Render
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
        Engine_Render_Char 1, 0, 32
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Private Sub SaveAsLbl_Click()
Dim RetNumber As Integer

    'Confirm
    If MsgBox("Are you sure you wish to save NPC " & Npcnumber & " as a new number?", vbYesNo) = vbNo Then Exit Sub
    
    'Get number
    RetNumber = Val(InputBox("Please enter the number to save the Object as."))
    If RetNumber = 0 Then Exit Sub
    
    'Check for overwrite
    If NpcExist(RetNumber, False) Then
        'are you sure you want to overwrite?
        If MsgBox("NPC " & RetNumber & " exists! Are you sure you wish to overwrite?", vbOKCancel, "Save as?") = vbOK Then
            Editor_SaveNPC RetNumber
        Else
            MsgBox "Save failed!", vbOKOnly, "Save failed!"
        End If
    Else
        'Save
        Editor_SaveNPC RetNumber
    End If
End Sub

Private Sub DeleteLbl_Click()
Dim RetNumber As Integer

    'Confirm
    If MsgBox("Are you sure you wish to delete NPC " & Npcnumber & " ?", vbYesNo) = vbNo Then Exit Sub
    'Check for overwrite
    If NpcExist(Npcnumber, True) Then
        MsgBox "DELETED!", vbOKOnly, "Gone!"
        Editor_LoadNPCsCombo
    Else
        MsgBox "Delete failed! NPC " & Npcnumber & " does not exist!", vbOKOnly, "Save failed!"
    End If
End Sub

Private Sub SaveLbl_Click()
    'Confirm
    If MsgBox("Are you sure you wish to save changes to NPC " & Npcnumber & "?", vbYesNo) = vbNo Then Exit Sub
    'Save
    Editor_SaveNPC Npcnumber
End Sub

Private Sub CharTxt_Change(Index As Integer)
Editor_SetNPCGrhs Index
End Sub
