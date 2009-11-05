VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Object Editor"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox HPPTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   720
      TabIndex        =   72
      ToolTipText     =   "Percentage of overall HP replenished upon using/equipting (positive or negative)"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox MPPTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   71
      ToolTipText     =   "Percentage of overall MP replenished upon using/equipting (positive or negative)"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox SPPTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   4080
      TabIndex        =   70
      ToolTipText     =   "Percentage of overall SP replenished upon using/equipting (positive or negative)"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox SWingsTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   3840
      TabIndex        =   68
      ToolTipText     =   "The helmet graphic to switch to when equipted/used"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox SHairTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1080
      TabIndex        =   67
      ToolTipText     =   "The hair graphic to switch to when equipted/used"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox SBodyTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   65
      ToolTipText     =   "The body graphic to switch to when equipted/used"
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox SHeadTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   63
      ToolTipText     =   "The head graphic to switch to when equipted/used"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox SWeapTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   1080
      TabIndex        =   61
      ToolTipText     =   "The weapon graphic to switch to when equipted/used"
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox SPTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   3960
      TabIndex        =   58
      ToolTipText     =   "How much SP is replenished upon using/equipting (positive or negative)"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox MPTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2280
      TabIndex        =   57
      ToolTipText     =   "How much MP is replenished upon using/equipting (positive or negative)"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox HPTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   600
      TabIndex        =   56
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   51
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   50
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   49
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   48
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   47
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   46
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   45
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   1200
      TabIndex        =   44
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   15
      Left            =   2160
      TabIndex        =   43
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   16
      Left            =   2160
      TabIndex        =   42
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   17
      Left            =   2160
      TabIndex        =   41
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   18
      Left            =   2160
      TabIndex        =   40
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   22
      Left            =   3120
      TabIndex        =   39
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   23
      Left            =   3120
      TabIndex        =   38
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   24
      Left            =   3120
      TabIndex        =   37
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   25
      Left            =   3120
      TabIndex        =   36
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   19
      Left            =   2160
      TabIndex        =   35
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   34
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   33
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   32
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   26
      Left            =   3120
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   1200
      TabIndex        =   30
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   1200
      TabIndex        =   29
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   14
      Left            =   1200
      TabIndex        =   28
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   20
      Left            =   2160
      TabIndex        =   27
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   27
      Left            =   3120
      TabIndex        =   26
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   21
      Left            =   2160
      TabIndex        =   25
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   28
      Left            =   3120
      TabIndex        =   24
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   29
      Left            =   4080
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   30
      Left            =   4080
      TabIndex        =   22
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   31
      Left            =   4080
      TabIndex        =   21
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   32
      Left            =   4080
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   33
      Left            =   4080
      TabIndex        =   19
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   34
      Left            =   4080
      TabIndex        =   18
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   35
      Left            =   4080
      TabIndex        =   17
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox WeaponTypeTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   3000
      TabIndex        =   15
      ToolTipText     =   "Type of weapon (only if the object type is weapon)"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox ObjTypeTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      ToolTipText     =   "Type of object"
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox PriceTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2280
      TabIndex        =   11
      ToolTipText     =   "The price of the object when purchased from a store"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2280
      TabIndex        =   9
      ToolTipText     =   "Grh number of the graphic the object will use"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer RenderTimer 
      Interval        =   50
      Left            =   4560
      Top             =   1800
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Name of the object"
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   4560
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label WeaponTypeLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   77
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label ObjTypeLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3600
      TabIndex        =   76
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP%:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   240
      TabIndex        =   75
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP%:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   1920
      TabIndex        =   74
      Top             =   4800
      Width           =   480
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP%:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   3600
      TabIndex        =   73
      Top             =   4800
      Width           =   450
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wings:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   3120
      TabIndex        =   69
      Top             =   5280
      Width           =   600
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hair:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   600
      TabIndex        =   66
      Top             =   5520
      Width           =   420
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   1800
      TabIndex        =   64
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Head:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   1800
      TabIndex        =   62
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   60
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paper-Doll:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   59
      Top             =   5040
      Width           =   1200
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   3600
      TabIndex        =   55
      Top             =   4560
      Width           =   315
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   1920
      TabIndex        =   54
      Top             =   4560
      Width           =   345
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   53
      Top             =   4560
      Width           =   330
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replenishers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   120
      TabIndex        =   52
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stat Modifiers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   14
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Object Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1680
      TabIndex        =   12
      Top             =   1560
      Width           =   1110
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grh:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   1080
      Width           =   375
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4575
      TabIndex        =   7
      Top             =   1440
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4200
      TabIndex        =   6
      Top             =   1080
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4560
      TabIndex        =   5
      Top             =   720
      Width           =   555
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   900
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   750
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   555
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

Private Sub GrhTxt_Change()
Dim i As Integer
On Error GoTo ErrOut
    
    i = Val(GrhTxt.Text)
    If Val(GrhTxt.Text) <= 0 Then ObjGrh.GrhIndex = 0 Else Engine_Init_Grh ObjGrh, Val(GrhTxt.Text)

    Exit Sub

ErrOut:

    GrhTxt.Text = "0"
    ObjGrh.GrhIndex = 0

End Sub

Private Sub HPPTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(HPPTxt.Text)
    If i > 100 Then HPPTxt.Text = 100
    
    Exit Sub
    
ErrOut:

    HPPTxt.Text = "0"
    
End Sub

Private Sub HPTxt_Change()
Dim i As Long
On Error GoTo ErrOut

    i = Val(HPTxt.Text)

    Exit Sub
    
ErrOut:

    HPTxt.Text = "0"

End Sub

Private Sub LoadLbl_Click()
Dim FileName As String
Dim TempNum As Integer
    On Error GoTo ErrOut

    'Confirm
    If MsgBox("Are you sure you wish to load another Object?" & vbCrLf & "Any changes made to the current Object will be lost!", vbYesNo) = vbNo Then Exit Sub
    
    'Load map
    With frmMain.CD
        .Filter = "Objects|*.obj"
        .DialogTitle = "Load"
        .FileName = ""
        .InitDir = OBJsPath
        .flags = cdlOFNFileMustExist
        .ShowOpen
    End With
    FileName = Right$(frmMain.CD.FileName, Len(frmMain.CD.FileName) - Len(OBJsPath))
    Editor_LoadOBJ Val(FileName)
    
    Exit Sub
    
ErrOut:

End Sub

Private Sub MPPTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(MPPTxt.Text)
    If i > 100 Then MPPTxt.Text = 100

    Exit Sub
    
ErrOut:

    MPPTxt.Text = "0"
    
End Sub

Private Sub MPTxt_Change()
Dim i As Long
On Error GoTo ErrOut

    i = Val(MPTxt.Text)

    Exit Sub
    
ErrOut:

    MPTxt.Text = "0"
    
End Sub

Private Sub ObjTypeLbl_Click()
Dim s As String

    s = "Object types:" & vbCrLf
    s = s & "1: Use Once" & vbCrLf
    s = s & "2: Weapon" & vbCrLf
    s = s & "3: Armor"
    MsgBox s, vbOKOnly

End Sub

Private Sub ObjTypeTxt_Change()
Dim i As Byte
On Error GoTo ErrOut

    i = Val(ObjTypeTxt.Text)

    Exit Sub
    
ErrOut:

    ObjTypeTxt.Text = "0"
    
End Sub

Private Sub PriceTxt_Change()
Dim i As Long
On Error GoTo ErrOut

    i = Val(PriceTxt.Text)

    Exit Sub
    
ErrOut:

    PriceTxt.Text = "0"
    
End Sub

Private Sub RenderTimer_Timer()

    'Usually I dont use timers, but we dont want to create a loop since we dont need intensive drawing
    ElapsedTime = 50
    
    'Render
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
        Engine_Render_Grh ObjGrh, 1, 1, 0, 1
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Private Sub SaveAsLbl_Click()
Dim RetNumber As Integer

    'Confirm
    If MsgBox("Are you sure you wish to save Object " & OpenIndex & " as a new number?", vbYesNo) = vbNo Then Exit Sub
    
    'Get number
    RetNumber = Val(InputBox("Please enter the number to save the Object as."))
    If RetNumber = 0 Then Exit Sub
    
    'Check for overwrite
    If Engine_FileExist(App.Path & "\OBJs\" & RetNumber & ".obj", vbNormal) Then
        If MsgBox("Object number " & RetNumber & " already exists, are you sure you wish to overwrite it?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Save
    Editor_SaveOBJ RetNumber

End Sub

Private Sub SaveLbl_Click()

    'Confirm
    If MsgBox("Are you sure you wish to save changes to Object " & OpenIndex & "?", vbYesNo) = vbNo Then Exit Sub
    
    'Save
    Editor_SaveOBJ OpenIndex

End Sub

Private Sub SBodyTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(SBodyTxt.Text)

    Exit Sub
    
ErrOut:

    SBodyTxt.Text = "0"
    
End Sub

Private Sub SHairTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(SHairTxt.Text)

    Exit Sub
    
ErrOut:

    SHairTxt.Text = "0"
    
End Sub

Private Sub SHeadTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(SHeadTxt.Text)

    Exit Sub
    
ErrOut:

    SHeadTxt.Text = "0"
    
End Sub

Private Sub SWingsTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(SWingsTxt.Text)

    Exit Sub
    
ErrOut:

    SWingsTxt.Text = "0"
    
End Sub

Private Sub SPPTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(SPPTxt.Text)
    If i > 100 Then SPPTxt.Text = 100

    Exit Sub
    
ErrOut:

    SPPTxt.Text = "0"
    
End Sub

Private Sub SPTxt_Change()
Dim i As Long
On Error GoTo ErrOut

    i = Val(SPTxt.Text)

    Exit Sub
    
ErrOut:

    SPTxt.Text = "0"
    
End Sub

Private Sub StatTxt_Change(Index As Integer)
Dim i As Long
On Error GoTo ErrOut

    i = Val(StatTxt(Index).Text)

    Exit Sub
    
ErrOut:

    StatTxt(Index).Text = "0"

End Sub

Private Sub StatTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Set the ToolTipText
    StatTxt(Index).ToolTipText = "Stat ID: " & Index

End Sub

Private Sub SWeapTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(SWeapTxt.Text)

    Exit Sub

ErrOut:

    SWeapTxt.Text = "0"
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub WeaponTypeLbl_Click()
Dim s As String

    s = "Weapon types:" & vbCrLf
    s = s & "0: Hands" & vbCrLf
    s = s & "1: Staff" & vbCrLf
    s = s & "2: Dagger" & vbCrLf
    s = s & "3: Sword"
    MsgBox s, vbOKOnly
    
End Sub

Private Sub WeaponTypeTxt_Change()
Dim i As Byte
On Error GoTo ErrOut

    i = Val(WeaponTypeTxt.Text)

    Exit Sub
    
ErrOut:

    WeaponTypeTxt.Text = "0"
    
End Sub
