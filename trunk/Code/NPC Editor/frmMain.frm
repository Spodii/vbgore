VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "NPC Editor"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   526
   ScaleMode       =   0  'User
   ScaleWidth      =   352
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox IDTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   74
      Text            =   "-1"
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox AmountTxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   72
      Text            =   "-1"
      Top             =   7440
      Width           =   735
   End
   Begin VB.ListBox OBJList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      ItemData        =   "frmMain.frx":9F72C
      Left            =   2520
      List            =   "frmMain.frx":9F72E
      TabIndex        =   71
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Timer RenderTimer 
      Interval        =   50
      Left            =   4320
      Top             =   2040
   End
   Begin VB.TextBox WeaponTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   66
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox BodyTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   65
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox HeadTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   64
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox HairTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   63
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox AITxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   62
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox DescTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   61
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2280
      TabIndex        =   60
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox HeadingTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   59
      Top             =   6360
      Width           =   495
   End
   Begin VB.TextBox RespawnTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   58
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox QuestTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   57
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox GiveGoldTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   56
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox GiveExpTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   55
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CheckBox HostileChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Hostile"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   53
      Top             =   2880
      Width           =   855
   End
   Begin VB.CheckBox AttackChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Attackable"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   52
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   35
      Left            =   4080
      TabIndex        =   46
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   34
      Left            =   4080
      TabIndex        =   45
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   33
      Left            =   4080
      TabIndex        =   44
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   32
      Left            =   4080
      TabIndex        =   43
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   31
      Left            =   4080
      TabIndex        =   42
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   30
      Left            =   4080
      TabIndex        =   41
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   29
      Left            =   4080
      TabIndex        =   40
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   28
      Left            =   3120
      TabIndex        =   39
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   21
      Left            =   2160
      TabIndex        =   38
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   27
      Left            =   3120
      TabIndex        =   37
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   20
      Left            =   2160
      TabIndex        =   36
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   1200
      TabIndex        =   35
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   1200
      TabIndex        =   34
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   1200
      TabIndex        =   33
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   26
      Left            =   3120
      TabIndex        =   32
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   31
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   30
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   240
      TabIndex        =   29
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   19
      Left            =   2160
      TabIndex        =   28
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   25
      Left            =   3120
      TabIndex        =   27
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   24
      Left            =   3120
      TabIndex        =   26
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   23
      Left            =   3120
      TabIndex        =   25
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   22
      Left            =   3120
      TabIndex        =   24
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   18
      Left            =   2160
      TabIndex        =   23
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   17
      Left            =   2160
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   16
      Left            =   2160
      TabIndex        =   21
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   2160
      TabIndex        =   20
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   1200
      TabIndex        =   19
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   17
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   16
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   240
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3720
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4680
      TabIndex        =   76
      Top             =   7440
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   3330
      TabIndex        =   75
      Top             =   7440
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   1800
      TabIndex        =   73
      Top             =   7440
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   17
      Left            =   2520
      TabIndex        =   70
      Top             =   5040
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4575
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   67
      Top             =   720
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   1680
      TabIndex        =   54
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Heading:"
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
      Left            =   240
      TabIndex        =   51
      Top             =   6360
      Width           =   780
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   50
      Top             =   6120
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   240
      TabIndex        =   49
      Top             =   5880
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   240
      TabIndex        =   48
      Top             =   5640
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   47
      Top             =   5400
      Width           =   885
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Misc:"
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
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   555
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   600
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
      TabIndex        =   9
      Top             =   600
      Width           =   900
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
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   1560
      TabIndex        =   8
      Top             =   1560
      Width           =   1185
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
      Index           =   6
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Width           =   495
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
      Index           =   5
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   780
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
      Index           =   4
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   420
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
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   525
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
      TabIndex        =   3
      Top             =   600
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   510
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

Private Sub AddLbl_Click()
Dim i As Integer

    'Check for a valid object index
    If Val(AmountTxt.Text) > UBound(ObjData) Or Val(IDTxt.Text) < LBound(ObjData) Then
        MsgBox "Invalid item index (" & Val(IDTxt.Text) & ")!" & vbCrLf _
            & "LBound: " & LBound(ObjData) & " UBound: " & UBound(ObjData), vbOKOnly
            Exit Sub
    End If

    'Make sure item doesn't already exist in the list
    If OpenNPC.NumVendItems > 0 Then
        For i = 1 To OpenNPC.NumVendItems
            If OpenNPC.VendItems(i).OBJIndex = Val(IDTxt.Text) Then
                MsgBox "Item already exists in the list", vbOKOnly
                Exit Sub
            End If
        Next i
    End If
    
    'Add item to the venditems list
    OpenNPC.NumVendItems = OpenNPC.NumVendItems + 1
    If OpenNPC.NumVendItems = 1 Then
        ReDim OpenNPC.VendItems(1 To OpenNPC.NumVendItems)
    Else
        ReDim Preserve OpenNPC.VendItems(1 To OpenNPC.NumVendItems)
    End If
    OpenNPC.VendItems(OpenNPC.NumVendItems).Amount = Val(AmountTxt.Text)
    OpenNPC.VendItems(OpenNPC.NumVendItems).OBJIndex = Val(IDTxt.Text)
    
    'Update the visual list
    Editor_UpdateVendItems

End Sub

Private Sub BodyTxt_Change()

    Editor_SetNPCGrhs

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

Private Sub HairTxt_Change()

    Editor_SetNPCGrhs

End Sub

Private Sub HeadTxt_Change()

    Editor_SetNPCGrhs

End Sub

Private Sub LoadLbl_Click()
Dim FileName As String
Dim TempNum As Integer
    On Error GoTo ErrOut

    'Confirm
    If MsgBox("Are you sure you wish to load another NPC?" & vbCrLf & "Any changes made to the current NPC will be lost!", vbYesNo) = vbNo Then Exit Sub
    
    'Load map
    With frmMain.CD
        .Filter = "NPCs|*.npc"
        .DialogTitle = "Load"
        .FileName = ""
        .InitDir = NPCPath
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End With
    FileName = Right$(frmMain.CD.FileName, Len(frmMain.CD.FileName) - Len(NPCPath))
    Editor_OpenNPC Val(FileName)
    
ErrOut:

End Sub

Private Sub OBJList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

    'Delete an object from the list
    If KeyCode = vbKeyDelete Then
        If OpenNPC.NumVendItems = 0 Then Exit Sub
        If OBJList.ListIndex + 1 = 0 Then Exit Sub
        
        'Remove the object
        OpenNPC.VendItems(OBJList.ListIndex + 1).OBJIndex = 0
        OpenNPC.VendItems(OBJList.ListIndex + 1).Amount = 0
        
        'Update the venditems list
        If OpenNPC.NumVendItems <> OBJList.ListIndex + 1 Then
            For i = (OBJList.ListIndex + 1) To (OpenNPC.NumVendItems - 1)
                OpenNPC.VendItems(i) = OpenNPC.VendItems(i + 1)
            Next i
        End If
        OpenNPC.NumVendItems = OpenNPC.NumVendItems - 1
        If OpenNPC.NumVendItems = 0 Then
            ReDim OpenNPC.VendItems(0)
        Else
            ReDim Preserve OpenNPC.VendItems(1 To OpenNPC.NumVendItems)
        End If
        
        'Update the visual list
        Editor_UpdateVendItems
        
    End If

End Sub

Private Sub RenderTimer_Timer()

    'Usually I dont use timers, but we dont want to create a loop since we dont need intensive drawing
    ElapsedTime = 50
    
    RotateCount = RotateCount + 1
    If RotateCount >= 30 Then
        CharList(1).Heading = CharList(1).Heading + 1
        If CharList(1).Heading > 8 Then CharList(1).Heading = 1
        CharList(1).HeadHeading = CharList(1).Heading
        RotateCount = 0
    End If
    
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
    If MsgBox("Are you sure you wish to save changes to NPC " & NPCNumber & " as a new number?", vbYesNo) = vbNo Then Exit Sub
    
    'Get number
    RetNumber = Val(InputBox("Please enter the number to save the NPC as."))
    If RetNumber = 0 Then Exit Sub
    
    'Check for overwrite
    If Engine_FileExist(NPCPath & RetNumber & ".npc", vbNormal) Then
        If MsgBox("NPC number " & RetNumber & " already exists, are you sure you wish to overwrite it?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    'Save
    Editor_SaveNPC RetNumber

End Sub

Private Sub SaveLbl_Click()

    'Confirm
    If MsgBox("Are you sure you wish to save changes to NPC " & NPCNumber & "?", vbYesNo) = vbNo Then Exit Sub
    
    'Save
    Editor_SaveNPC NPCNumber

End Sub

Private Sub WeaponTxt_Change()

    Editor_SetNPCGrhs

End Sub
