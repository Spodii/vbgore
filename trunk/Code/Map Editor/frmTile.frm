VERSION 5.00
Begin VB.Form frmTile 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tile"
   ClientHeight    =   6225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTile.frx":0000
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   108
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox SfxTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "The number of the .wav file that will be looped on the tile for stuff like waterfalls, birds, etc - set to 0 for nothing"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox WYTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "The Y co-ordinate the user will warp to when stepping on the tile"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox WXTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "The X co-ordinate the user will warp to when stepping on the tile"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox WMapTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "The map the user will warp to when stepping on the tile"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "The graphic index of the layer"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "The amount of the objects on the tile"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox ObjTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "The index of the Object placed on the tile by the *.obj file number"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox NPCTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "The index of the NPC placed on the tile by the *.npc file number"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox MailboxTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   $"frmTile.frx":20D7E
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox BlockedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The blocked value of tile - reccomended you set this value with the Block form unless you know the correct values you want"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   13
      Text            =   "0"
      ToolTipText     =   "Bottom-right light value of the layer"
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Text            =   "0"
      ToolTipText     =   "Bottom-left light value of the layer"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Text            =   "0"
      ToolTipText     =   "Top-right light value of the layer"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Text            =   "0"
      ToolTipText     =   "Top-left light value of the layer"
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sfx:"
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
      Index           =   22
      Left            =   240
      TabIndex        =   55
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label GrhSelectLbl 
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
      Index           =   0
      Left            =   1440
      TabIndex        =   54
      Top             =   3240
      Width           =   90
   End
   Begin VB.Label LightLbl 
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
      Index           =   4
      Left            =   1440
      TabIndex        =   53
      Top             =   5400
      Width           =   90
   End
   Begin VB.Label LightLbl 
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
      Index           =   3
      Left            =   1440
      TabIndex        =   52
      Top             =   4920
      Width           =   90
   End
   Begin VB.Label LightLbl 
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
      Index           =   2
      Left            =   1440
      TabIndex        =   51
      Top             =   4440
      Width           =   90
   End
   Begin VB.Label LightLbl 
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
      Index           =   1
      Left            =   1440
      TabIndex        =   50
      Top             =   3960
      Width           =   90
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
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
      Index           =   21
      Left            =   960
      TabIndex        =   49
      Top             =   2280
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
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
      Left            =   360
      TabIndex        =   48
      Top             =   2280
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warp:"
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
      TabIndex        =   47
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label LayerLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   120
      TabIndex        =   46
      ToolTipText     =   "Click to view layer 1"
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label OldGLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   45
      ToolTipText     =   "What the graphic index was when you last right-clicked the tile"
      Top             =   3480
      Width           =   90
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old:"
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
      TabIndex        =   44
      Top             =   3480
      Width           =   360
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lights:"
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
      Left            =   120
      TabIndex        =   43
      Top             =   3720
      Width           =   585
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
      Index           =   12
      Left            =   120
      TabIndex        =   42
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label LayerLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
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
      Left            =   1320
      TabIndex        =   41
      ToolTipText     =   "Click to view layer 6"
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label LayerLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   1080
      TabIndex        =   40
      ToolTipText     =   "Click to view layer 5"
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label LayerLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
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
      Left            =   840
      TabIndex        =   39
      ToolTipText     =   "Click to view layer 4"
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label LayerLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
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
      Left            =   600
      TabIndex        =   38
      ToolTipText     =   "Click to view layer 3"
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label LayerLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Left            =   360
      TabIndex        =   37
      ToolTipText     =   "Click to view layer 2"
      Top             =   3000
      Width           =   120
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old:"
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
      Left            =   120
      TabIndex        =   36
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old:"
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
      Left            =   120
      TabIndex        =   34
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old:"
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
      Left            =   120
      TabIndex        =   33
      Top             =   5640
      Width           =   360
   End
   Begin VB.Label OldLLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   32
      ToolTipText     =   "What the light was when you last right-clicked the tile"
      Top             =   4200
      Width           =   90
   End
   Begin VB.Label OldLLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   31
      ToolTipText     =   "What the light was when you last right-clicked the tile"
      Top             =   4680
      Width           =   90
   End
   Begin VB.Label OldLLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   30
      ToolTipText     =   "What the light was when you last right-clicked the tile"
      Top             =   5160
      Width           =   90
   End
   Begin VB.Label OldLLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   29
      ToolTipText     =   "What the light was when you last right-clicked the tile"
      Top             =   5640
      Width           =   90
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#:"
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
      Left            =   600
      TabIndex        =   28
      Top             =   1800
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBJ:"
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
      Left            =   240
      TabIndex        =   27
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC:"
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
      TabIndex        =   26
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mailbox:"
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
      Index           =   8
      Left            =   240
      TabIndex        =   25
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blocked:"
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
      Index           =   7
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   765
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4:"
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
      Left            =   240
      TabIndex        =   23
      Top             =   5400
      Width           =   180
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3:"
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
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   180
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2:"
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
      Left            =   240
      TabIndex        =   21
      Top             =   4440
      Width           =   180
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1:"
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
      Left            =   240
      TabIndex        =   20
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label YLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   19
      ToolTipText     =   "Tile Y position of the tile you are viewing"
      Top             =   600
      Width           =   270
   End
   Begin VB.Label XLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   18
      ToolTipText     =   "Tile X position of the tile you are viewing"
      Top             =   600
      Width           =   270
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Layers:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   630
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
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
      Left            =   840
      TabIndex        =   16
      Top             =   600
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   195
   End
End
Attribute VB_Name = "frmTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetInfo(ByVal tX As Byte, ByVal tY As Byte)

    'Check for valid selected layer value
    If SelectedLayer < 1 Then SelectedLayer = 1
    If SelectedLayer > 6 Then SelectedLayer = 6

    'Set the information
    If MapData(tX, tY).NPCIndex > 0 Then frmTile.NPCTxt.Text = CharList(MapData(tX, tY).NPCIndex).NPCNumber Else frmTile.NPCTxt.Text = 0
    ObjTxt.Text = MapData(tX, tY).ObjInfo.ObjIndex
    AmountTxt.Text = MapData(tX, tY).ObjInfo.Amount
    BlockedTxt.Text = MapData(tX, tY).Blocked
    If frmBlock.Visible = True Then
        If MapData(tX, tY).Blocked And 1 Then frmBlock.BlockChk(1).Value = 1 Else frmBlock.BlockChk(1).Value = 0
        If MapData(tX, tY).Blocked And 2 Then frmBlock.BlockChk(2).Value = 1 Else frmBlock.BlockChk(2).Value = 0
        If MapData(tX, tY).Blocked And 4 Then frmBlock.BlockChk(3).Value = 1 Else frmBlock.BlockChk(3).Value = 0
        If MapData(tX, tY).Blocked And 8 Then frmBlock.BlockChk(4).Value = 1 Else frmBlock.BlockChk(4).Value = 0
        If MapData(tX, tY).BlockedAttack Then frmBlock.BlockAttackChk.Value = 1 Else frmBlock.BlockAttackChk.Value = 0
    End If
    WMapTxt.Text = MapData(tX, tY).TileExit.Map
    WXTxt.Text = MapData(tX, tY).TileExit.X
    WYTxt.Text = MapData(tX, tY).TileExit.Y
    XLbl.Caption = tX
    YLbl.Caption = tY
    MailboxTxt.Text = MapData(tX, tY).Mailbox
    LayerLbl_Click SelectedLayer
    
End Sub

Private Sub AmountTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub BlockedTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Data2Path & "MapEditor.ini", "TILE", "X", frmTile.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "TILE", "Y", frmTile.Top
    HideFrmTile
    
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
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub
Private Sub GrhSelectLbl_Click(Index As Integer)

    ShowFrmTileSelect 0

End Sub

Private Sub GrhTxt_KeyPress(KeyAscii As Integer)

    'Avoid invalid numbers
    If GrhTxt.Text < 0 Then Exit Sub
    If GrhTxt.Text > UBound(GrhData) Then Exit Sub
    If XLbl.Caption < XMinMapSize Then Exit Sub
    If XLbl.Caption > XMaxMapSize Then Exit Sub
    If YLbl.Caption < YMinMapSize Then Exit Sub
    If YLbl.Caption > YMaxMapSize Then Exit Sub
    
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    'Change the graphic
    Engine_Init_Grh MapData(XLbl.Caption, YLbl.Caption).Graphic(SelectedLayer), CInt(GrhTxt.Text)
    
End Sub

Private Sub LayerLbl_Click(Index As Integer)
Dim i As Byte

    'Avoid invalid numbers
    If XLbl.Caption < XMinMapSize Then Exit Sub
    If XLbl.Caption > XMaxMapSize Then Exit Sub
    If YLbl.Caption < YMinMapSize Then Exit Sub
    If YLbl.Caption > YMaxMapSize Then Exit Sub
    
    'Set the layer
    SelectedLayer = Index
    
    'Set colors
    For i = 1 To 6
        If i <> Index Then
            LayerLbl(i).ForeColor = &H8000000F
        Else
            LayerLbl(i).ForeColor = 255
        End If
    Next i
    
    'Set values
    GrhTxt.Text = MapData(XLbl.Caption, YLbl.Caption).Graphic(Index).GrhIndex
    OldGLbl.Caption = GrhTxt.Text
    For i = 1 To 4
        LightTxt(i).Text = MapData(XLbl.Caption, YLbl.Caption).Light(Index + ((i - 1) * 4))
        OldLLbl(i).Caption = LightTxt(i).Text
    Next i
    ShadowChk.Value = MapData(XLbl.Caption, YLbl.Caption).Shadow(Index)
    
End Sub

Private Sub LightLbl_Click(Index As Integer)

    ShowFrmARGB LightTxt(Index)

End Sub

Private Sub LightTxt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long
On Error GoTo ErrOut
    
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    i = Val(LightTxt(Index).Text)

    'Set the light
    MapData(XLbl.Caption, YLbl.Caption).Light(((SelectedLayer - 1) * 4) + Index) = CLng(LightTxt(Index).Text)
    
    Exit Sub
    
ErrOut:

    LightTxt(Index).Text = 0
    
End Sub

Private Sub MailboxTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = CByte(MailboxTxt.Text)
    If i > 1 Then i = 1
    If i < 0 Then i = 0
    
    'Set as mailbox
    MapData(XLbl.Caption, YLbl.Caption).Mailbox = i

    Exit Sub

ErrOut:

    MailboxTxt.Text = "0"

End Sub

Private Sub MailboxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub NPCTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub ObjTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub SfxTxt_Change()
Dim i As Integer
On Error GoTo ErrOut
    
    i = Val(SfxTxt.Text)
    MapData(XLbl.Caption, YLbl.Caption).Sfx = i
    
    Exit Sub
    
ErrOut:

    SfxTxt.Text = 0

End Sub

Private Sub SfxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub ShadowChk_Click()

    MapData(XLbl.Caption, YLbl.Caption).Shadow(SelectedLayer) = ShadowChk.Value

End Sub

Private Sub WMapTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub WXTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub WYTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

