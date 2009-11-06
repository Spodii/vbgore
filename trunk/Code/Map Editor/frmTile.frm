VERSION 5.00
Begin VB.Form frmTile 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tile"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox OldLLbl 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   48
      Text            =   "0"
      ToolTipText     =   "Top-left light value of the layer"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox OldLLbl 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   47
      Text            =   "0"
      ToolTipText     =   "Top-left light value of the layer"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox OldLLbl 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   46
      Text            =   "0"
      ToolTipText     =   "Top-left light value of the layer"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox OldLLbl 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   45
      Text            =   "0"
      ToolTipText     =   "Top-left light value of the layer"
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox OldGLbl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   44
      Text            =   "0"
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox SignTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   42
      Text            =   "0"
      ToolTipText     =   "The number of the sign from Signs.dat"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox SfxTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "The number of the .wav file that will be looped on the tile for stuff like waterfalls, birds, etc - set to 0 for nothing"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shadow"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   7680
      Width           =   975
   End
   Begin VB.TextBox WYTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "The Y co-ordinate the user will warp to when stepping on the tile"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox WXTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "The X co-ordinate the user will warp to when stepping on the tile"
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox WMapTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "The map the user will warp to when stepping on the tile"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "The graphic index of the layer"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox NPCTxt 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "The index of the NPC placed on the tile by the *.npc file number"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox MailboxTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   $"frmTile.frx":0000
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox BlockedTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The blocked value of tile - reccomended you set this value with the Block form unless you know the correct values you want"
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Text            =   "0"
      ToolTipText     =   "Bottom-right light value of the layer"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Text            =   "0"
      ToolTipText     =   "Bottom-left light value of the layer"
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   480
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "Top-right light value of the layer"
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "Top-left light value of the layer"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sign:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   240
      TabIndex        =   43
      Top             =   1080
      Width           =   450
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   22
      Left            =   240
      TabIndex        =   41
      Top             =   2520
      Width           =   345
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   21
      Left            =   960
      TabIndex        =   40
      Top             =   2160
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   39
      Top             =   2160
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   120
      TabIndex        =   38
      Top             =   1800
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "Click to view layer 1"
      Top             =   3120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   240
      TabIndex        =   36
      Top             =   3840
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   0
      TabIndex        =   35
      Top             =   4560
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   180
      TabIndex        =   34
      Top             =   3480
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   1320
      TabIndex        =   33
      ToolTipText     =   "Click to view layer 6"
      Top             =   3120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   1080
      TabIndex        =   32
      ToolTipText     =   "Click to view layer 5"
      Top             =   3120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   31
      ToolTipText     =   "Click to view layer 4"
      Top             =   3120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   30
      ToolTipText     =   "Click to view layer 3"
      Top             =   3120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   29
      ToolTipText     =   "Click to view layer 2"
      Top             =   3120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   28
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   27
      Top             =   5880
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   26
      Top             =   6600
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   120
      TabIndex        =   25
      Top             =   7320
      Width           =   360
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   24
      Top             =   1440
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   720
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   22
      Top             =   360
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   6960
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   20
      Top             =   6240
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   5520
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   18
      Top             =   4800
      Width           =   180
   End
   Begin VB.Label YLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      ToolTipText     =   "Tile Y position of the tile you are viewing"
      Top             =   120
      Width           =   270
   End
   Begin VB.Label XLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   16
      ToolTipText     =   "Tile X position of the tile you are viewing"
      Top             =   120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   2880
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   14
      Top             =   120
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "frmTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetTileInfo(ByVal tX As Byte, ByVal tY As Byte)

    'Check for valid selected layer value
    If SelectedLayer < 1 Then SelectedLayer = 1
    If SelectedLayer > 6 Then SelectedLayer = 6

    'Set the information
    If MapData(tX, tY).NPCIndex > 0 Then frmTile.NPCTxt.Text = CharList(MapData(tX, tY).NPCIndex).NPCNumber Else frmTile.NPCTxt.Text = 0
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
    SignTxt.Text = MapData(tX, tY).Sign
    SfxTxt.Text = MapData(tX, tY).Sfx
    LayerLbl_Click SelectedLayer
    
End Sub

Private Sub BlockedTxt_Change()
Dim i As Byte
On Error GoTo ErrOut

    i = CByte(BlockedTxt.Text)
    If i < 0 Then i = 0
    
    'Set the sign value
    MapData(XLbl.Caption, YLbl.Caption).Blocked = i

    Exit Sub

ErrOut:

End Sub

Private Sub BlockedTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub BlockedTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    SetInfo "The blocked value of tile - reccomended you set this value with the Block form unless you know the correct values you want."

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub GrhTxt_Change()
Dim i As Long
On Error GoTo ErrOut

    i = CLng(GrhTxt.Text)
    If i < 0 Then i = 0
    
    'Set the grh value
    Engine_Init_Grh MapData(XLbl.Caption, YLbl.Caption).Graphic(SelectedLayer), i
    Engine_CreateTileLayers
    
    Exit Sub

ErrOut:

    Engine_Init_Grh MapData(XLbl.Caption, YLbl.Caption).Graphic(SelectedLayer), 0
    
End Sub

Private Sub GrhTxt_KeyPress(KeyAscii As Integer)

    'Avoid invalid numbers
    If GrhTxt.Text = vbNullString Then Exit Sub
    If Val(GrhTxt.Text) < 0 Then Exit Sub
    If Val(GrhTxt.Text) > UBound(GrhData) Then Exit Sub
    If Val(XLbl.Caption) < 1 Then Exit Sub
    If Val(XLbl.Caption) > MapInfo.Width Then Exit Sub
    If Val(YLbl.Caption) < 1 Then Exit Sub
    If Val(YLbl.Caption) > MapInfo.Height Then Exit Sub
    
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> vbKeyDelete Then
                If KeyAscii <> vbKeyBack Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If

End Sub

Private Sub GrhTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    SetInfo "The grh index of the layer."

End Sub

Private Sub LayerLbl_Click(Index As Integer)
Dim i As Byte

    'Avoid invalid numbers
    If Val(XLbl.Caption) < 1 Then Exit Sub
    If Val(XLbl.Caption) > MapInfo.Width Then Exit Sub
    If Val(YLbl.Caption) < 1 Then Exit Sub
    If Val(YLbl.Caption) > MapInfo.Height Then Exit Sub
    
    'Set the layer
    SelectedLayer = Index
    
    'Set colors
    For i = 1 To 6
        If i <> Index Then
            LayerLbl(i).ForeColor = &H80000008
        Else
            LayerLbl(i).ForeColor = 255
        End If
    Next i
    
    'Set values
    GrhTxt.Text = MapData(XLbl.Caption, YLbl.Caption).Graphic(Index).GrhIndex
    OldGLbl.Text = GrhTxt.Text
    For i = 1 To 4
        LightTxt(i).Text = MapData(XLbl.Caption, YLbl.Caption).Light(i + ((Index - 1) * 4))
        OldLLbl(i).Text = LightTxt(i).Text
    Next i
    ShadowChk.Value = MapData(XLbl.Caption, YLbl.Caption).Shadow(Index)
    
End Sub

Private Sub LayerLbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Click to view information for layer " & Index & "."

End Sub

Private Sub LightTxt_Change(Index As Integer)
Dim i As Long
On Error GoTo ErrOut

    i = CLng(LightTxt(Index).Text)

    'Set the sign value
    MapData(XLbl.Caption, YLbl.Caption).Light(((SelectedLayer - 1) * 4) + Index) = i
    
    Exit Sub

ErrOut:

    MapData(XLbl.Caption, YLbl.Caption).Light(((SelectedLayer - 1) * 4) + Index) = 0
    
End Sub

Private Sub LightTxt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Long
On Error GoTo ErrOut
    
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If Chr$(KeyAscii) <> "-" Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        If KeyAscii <> 8 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    i = Val(LightTxt(Index).Text)

    'Set the light
    MapData(XLbl.Caption, YLbl.Caption).Light(((SelectedLayer - 1) * 4) + Index) = CLng(LightTxt(Index).Text)
    
    Exit Sub
    
ErrOut:
    
End Sub

Private Sub LightTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s As String

    Select Case Index
        Case 1: s = "top-left"
        Case 2: s = "top-right"
        Case 3: s = "bottom-left"
        Case 4: s = "bottom-right"
    End Select
    
    SetInfo "Light value in the tile's " & s & " corner."

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

End Sub

Private Sub MailboxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub MailboxTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "If the tile has a mailbox, 1 for yes, 0 for no - this will not change the looks of the tile, only allow users to check their mail while clicking the tile."

End Sub

Private Sub NPCTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub NPCTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The index of the NPC placed on the tile by the *.npc file number."

End Sub

Private Sub OldGLbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The graphic index that used to be on the layer."

End Sub

Private Sub SfxTxt_Change()
Dim i As Integer
On Error GoTo ErrOut
    
    i = Val(SfxTxt.Text)
    MapData(XLbl.Caption, YLbl.Caption).Sfx = i
    
    Exit Sub
    
ErrOut:

End Sub

Private Sub SfxTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SfxTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The number of the .wav file that will be looped on the tile for stuff like waterfalls, birds, etc - set to 0 for nothing."

End Sub

Private Sub ShadowChk_Click()

    MapData(XLbl.Caption, YLbl.Caption).Shadow(SelectedLayer) = ShadowChk.Value

End Sub

Private Sub ShadowChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Whether or not a shadow is placed on this layer."

End Sub

Private Sub SignTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = CInt(SignTxt.Text)
    If i < 0 Then i = 0
    
    'Set the sign value
    MapData(XLbl.Caption, YLbl.Caption).Sign = i

    Exit Sub

ErrOut:

End Sub

Private Sub SignTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SignTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The number of the sign from the Signs.dat file."

End Sub

Private Sub WMapTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = CInt(WMapTxt.Text)
    If i < 0 Then i = 0
    
    'Set the sign value
    MapData(XLbl.Caption, YLbl.Caption).TileExit.Map = i

    Exit Sub

ErrOut:

End Sub

Private Sub WMapTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub WMapTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The map the user will warp to when stepping on the tile."

End Sub

Private Sub WXTxt_Change()
Dim i As Byte
On Error GoTo ErrOut

    i = CByte(WXTxt.Text)
    If i < 0 Then i = 0
    If i > MapInfo.Width Then i = MapInfo.Width
    
    'Set the sign value
    MapData(XLbl.Caption, YLbl.Caption).TileExit.X = i

    Exit Sub

ErrOut:
    
End Sub

Private Sub WXTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub WXTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The X co-ordinate the user will warp to when stepping on the tile."

End Sub

Private Sub WYTxt_Change()
Dim i As Byte
On Error GoTo ErrOut

    i = CByte(WYTxt.Text)
    If i < 0 Then i = 0
    If i > MapInfo.Height Then i = MapInfo.Height
    
    'Set the sign value
    MapData(XLbl.Caption, YLbl.Caption).TileExit.Y = i

    Exit Sub

ErrOut:

End Sub

Private Sub WYTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                If KeyAscii <> vbKeyDelete Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub WYTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The Y co-ordinate the user will warp to when stepping on the tile."

End Sub

