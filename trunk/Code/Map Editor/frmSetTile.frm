VERSION 5.00
Begin VB.Form frmSetTile 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Set Tile"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetTile.frx":0000
   ScaleHeight     =   524
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   171
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   108
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   6120
      Width           =   255
   End
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   107
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   106
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   105
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   104
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   103
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   840
      TabIndex        =   41
      ToolTipText     =   "Set light layer 6"
      Top             =   5400
      Width           =   735
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   840
      TabIndex        =   33
      ToolTipText     =   "Set light layer 5"
      Top             =   4440
      Width           =   735
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   25
      ToolTipText     =   "Set light layer 4"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "Set light layer 3"
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   840
      TabIndex        =   9
      ToolTipText     =   "Set light layer 2"
      Top             =   1560
      Width           =   735
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Light"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "Set light layer 1"
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1560
      TabIndex        =   42
      ToolTipText     =   "Set layer 4"
      Top             =   5400
      Width           =   975
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   34
      ToolTipText     =   "Set layer 4"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1560
      TabIndex        =   26
      ToolTipText     =   "Set layer 4"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   1560
      TabIndex        =   18
      ToolTipText     =   "Set layer 4"
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      ToolTipText     =   "Set layer 4"
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Shadow"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "Set layer 4"
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Grh"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   40
      ToolTipText     =   "Set graphic layer 6"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Grh"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   32
      ToolTipText     =   "Set graphic layer 5"
      Top             =   4440
      Width           =   615
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Grh"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   24
      ToolTipText     =   "Set graphic layer 4"
      Top             =   3480
      Width           =   615
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Grh"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Set graphic layer 3"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Grh"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Set graphic layer 2"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Grh"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Set graphic layer 1"
      Top             =   600
      Width           =   615
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   66
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   1305
      Width           =   255
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   65
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   2265
      Width           =   255
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   64
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   3225
      Width           =   255
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   63
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   4185
      Width           =   255
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   62
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   5145
      Width           =   255
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2280
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   61
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   6105
      Width           =   255
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   240
      MaxLength       =   11
      TabIndex        =   43
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   44
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   240
      MaxLength       =   11
      TabIndex        =   45
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   46
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   600
      MaxLength       =   5
      TabIndex        =   47
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   6090
      Width           =   615
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   240
      MaxLength       =   11
      TabIndex        =   35
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   36
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   240
      MaxLength       =   11
      TabIndex        =   37
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   38
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   600
      MaxLength       =   5
      TabIndex        =   39
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   5130
      Width           =   615
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   240
      MaxLength       =   11
      TabIndex        =   27
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   28
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   240
      MaxLength       =   11
      TabIndex        =   29
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   30
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   600
      MaxLength       =   5
      TabIndex        =   31
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   4170
      Width           =   615
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   240
      MaxLength       =   11
      TabIndex        =   19
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   20
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   240
      MaxLength       =   11
      TabIndex        =   21
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   22
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   600
      MaxLength       =   5
      TabIndex        =   23
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   3210
      Width           =   615
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   240
      MaxLength       =   11
      TabIndex        =   11
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   12
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   240
      MaxLength       =   11
      TabIndex        =   13
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   14
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   600
      MaxLength       =   5
      TabIndex        =   15
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   2250
      Width           =   615
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
      ScaleWidth      =   135
      TabIndex        =   50
      ToolTipText     =   "Preview of the tile, with all lights and graphic layers included"
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   600
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   6
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   240
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      MaxLength       =   11
      TabIndex        =   3
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   840
      Width           =   975
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shdw:"
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
      Left            =   1320
      TabIndex        =   102
      Top             =   6120
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shdw:"
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
      Left            =   1320
      TabIndex        =   101
      Top             =   5160
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shdw:"
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
      Left            =   1320
      TabIndex        =   100
      Top             =   4200
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shdw:"
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
      Left            =   1320
      TabIndex        =   99
      Top             =   3240
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shdw:"
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
      Left            =   1320
      TabIndex        =   98
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shdw:"
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
      Left            =   1320
      TabIndex        =   97
      Top             =   1320
      Width           =   540
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
      Index           =   6
      Left            =   1200
      TabIndex        =   96
      Top             =   6120
      Width           =   90
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
      Index           =   5
      Left            =   1200
      TabIndex        =   95
      Top             =   5160
      Width           =   90
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
      Index           =   4
      Left            =   1200
      TabIndex        =   94
      Top             =   4200
      Width           =   90
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
      Index           =   3
      Left            =   1200
      TabIndex        =   93
      Top             =   3240
      Width           =   90
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
      Index           =   2
      Left            =   1200
      TabIndex        =   92
      Top             =   2280
      Width           =   90
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
      Index           =   1
      Left            =   1200
      TabIndex        =   91
      Top             =   1320
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
      Index           =   24
      Left            =   2280
      TabIndex        =   90
      Top             =   5880
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
      Index           =   22
      Left            =   2280
      TabIndex        =   89
      Top             =   5640
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
      Index           =   23
      Left            =   1200
      TabIndex        =   88
      Top             =   5880
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
      Index           =   21
      Left            =   1200
      TabIndex        =   87
      Top             =   5640
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
      Index           =   20
      Left            =   2280
      TabIndex        =   86
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
      Index           =   18
      Left            =   2280
      TabIndex        =   85
      Top             =   4680
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
      Index           =   19
      Left            =   1200
      TabIndex        =   84
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
      Index           =   17
      Left            =   1200
      TabIndex        =   83
      Top             =   4680
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
      Index           =   16
      Left            =   2280
      TabIndex        =   82
      Top             =   3960
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
      Index           =   14
      Left            =   2280
      TabIndex        =   81
      Top             =   3720
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
      Index           =   15
      Left            =   1200
      TabIndex        =   80
      Top             =   3960
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
      Index           =   13
      Left            =   1200
      TabIndex        =   79
      Top             =   3720
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
      Index           =   12
      Left            =   2280
      TabIndex        =   78
      Top             =   3000
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
      Index           =   10
      Left            =   2280
      TabIndex        =   77
      Top             =   2760
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
      Index           =   11
      Left            =   1200
      TabIndex        =   76
      Top             =   3000
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
      Index           =   9
      Left            =   1200
      TabIndex        =   75
      Top             =   2760
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
      Index           =   8
      Left            =   2280
      TabIndex        =   74
      Top             =   2040
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
      Index           =   6
      Left            =   2280
      TabIndex        =   73
      Top             =   1800
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
      Index           =   7
      Left            =   1200
      TabIndex        =   72
      Top             =   2040
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
      Index           =   5
      Left            =   1200
      TabIndex        =   71
      Top             =   1800
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
      Left            =   2280
      TabIndex        =   70
      Top             =   1080
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
      Left            =   2280
      TabIndex        =   69
      Top             =   840
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
      Left            =   1200
      TabIndex        =   68
      Top             =   1080
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
      Left            =   1200
      TabIndex        =   67
      Top             =   840
      Width           =   90
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6:"
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
      Left            =   0
      TabIndex        =   60
      Top             =   5400
      Width           =   180
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
      Index           =   10
      Left            =   120
      TabIndex        =   59
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5:"
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
      Left            =   0
      TabIndex        =   58
      Top             =   4440
      Width           =   180
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
      Index           =   8
      Left            =   120
      TabIndex        =   57
      Top             =   5160
      Width           =   375
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
      Index           =   7
      Left            =   0
      TabIndex        =   56
      Top             =   3480
      Width           =   180
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
      Index           =   6
      Left            =   120
      TabIndex        =   55
      Top             =   4200
      Width           =   375
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
      Left            =   0
      TabIndex        =   54
      Top             =   2520
      Width           =   180
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
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   3240
      Width           =   375
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
      Index           =   3
      Left            =   0
      TabIndex        =   52
      Top             =   1560
      Width           =   180
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
      Index           =   2
      Left            =   120
      TabIndex        =   51
      Top             =   2280
      Width           =   375
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
      Left            =   120
      TabIndex        =   49
      Top             =   1320
      Width           =   375
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
      Index           =   0
      Left            =   0
      TabIndex        =   48
      Top             =   600
      Width           =   180
   End
End
Attribute VB_Name = "frmSetTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DrawPreview()
Dim i As Byte
Dim TempGrh As Grh
Dim TempRect As RECT

 
    TempGrh.FrameCounter = 1
    
    'Set the map set preview
    For i = 1 To 6
        If LayerChk(i).Value = 1 Then
            If Val(GrhTxt(i).Text) < 1 Then
                PreviewMapGrh(i).GrhIndex = 0
            Else
                If PreviewMapGrh(i).GrhIndex <> Val(GrhTxt(i).Text) Then
                    Engine_Init_Grh PreviewMapGrh(i), Val(GrhTxt(i).Text)
                End If
            End If
        Else
            PreviewMapGrh(i).GrhIndex = 0
        End If
    Next i
    
    'Set the view area
    TempRect.bottom = 79
    TempRect.Right = 135

    'Draw the preview
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
    
        'Draw the grhs
        For i = 1 To 6
            If LayerChk(i).Value = 1 Then
                TempGrh.GrhIndex = Val(GrhTxt(i).Text)
                Engine_Render_Grh TempGrh, 0, 0, 0, 0, False, Val(LightTxt((i - 1) * 4 + 1).Text), Val(LightTxt((i - 1) * 4 + 2).Text), Val(LightTxt((i - 1) * 4 + 3).Text), Val(LightTxt((i - 1) * 4 + 4).Text)
            End If
        Next i
        
    D3DDevice.EndScene
    D3DDevice.Present TempRect, TempRect, frmSetTile.PreviewPic.hwnd, ByVal 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = 1
    Engine_Var_Write Data2Path & "MapEditor.ini", "SETTILE", "X", frmSetTile.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "SETTILE", "Y", frmSetTile.Top
    HideFrmSetTile
    
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
    
    ShowFrmTileSelect Index

End Sub

Private Sub GrhTxt_Change(Index As Integer)
Dim i As Integer
On Error GoTo ErrOut

    i = Val(GrhTxt(Index).Text)
    
    'Check for valid range
    If Val(GrhTxt(Index).Text) < 0 Then GrhTxt(Index).Text = "0"
    If Val(GrhTxt(Index).Text) > UBound(GrhData) Then Exit Sub

    DrawPreview
    
    Exit Sub
    
ErrOut:

    GrhTxt(Index).Text = 0

End Sub

Private Sub GrhTxt_KeyPress(Index As Integer, KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub LayerChk_Click(Index As Integer)

    DrawPreview

End Sub

Private Sub LightLbl_Click(Index As Integer)

    'Bring up info box
    ShowFrmARGB LightTxt(Index)

End Sub

Private Sub LightTxt_Change(Index As Integer)
Dim TempRect As RECT
Dim i As Long
Dim j As Byte
On Error GoTo ErrOut
    
    'Check for a valid light value
    i = Val(LightTxt(Index).Text)

    DrawPreview
    
    'Set the view area
    TempRect.bottom = 15
    TempRect.Right = 15
    
    'Draw the light preview
    For i = 1 To 6
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        D3DDevice.BeginScene
            Engine_Render_Rectangle 0, 0, 15, 15, 1, 1, 1, 1, 1, 1, 0, 0, Val(LightTxt((i - 1) * 4 + 1).Text), Val(LightTxt((i - 1) * 4 + 2).Text), Val(LightTxt((i - 1) * 4 + 3).Text), Val(LightTxt((i - 1) * 4 + 4).Text)
        D3DDevice.EndScene
        D3DDevice.Present TempRect, TempRect, frmSetTile.LightPic(i).hwnd, ByVal 0
    Next i

Exit Sub

ErrOut:

    LightTxt(Index).Text = "0"

End Sub

Private Sub LightTxt_KeyPress(Index As Integer, KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If Chr$(KeyAscii) <> "-" Then
                If KeyAscii <> 8 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub ShadowTxt_Change(Index As Integer)
Dim i As Long
On Error GoTo ErrOut

    i = Val(ShadowTxt(Index).Text)
    If i > 0 Then ShadowTxt(Index).Text = 1
    If i < 0 Then ShadowTxt(Index).Text = 0
    
    Exit Sub
    
ErrOut:
    ShadowTxt(Index).Text = 0
End Sub

Private Sub ShadowTxt_KeyPress(Index As Integer, KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub
