VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "Game Configuration"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "Window Controls"
      Height          =   1575
      Left            =   240
      TabIndex        =   78
      Top             =   5400
      Width           =   5895
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   14
         Left            =   1200
         TabIndex        =   83
         Text            =   "Text1"
         ToolTipText     =   "Hide / show character stats window"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   13
         Left            =   4200
         TabIndex        =   82
         Text            =   "Text1"
         ToolTipText     =   "Hide / show the quick bar"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   12
         Left            =   1200
         TabIndex        =   81
         Text            =   "Text1"
         ToolTipText     =   "Hide / show game menu"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   11
         Left            =   4200
         TabIndex        =   80
         Text            =   "Text1"
         ToolTipText     =   "Hide / show the user inventory"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   10
         Left            =   1200
         TabIndex        =   79
         Text            =   "Text1"
         ToolTipText     =   "Hide / show chat window"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Stats:"
         Height          =   195
         Left            =   120
         TabIndex        =   88
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Quick bar:"
         Height          =   195
         Left            =   2880
         TabIndex        =   87
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Menu:"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Inventory:"
         Height          =   195
         Left            =   2880
         TabIndex        =   85
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Chat:"
         Height          =   195
         Left            =   120
         TabIndex        =   84
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Quick Bar Hot-Keys"
      Height          =   1935
      Left            =   240
      TabIndex        =   53
      Top             =   3360
      Width           =   5895
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   27
         Left            =   4320
         TabIndex        =   65
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 12"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   26
         Left            =   2400
         TabIndex        =   64
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 11"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   25
         Left            =   480
         TabIndex        =   63
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 10"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   24
         Left            =   4320
         TabIndex        =   62
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 9"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   23
         Left            =   2400
         TabIndex        =   61
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 8"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   22
         Left            =   480
         TabIndex        =   60
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 7"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   21
         Left            =   4320
         TabIndex        =   59
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 6"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   20
         Left            =   2400
         TabIndex        =   58
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 5"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   19
         Left            =   480
         TabIndex        =   57
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 4"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   18
         Left            =   4320
         TabIndex        =   56
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 3"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   17
         Left            =   2400
         TabIndex        =   55
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 2"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   16
         Left            =   480
         TabIndex        =   54
         Text            =   "Text1"
         ToolTipText     =   "Hot-key slot 1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "12:"
         Height          =   195
         Left            =   3960
         TabIndex        =   77
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "11:"
         Height          =   195
         Left            =   2040
         TabIndex        =   76
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "10:"
         Height          =   195
         Left            =   120
         TabIndex        =   75
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "9:"
         Height          =   195
         Left            =   3960
         TabIndex        =   74
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "8:"
         Height          =   195
         Left            =   2040
         TabIndex        =   73
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "7:"
         Height          =   195
         Left            =   120
         TabIndex        =   72
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "6:"
         Height          =   195
         Left            =   3960
         TabIndex        =   71
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "5:"
         Height          =   195
         Left            =   2040
         TabIndex        =   70
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "4:"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "3:"
         Height          =   195
         Left            =   3960
         TabIndex        =   68
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "2:"
         Height          =   195
         Left            =   2040
         TabIndex        =   67
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "1:"
         Height          =   195
         Left            =   120
         TabIndex        =   66
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "General Controls"
      Height          =   2895
      Left            =   240
      TabIndex        =   24
      Top             =   360
      Width           =   5895
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   30
         Left            =   4200
         TabIndex        =   38
         Text            =   "Text1"
         ToolTipText     =   "Instantly reply to the last person to whisper to you"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   29
         Left            =   1200
         TabIndex        =   37
         Text            =   "Text1"
         ToolTipText     =   "Target the closest NPC the user is facing"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   28
         Left            =   4200
         TabIndex        =   36
         Text            =   "Text1"
         ToolTipText     =   "Reset the GUI window positions to the default locations"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   15
         Left            =   4200
         TabIndex        =   35
         Text            =   "Text1"
         ToolTipText     =   "Toggle mini-map display"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   34
         Text            =   "Text1"
         ToolTipText     =   "Zoom out to the normal view (Motion Blur must be enabled)"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   8
         Left            =   4200
         TabIndex        =   33
         Text            =   "Text1"
         ToolTipText     =   "Zoom in on the center of the screen (Motion Blur must be enabled)"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   7
         Left            =   4200
         TabIndex        =   32
         Text            =   "Text1"
         ToolTipText     =   "Scroll the chat buffer down (newer messages)"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   6
         Left            =   4200
         TabIndex        =   31
         Text            =   "Text1"
         ToolTipText     =   "Scroll the chat buffer up (older messages)"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   30
         Text            =   "Text1"
         ToolTipText     =   "Move character right (East)"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   29
         Text            =   "Text1"
         ToolTipText     =   "Move character left (West)"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   28
         Text            =   "Text1"
         ToolTipText     =   "Move character down (South)"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   27
         Text            =   "Text1"
         ToolTipText     =   "Move character up (North)"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   26
         Text            =   "Text1"
         ToolTipText     =   "Picking up items off the ground"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox KeyTxt 
         BackColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   25
         Text            =   "Text1"
         ToolTipText     =   "Basic weapon attack"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Quick reply:"
         Height          =   195
         Left            =   2880
         TabIndex        =   52
         Top             =   2520
         Width           =   1230
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Quick target:"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Reset GUI:"
         Height          =   195
         Left            =   2880
         TabIndex        =   50
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Mini-map:"
         Height          =   195
         Left            =   2880
         TabIndex        =   49
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Zoom out:"
         Height          =   195
         Left            =   2880
         TabIndex        =   48
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Zoom in:"
         Height          =   195
         Left            =   2880
         TabIndex        =   47
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Scroll chat down:"
         Height          =   195
         Left            =   2880
         TabIndex        =   46
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Scroll chat up:"
         Height          =   195
         Left            =   2880
         TabIndex        =   45
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move right:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move left:"
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move down:"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Move up:"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Pick up:"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Attack:"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H80000005&
      Caption         =   "Audio"
      Height          =   1215
      Left            =   6480
      TabIndex        =   20
      Top             =   4200
      Width           =   2295
      Begin VB.CheckBox SoundsChk 
         BackColor       =   &H80000005&
         Caption         =   "Enable Sounds"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Enable sound effects"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox MusicChk 
         BackColor       =   &H80000005&
         Caption         =   "Enable Music"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Enable background music"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox ReverseChk 
         BackColor       =   &H80000005&
         Caption         =   "Reverse Speakers"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Reverse the left and right speakers (tick if the game pans in the wrong direction)"
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000005&
      Caption         =   "Video"
      Height          =   3735
      Left            =   6480
      TabIndex        =   6
      Top             =   360
      Width           =   2295
      Begin VB.CheckBox AltRenderChk 
         BackColor       =   &H80000005&
         Caption         =   "Alternate Render"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Tick this if the graphics are rendering incorrectly (results in a performance decrease)"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox AltRenderMapChk 
         BackColor       =   &H80000005&
         Caption         =   "Alternate Render Map"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Tick this if the map is rendering incorrectly (results in a performance decrease and bad map lighting)"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CheckBox AltRenderTextChk 
         BackColor       =   &H80000005&
         Caption         =   "Alternate Render Text"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Tick this if the text is rendering incorrectly (results in a performance decrease)"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox FullScreenChk 
         BackColor       =   &H80000005&
         Caption         =   "Fullscreen Mode"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Enable full-screen mode"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox MotionBlurChk 
         BackColor       =   &H80000005&
         Caption         =   "Motion Blur Effects"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Enable motion blur effects"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox WeatherChk 
         BackColor       =   &H80000005&
         Caption         =   "Weather Effects"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Draw weather effects (rain, fog, snow, etc)"
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox ChatBubblesChk 
         BackColor       =   &H80000005&
         Caption         =   "Chat Bubbles"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Draw character chat bubbles above their head - Warning! Some NPCs only chat with chat bubbles!"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox Bit32Chk 
         BackColor       =   &H80000005&
         Caption         =   "32-Bit Pixel Depth"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Use 32-bit pixel depth resolution instead of 16-bit (slower, but very slightly better coloring)"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox FPSCapChk 
         BackColor       =   &H80000005&
         Caption         =   "Enable FPS Cap"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Tick to enable a FPS cap which will free up the CPU to other processes (recommended)"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox FPSTxt 
         BackColor       =   &H8000000E&
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "The highest FPS to aim for (recommended 60 since values higher than 60 will be hardly noticeable)"
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox VSyncChk 
         BackColor       =   &H80000005&
         Caption         =   "V-Sync"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Syncronize the verticle refresh to the monitor (not recommended)"
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox CompressChk 
         BackColor       =   &H80000005&
         Caption         =   "Texture Compression"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   $"frmMain.frx":17D2A
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "FPS Cap:"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   3360
         Width           =   675
      End
   End
   Begin VB.CommandButton DefaultsCmd 
      Caption         =   "Restore Default Controls"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Restore the default controls (does not apply to the Video and Audio game settings!)"
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000005&
      Caption         =   "Game Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000005&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
   Begin VB.CommandButton CloseCmd 
      Caption         =   "Close Without Saving"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Close the configuration application without saving changes"
      Top             =   6360
      Width           =   2535
   End
   Begin VB.CommandButton SaveCmd 
      Caption         =   "Save Changes"
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Save changes to the game configuration"
      Top             =   6000
      Width           =   2535
   End
   Begin VB.CommandButton LoadCmd 
      Caption         =   "Load Saved Settings"
      Height          =   255
      Left            =   6360
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Restore the configuration settings currently in the Game.ini file"
      Top             =   5640
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const KeyPress_Shift As Integer = 2 ^ 12
Private Const KeyPress_Control As Integer = 2 ^ 13
Private Const KeyPress_Alt As Integer = 2 ^ 14

Private Type KeyDefinitions
    MiniMap As Integer
    PickUpObj As Integer
    QuickBar(1 To 12) As Integer
    Attack As Integer
    ChatBufferUp As Integer
    ChatBufferDown As Integer
    InventoryWindow As Integer
    QuickBarWindow As Integer
    ChatWindow As Integer
    StatWindow As Integer
    MenuWindow As Integer
    ZoomIn As Integer
    ZoomOut As Integer
    MoveNorth As Integer
    MoveEast As Integer
    MoveSouth As Integer
    MoveWest As Integer
    ResetGUI As Integer
    QuickTarget As Integer
    QuickReply As Integer
End Type
Private KeyDefinitions As KeyDefinitions

Private HasChanged As Boolean

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Function KeyName(ByVal KeyCode As Integer) As String
Dim s As String

    'Check for shift, alt and control
    If KeyCode And KeyPress_Shift Then s = "(SHIFT)": KeyCode = KeyCode Xor KeyPress_Shift
    If KeyCode And KeyPress_Control Then s = "(CTRL)": KeyCode = KeyCode Xor KeyPress_Control
    If KeyCode And KeyPress_Alt Then s = "(ALT)": KeyCode = KeyCode Xor KeyPress_Alt
    
    'Remove the Shift, Control and Alt bits
    KeyCode = KeyCode And 2047

    'Check for known names
    Select Case KeyCode
        
        Case 1          'Left-click
        Case 2          'Right-click
        Case 3          'Cancel
        Case 4          'Middle-click
        
        Case 16, 160
            KeyName = "(SHIFT)"
        Case 17, 162
            KeyName = "(CTRL)"
        Case 18, 164
            KeyName = "(ALT)"
            
        Case 8
            KeyName = "(BACK)"
        Case 9
            KeyName = "(TAB)"
        Case 12
            KeyName = "(CLEAR)"
        Case 13
            KeyName = "(RETURN)"
        Case 19
            KeyName = "(PAUSE)"
        Case 20
            KeyName = "(CAP)"
        Case 27
            KeyName = "(ESC)"
        Case 32
            KeyName = "(SPACE)"
        Case 33
            KeyName = "(PGUP)"
        Case 34
            KeyName = "(PGDOWN)"
        Case 35
            KeyName = "(END)"
        Case 36
            KeyName = "(HOME)"
        Case 37
            KeyName = "(LEFT)"
        Case 38
            KeyName = "(UP)"
        Case 39
            KeyName = "(RIGHT)"
        Case 40
            KeyName = "(DOWN)"
        Case 41
            KeyName = "(SELECT)"
        Case 42
            KeyName = "(PRINT)"
        Case 43
            KeyName = "(EXECUTE)"
        Case 44
            KeyName = "(SNAPSHOT)"
        Case 45
            KeyName = "(INS)"
        Case 46
            KeyName = "(DEL)"
        Case 47
            KeyName = "(HELP)"
        Case 112 To 127
            KeyName = "F" & (KeyCode - 111)
        Case 144
            KeyName = "(NUMLCK)"
        Case 145
            KeyName = "(SCRLLCK)"
        Case Else
            If KeyCode >= 32 Then
                KeyName = UCase$(Chr$(KeyCode))
            Else
                KeyName = "(UNKNOWN)"
            End If
    End Select
    
    If s <> vbNullString Then
        KeyName = s & " + " & KeyName
    End If
    
End Function

Private Function GetKeyValue(ByVal KeyCode As Integer) As Integer

    'Only add on Shift, Control or Alt combos if they aren't pressed
    If KeyCode <> 16 Then
        If KeyCode <> 17 Then
            If KeyCode <> 18 Then
                If GetAsyncKeyState(16) Then GetKeyValue = GetKeyValue Or KeyPress_Shift
                If GetAsyncKeyState(17) Then GetKeyValue = GetKeyValue Or KeyPress_Control
                If GetAsyncKeyState(18) Then GetKeyValue = GetKeyValue Or KeyPress_Alt
            End If
        End If
    End If
    
    'Add on the keycode
    GetKeyValue = GetKeyValue Or KeyCode
    
    'Clear the previous alt/control/shift key presses
    GetAsyncKeyState 16
    GetAsyncKeyState 17
    GetAsyncKeyState 18

End Function

Private Sub AltRenderChk_Click()

    HasChanged = True

End Sub

Private Sub AltRenderMapChk_Click()

    HasChanged = True

End Sub

Private Sub AltRenderTextChk_Click()

    HasChanged = True

End Sub

Private Sub Bit32Chk_Click()

    HasChanged = True

End Sub

Private Sub ChatBubblesChk_Click()

    HasChanged = True

End Sub

Private Sub CloseCmd_Click()

    Unload Me

End Sub

Private Sub CompressChk_Click()

    HasChanged = True

End Sub

Private Sub Form_Load()

    'Set the file paths
    InitFilePaths

    'Clear the key cache
    Input_Keys_ClearQueue
    
    'Load the key config
    Input_KeyDefinitions_Load
    
    'Load the game config
    LoadConfig

End Sub

Private Sub Input_Keys_ClearQueue()

'*****************************************************************
'Clears the GetAsyncKeyState queue to prevent key presses from a long time
' ago falling into "have been pressed"
'*****************************************************************
Dim i As Long

    For i = 0 To 255
        GetAsyncKeyState i
    Next i

End Sub

Private Sub Input_KeyDefinitions_Load()

'*****************************************************************
'Load the key definitions
'*****************************************************************
Dim i As Long

    KeyDefinitions.Attack = Val(Var_Get(DataPath & "Game.ini", "INPUT", "Attack"))
    KeyDefinitions.ChatBufferDown = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatBufferDown"))
    KeyDefinitions.ChatBufferUp = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatBufferUp"))
    KeyDefinitions.ChatWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ChatWindow"))
    KeyDefinitions.InventoryWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "InventoryWindow"))
    KeyDefinitions.MenuWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MenuWindow"))
    KeyDefinitions.MiniMap = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MiniMap"))
    KeyDefinitions.MoveEast = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveEast"))
    KeyDefinitions.MoveNorth = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveNorth"))
    KeyDefinitions.MoveSouth = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveSouth"))
    KeyDefinitions.MoveWest = Val(Var_Get(DataPath & "Game.ini", "INPUT", "MoveWest"))
    KeyDefinitions.PickUpObj = Val(Var_Get(DataPath & "Game.ini", "INPUT", "PickUpObj"))
    KeyDefinitions.QuickBarWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickBarWindow"))
    KeyDefinitions.StatWindow = Val(Var_Get(DataPath & "Game.ini", "INPUT", "StatWindow"))
    KeyDefinitions.ZoomIn = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ZoomIn"))
    KeyDefinitions.ZoomOut = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ZoomOut"))
    KeyDefinitions.ResetGUI = Val(Var_Get(DataPath & "Game.ini", "INPUT", "ResetGUI"))
    KeyDefinitions.QuickTarget = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickTarget"))
    KeyDefinitions.QuickReply = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickReply"))
    For i = 1 To 12
        KeyDefinitions.QuickBar(i) = Val(Var_Get(DataPath & "Game.ini", "INPUT", "QuickBar" & i))
    Next i
    
    'Only used in the config editor
    SetTextBoxes
    
End Sub

Private Sub Input_KeyDefinitions_Save()

'*****************************************************************
'Save the key definitions
'*****************************************************************
Dim i As Long

    Var_Write DataPath & "Game.ini", "INPUT", "Attack", KeyDefinitions.Attack
    Var_Write DataPath & "Game.ini", "INPUT", "ChatBufferDown", KeyDefinitions.ChatBufferDown
    Var_Write DataPath & "Game.ini", "INPUT", "ChatBufferUp", KeyDefinitions.ChatBufferUp
    Var_Write DataPath & "Game.ini", "INPUT", "ChatWindow", KeyDefinitions.ChatWindow
    Var_Write DataPath & "Game.ini", "INPUT", "InventoryWindow", KeyDefinitions.InventoryWindow
    Var_Write DataPath & "Game.ini", "INPUT", "MenuWindow", KeyDefinitions.MenuWindow
    Var_Write DataPath & "Game.ini", "INPUT", "MiniMap", KeyDefinitions.MiniMap
    Var_Write DataPath & "Game.ini", "INPUT", "MoveEast", KeyDefinitions.MoveEast
    Var_Write DataPath & "Game.ini", "INPUT", "MoveNorth", KeyDefinitions.MoveNorth
    Var_Write DataPath & "Game.ini", "INPUT", "MoveSouth", KeyDefinitions.MoveSouth
    Var_Write DataPath & "Game.ini", "INPUT", "MoveWest", KeyDefinitions.MoveWest
    Var_Write DataPath & "Game.ini", "INPUT", "PickUpObj", KeyDefinitions.PickUpObj
    Var_Write DataPath & "Game.ini", "INPUT", "QuickBarWindow", KeyDefinitions.QuickBarWindow
    Var_Write DataPath & "Game.ini", "INPUT", "StatWindow", KeyDefinitions.StatWindow
    Var_Write DataPath & "Game.ini", "INPUT", "ZoomIn", KeyDefinitions.ZoomIn
    Var_Write DataPath & "Game.ini", "INPUT", "ZoomOut", KeyDefinitions.ZoomOut
    Var_Write DataPath & "Game.ini", "INPUT", "ResetGUI", KeyDefinitions.ResetGUI
    Var_Write DataPath & "Game.ini", "INPUT", "QuickTarget", KeyDefinitions.QuickTarget
    Var_Write DataPath & "Game.ini", "INPUT", "QuickReply", KeyDefinitions.QuickReply
    For i = 1 To 12
        Var_Write DataPath & "Game.ini", "INPUT", "QuickBar" & i, KeyDefinitions.QuickBar(i)
    Next i

End Sub

Private Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
Dim sSpaces As String

    sSpaces = Space$(1000)
    GetPrivateProfileString Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    Var_Get = RTrim$(sSpaces)
    If Len(Var_Get) > 0 Then
        Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
    Else
        Var_Get = vbNullString
    End If
    
End Function

Private Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    WritePrivateProfileString Main, Var, Value, File

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If HasChanged Then
        If MsgBox("Are you sure you wish to quit? Any unsaved changes will be lost!", vbYesNo) = vbNo Then
            Cancel = 1
            UnloadMode = 0
            Exit Sub
        End If
    End If
    
End Sub

Private Sub Form_Resize()

    Me.Width = 9135
    Me.Height = 7710

End Sub

Private Sub FPSCapChk_Click()

    HasChanged = True

    'Enable/disable FPS limit
    FPSTxt.Enabled = (FPSCapChk.Value = 1)

End Sub

Private Sub FPSTxt_GotFocus()

    'Set the high-light
    FPSTxt.BackColor = &H80000013

End Sub

Private Sub FPSTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    'Clear the key so no text will be entered in the control
    KeyCode = 0
    
End Sub

Private Sub FPSTxt_KeyPress(KeyAscii As Integer)

    'Check for numeric or backspace
    If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> 8 Then KeyAscii = 0
    
    If KeyAscii > 0 Then HasChanged = True
    
End Sub

Private Sub FPSTxt_LostFocus()
    
    'Remove the high-light
    FPSTxt.BackColor = &H80000005

End Sub

Private Sub FullScreenChk_Click()

    HasChanged = True

End Sub

Private Sub KeyTxt_GotFocus(Index As Integer)

    'Set the high-light
    KeyTxt(Index).BackColor = &H80000013

End Sub

Private Sub SetTextBoxes()
Dim i As Long

    'Set the values in the text boxes
    With KeyDefinitions
        KeyTxt(0).Text = KeyName(.Attack)
        KeyTxt(1).Text = KeyName(.PickUpObj)
        KeyTxt(2).Text = KeyName(.MoveNorth)
        KeyTxt(3).Text = KeyName(.MoveSouth)
        KeyTxt(4).Text = KeyName(.MoveWest)
        KeyTxt(5).Text = KeyName(.MoveEast)
        KeyTxt(6).Text = KeyName(.ChatBufferUp)
        KeyTxt(7).Text = KeyName(.ChatBufferDown)
        KeyTxt(8).Text = KeyName(.ZoomIn)
        KeyTxt(9).Text = KeyName(.ZoomOut)
        KeyTxt(10).Text = KeyName(.ChatWindow)
        KeyTxt(11).Text = KeyName(.InventoryWindow)
        KeyTxt(12).Text = KeyName(.MenuWindow)
        KeyTxt(13).Text = KeyName(.QuickBarWindow)
        KeyTxt(14).Text = KeyName(.StatWindow)
        KeyTxt(15).Text = KeyName(.MiniMap)
        For i = 1 To 12
            KeyTxt(15 + i).Text = KeyName(.QuickBar(i))
        Next i
        KeyTxt(28).Text = KeyName(.ResetGUI)
        KeyTxt(29).Text = KeyName(.QuickTarget)
        KeyTxt(30).Text = KeyName(.QuickReply)
    End With

End Sub

Private Sub DefaultsCmd_Click()
Dim i As Long

    If MsgBox("Are you sure you wish to restore the default control settings?" & vbNewLine & "Any unsaved changes will be lost!", vbYesNo) = vbNo Then Exit Sub

    'Set to the default settings used
    With KeyDefinitions
        .Attack = 17
        .PickUpObj = 18
        .MoveNorth = 87
        .MoveEast = 68
        .MoveSouth = 83
        .MoveWest = 65
        .ChatBufferUp = 33
        .ChatBufferDown = 34
        .ZoomIn = 104
        .ZoomOut = 98
        .ChatWindow = KeyPress_Control Or 67
        .InventoryWindow = KeyPress_Control Or 87
        .MenuWindow = 27
        .QuickBarWindow = KeyPress_Control Or 81
        .StatWindow = KeyPress_Control Or 83
        .MiniMap = 9
        For i = 1 To 12
            .QuickBar(i) = 111 + i
        Next i
        .ResetGUI = KeyPress_Shift Or 123
        .QuickTarget = 69
        .QuickReply = 82
    End With
    
    'Display the changes
    SetTextBoxes
    HasChanged = False

End Sub

Private Sub KeyTxt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Long
        
    'Get the value
    i = GetKeyValue(KeyCode)
    
    'Display the value
    KeyTxt(Index).Text = KeyName(i)
    
    'Set the value to the appropriate variable
    With KeyDefinitions
        Select Case Index
            Case 0: .Attack = i
            Case 1: .PickUpObj = i
            Case 2: .MoveNorth = i
            Case 3: .MoveSouth = i
            Case 4: .MoveWest = i
            Case 5: .MoveEast = i
            Case 6: .ChatBufferUp = i
            Case 7: .ChatBufferDown = i
            Case 8: .ZoomIn = i
            Case 9: .ZoomOut = i
            Case 10: .ChatWindow = i
            Case 11: .InventoryWindow = i
            Case 12: .MenuWindow = i
            Case 13: .QuickBarWindow = i
            Case 14: .StatWindow = i
            Case 15: .MiniMap = i
            Case 16 To 27: .QuickBar(Index - 15) = i
            Case 28: .ResetGUI = i
            Case 29: .QuickTarget = i
            Case 30: .QuickReply = i
        End Select
    End With
    
    'Clear the key so no text will be entered in the control
    KeyCode = 0
    Shift = 0
    
    'A change has been made
    HasChanged = True
    
End Sub

Private Sub KeyTxt_KeyPress(Index As Integer, KeyAscii As Integer)

    'Clear the key so no text will be entered in the control
    KeyAscii = 0

End Sub

Private Sub KeyTxt_LostFocus(Index As Integer)
    
    'Remove the high-light
    KeyTxt(Index).BackColor = &H80000005

End Sub

Private Sub LoadCmd_Click()

    If MsgBox("Are you sure you wish to load the last saved settings?", vbYesNo) = vbNo Then Exit Sub
    Input_KeyDefinitions_Load
    LoadConfig
    HasChanged = False

End Sub

Private Sub MotionBlurChk_Click()

    HasChanged = True

End Sub

Private Sub MusicChk_Click()

    HasChanged = True

End Sub

Private Sub ReverseChk_Click()

    HasChanged = True

End Sub

Private Sub SaveCmd_Click()

    If MsgBox("Are you sure you wish to save the current settings?", vbYesNo) = vbNo Then Exit Sub
    Input_KeyDefinitions_Save
    SaveConfig
    HasChanged = False

End Sub

Private Sub LoadConfig()
Dim f As String

    f = DataPath & "Game.ini"
    
    FullScreenChk.Value = IIf(Val(Var_Get(f, "INIT", "Windowed")) = 0, 1, 0)
    MotionBlurChk.Value = IIf(Val(Var_Get(f, "INIT", "UseMotionBlur")) <> 0, 1, 0)
    WeatherChk.Value = IIf(Val(Var_Get(f, "INIT", "UseWeather")) <> 0, 1, 0)
    ChatBubblesChk.Value = IIf(Val(Var_Get(f, "INIT", "DisableChatBubbles")) = 0, 1, 0)
    Bit32Chk.Value = IIf(Val(Var_Get(f, "INIT", "32bit")) <> 0, 1, 0)
    VSyncChk.Value = IIf(Val(Var_Get(f, "INIT", "VSync")) <> 0, 1, 0)
    CompressChk.Value = IIf(Val(Var_Get(f, "INIT", "TextureCompression")) <> 0, 1, 0)
    
    AltRenderChk.Value = IIf(Val(Var_Get(f, "INIT", "AlternateRender")) <> 0, 1, 0)
    AltRenderMapChk.Value = IIf(Val(Var_Get(f, "INIT", "AlternateRenderMap")) <> 0, 1, 0)
    AltRenderTextChk.Value = IIf(Val(Var_Get(f, "INIT", "AlternateRenderText")) <> 0, 1, 0)
    
    FPSCapChk.Value = IIf(Val(Var_Get(f, "INIT", "FPSCap")) <> 0, 1, 0)
    FPSTxt.Text = Val(Var_Get(f, "INIT", "FPSCap"))
    FPSTxt.Enabled = (FPSCapChk.Value = 1)
    
    SoundsChk.Value = IIf(Val(Var_Get(f, "INIT", "UseSfx")) <> 0, 1, 0)
    ReverseChk.Value = IIf(Val(Var_Get(f, "INIT", "ReverseSound")) <> 0, 1, 0)
    MusicChk.Value = IIf(Val(Var_Get(f, "INIT", "UseMusic")) <> 0, 1, 0)

End Sub

Private Sub SaveConfig()
Dim f As String

    f = DataPath & "Game.ini"

    Var_Write f, "INIT", "Windowed", IIf(FullScreenChk.Value = 0, 1, 0)
    Var_Write f, "INIT", "UseMotionBlur", MotionBlurChk.Value
    Var_Write f, "INIT", "UseWeather", WeatherChk.Value
    Var_Write f, "INIT", "DisableChatBubbles", IIf(ChatBubblesChk.Value = 0, 1, 0)
    Var_Write f, "INIT", "32bit", Bit32Chk.Value
    Var_Write f, "INIT", "VSync", VSyncChk.Value
    Var_Write f, "INIT", "TextureCompression", CompressChk.Value
    
    Var_Write f, "INIT", "AlternateRender", AltRenderChk.Value
    Var_Write f, "INIT", "AlternateRenderMap", AltRenderMapChk.Value
    Var_Write f, "INIT", "AlternateRenderText", AltRenderTextChk.Value
    
    If FPSCapChk.Value = 1 Then
        Var_Write f, "INIT", "FPSCap", Val(FPSTxt.Text)
    Else
        Var_Write f, "INIT", "FPSCap", "0"
    End If
    
    Var_Write f, "INIT", "UseSfx", SoundsChk.Value
    Var_Write f, "INIT", "ReverseSound", ReverseChk.Value
    Var_Write f, "INIT", "UseMusic", MusicChk.Value

End Sub

Private Sub SoundsChk_Click()

    HasChanged = True

End Sub

Private Sub VSyncChk_Click()

    HasChanged = True

End Sub

Private Sub WeatherChk_Click()

    HasChanged = True

End Sub
