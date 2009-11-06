VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Object Editor"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375.001
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   2640
      TabIndex        =   165
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   2640
      TabIndex        =   164
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   2640
      TabIndex        =   163
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   2640
      TabIndex        =   162
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   2640
      TabIndex        =   161
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   2640
      TabIndex        =   160
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   31
      Left            =   8520.001
      TabIndex        =   158
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   8520.001
      TabIndex        =   156
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   29
      Left            =   8520.001
      TabIndex        =   154
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   28
      Left            =   8520.001
      TabIndex        =   152
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   27
      Left            =   8520.001
      TabIndex        =   150
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   26
      Left            =   8520.001
      TabIndex        =   148
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   8520.001
      TabIndex        =   146
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   24
      Left            =   8520.001
      TabIndex        =   144
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   23
      Left            =   8520.001
      TabIndex        =   142
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   8520.001
      TabIndex        =   140
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   8520.001
      TabIndex        =   138
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   8520.001
      TabIndex        =   136
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   8520.001
      TabIndex        =   134
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   8520.001
      TabIndex        =   132
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   8520.001
      TabIndex        =   130
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   8520.001
      TabIndex        =   128
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   6720
      TabIndex        =   126
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   6720
      TabIndex        =   124
      Top             =   3840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   6720
      TabIndex        =   122
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   6720
      TabIndex        =   120
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   6720
      TabIndex        =   118
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   6720
      TabIndex        =   116
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   6720
      TabIndex        =   114
      Top             =   2640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   6720
      TabIndex        =   112
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   6720
      TabIndex        =   110
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   6720
      TabIndex        =   108
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   6720
      TabIndex        =   106
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   6720
      TabIndex        =   104
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   6720
      TabIndex        =   102
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6720
      TabIndex        =   100
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   6720
      TabIndex        =   98
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6720
      TabIndex        =   96
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox RepTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   4800
      TabIndex        =   93
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepPercTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   6360
      TabIndex        =   92
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   2640
      TabIndex        =   90
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   2640
      TabIndex        =   88
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ReqTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   4800
      TabIndex        =   87
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ReqTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4800
      TabIndex        =   85
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ReqTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   83
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ReqTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   81
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ReqTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   79
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox ReqTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4800
      TabIndex        =   77
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox StackTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   73
      ToolTipText     =   "Name of the object"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox FXTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   71
      ToolTipText     =   "Name of the object"
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox ProjecTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   69
      ToolTipText     =   "Name of the object"
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox RotTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   67
      ToolTipText     =   "Name of the object"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox RangeTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   66
      ToolTipText     =   "Name of the object"
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox RepPercTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   6360
      TabIndex        =   62
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   4800
      TabIndex        =   60
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepPercTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   6360
      TabIndex        =   58
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4800
      TabIndex        =   56
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepPercTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   6360
      TabIndex        =   54
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   52
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepPercTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   6360
      TabIndex        =   50
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   4800
      TabIndex        =   48
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox RepPercTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   6360
      TabIndex        =   46
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   2640
      TabIndex        =   45
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   1080
      TabIndex        =   35
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   1080
      TabIndex        =   34
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   1080
      TabIndex        =   33
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1080
      TabIndex        =   32
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   30
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   28
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Classes 
      BackColor       =   &H80000005&
      Caption         =   "Rogue"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   27
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox Classes 
      BackColor       =   &H80000005&
      Caption         =   "Mage"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   26
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox Classes 
      BackColor       =   &H80000005&
      Caption         =   "Warrior"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   25
      Top             =   360
      Width           =   855
   End
   Begin VB.ComboBox ObjTypeCombo 
      Height          =   315
      ItemData        =   "frmMain.frx":17D70
      Left            =   2400
      List            =   "frmMain.frx":17D72
      TabIndex        =   23
      Text            =   "Object Type"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox SelectObjCombo 
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Text            =   "Select an Object"
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox SpriteTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox RepTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   4800
      TabIndex        =   16
      ToolTipText     =   "How much HP is replenished upon using/equipting (positive or negative)"
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox WeaponTypeTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      ToolTipText     =   "Type of weapon (only if the object type is weapon)"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox PriceTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   11
      ToolTipText     =   "The price of the object when purchased from a store"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   9
      ToolTipText     =   "Grh number of the graphic the object will use"
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Timer RenderTimer 
      Interval        =   50
      Left            =   8520.001
      Top             =   5640
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Name of the object"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8400.001
      Top             =   5640
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
      Left            =   3120
      TabIndex        =   172
      Top             =   600
      Width           =   705
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   171
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   170
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   169
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   168
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   167
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   166
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   159
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   157
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   155
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   153
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   151
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   149
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   147
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   145
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   143
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   141
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   139
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   137
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   135
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   133
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   131
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   7920
      TabIndex        =   129
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   127
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   125
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   123
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   121
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   119
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   117
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   115
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   113
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   111
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   109
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   107
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   105
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   103
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   101
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   99
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label StatLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   6120
      TabIndex        =   97
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   95
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepPercLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      TabIndex        =   94
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   91
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   89
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ReqLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   86
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ReqLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   84
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ReqLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   82
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ReqLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   80
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ReqLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   78
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label ReqLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   76
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line28 
      X1              =   376
      X2              =   272
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Line Line27 
      X1              =   376
      X2              =   376
      Y1              =   296
      Y2              =   160
   End
   Begin VB.Line Line26 
      X1              =   376
      X2              =   272
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Req:"
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
      Index           =   16
      Left            =   4440
      TabIndex        =   75
      Top             =   2400
      Width           =   510
   End
   Begin VB.Line Line25 
      X1              =   272
      X2              =   272
      Y1              =   296
      Y2              =   160
   End
   Begin VB.Line Line24 
      X1              =   272
      X2              =   272
      Y1              =   152
      Y2              =   8
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Req. Class:"
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
      Index           =   15
      Left            =   4200
      TabIndex        =   74
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line23 
      X1              =   376
      X2              =   272
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line22 
      X1              =   376
      X2              =   376
      Y1              =   152
      Y2              =   8
   End
   Begin VB.Line Line21 
      X1              =   376
      X2              =   272
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stack:"
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
      Left            =   240
      TabIndex        =   72
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FX:"
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
      Left            =   2760
      TabIndex        =   70
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projectile Grh:"
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
      Left            =   1800
      TabIndex        =   68
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rotation:"
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
      Left            =   2280
      TabIndex        =   65
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range:"
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
      Left            =   2400
      TabIndex        =   64
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label RepPercLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      TabIndex        =   63
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   61
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepPercLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      TabIndex        =   59
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   57
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepPercLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      TabIndex        =   55
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   53
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepPercLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      TabIndex        =   51
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   49
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label RepPercLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   2040
      TabIndex        =   44
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   43
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   42
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   41
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   40
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   38
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   37
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   36
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line20 
      X1              =   8
      X2              =   8
      Y1              =   416
      Y2              =   256
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
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
      Left            =   1800
      TabIndex        =   24
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Line19 
      X1              =   8
      X2              =   264
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line18 
      X1              =   264
      X2              =   264
      Y1              =   8
      Y2              =   72
   End
   Begin VB.Line Line17 
      X1              =   264
      X2              =   8
      Y1              =   72
      Y2              =   72
   End
   Begin VB.Line Line16 
      X1              =   8
      X2              =   8
      Y1              =   8
      Y2              =   72
   End
   Begin VB.Line Line15 
      X1              =   8
      X2              =   8
      Y1              =   80
      Y2              =   248
   End
   Begin VB.Line Line14 
      X1              =   8
      X2              =   264
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line Line13 
      X1              =   264
      X2              =   264
      Y1              =   80
      Y2              =   248
   End
   Begin VB.Line Line12 
      X1              =   8
      X2              =   264
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line Line11 
      X1              =   480
      X2              =   272
      Y1              =   304
      Y2              =   304
   End
   Begin VB.Line Line10 
      X1              =   272
      X2              =   272
      Y1              =   304
      Y2              =   416
   End
   Begin VB.Line Line9 
      X1              =   480
      X2              =   272
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line8 
      X1              =   264
      X2              =   8
      Y1              =   416
      Y2              =   416
   End
   Begin VB.Line Line7 
      X1              =   264
      X2              =   264
      Y1              =   416
      Y2              =   256
   End
   Begin VB.Line Line6 
      X1              =   264
      X2              =   8
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line Line5 
      X1              =   480
      X2              =   480
      Y1              =   416
      Y2              =   304
   End
   Begin VB.Line Line4 
      X1              =   384
      X2              =   384
      Y1              =   8
      Y2              =   296
   End
   Begin VB.Line Line3 
      X1              =   384
      X2              =   616
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Line Line2 
      X1              =   616
      X2              =   616
      Y1              =   296
      Y2              =   8
   End
   Begin VB.Line Line1 
      X1              =   384
      X2              =   616
      Y1              =   8
      Y2              =   8
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
      Index           =   6
      Left            =   5880
      TabIndex        =   21
      Top             =   240
      Width           =   600
   End
   Begin VB.Label SpriteLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   480
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   495
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
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3840
      TabIndex        =   18
      Top             =   2520
      Width           =   90
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label RepLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
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
      Left            =   4200
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   4200
      TabIndex        =   14
      Top             =   4560
      Width           =   1455
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1800
      TabIndex        =   12
      Top             =   2520
      Width           =   1260
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   1800
      TabIndex        =   10
      Top             =   1920
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
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
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   600
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
      Left            =   1920
      TabIndex        =   6
      Top             =   600
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
      Left            =   1080
      TabIndex        =   5
      Top             =   600
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
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   1680
      TabIndex        =   3
      Top             =   1200
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
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
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
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

Private Sub DeleteLbl_Click()

    'Confirm
    If MsgBox("Are you sure you wish to delete Object " & Objnumber & " ?", vbYesNo) = vbNo Then Exit Sub
    'Check for overwrite
    If ObjExist(Objnumber, True) Then
        MsgBox "DELETED!", vbOKOnly, "Gone!"
        Engine_Load_ObjCombo
    Else
        MsgBox "Delete failed! Object " & Objnumber & " does not exist!", vbOKOnly, "Save failed!"
    End If
End Sub

Private Sub LoadLbl_Click()
Dim TempStr() As String
    'Confirm
    If MsgBox("Are you sure you wish to load another object?" & vbCrLf & "Any changes made to the current object will be lost!", vbYesNo) = vbNo Then Exit Sub
    
    'Load object
        TempStr = Split(frmMain.SelectObjCombo.Text, "-")
    
    Editor_LoadOBJ Val(TempStr(0))

End Sub

Private Sub RenderTimer_Timer()

    'Usually I dont use timers, but we dont want to create a loop since we dont need intensive drawing
    ElapsedTime = 50
    
    'Render
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
        Engine_Render_Grh ObjGrh, 1, 1, 0, 1, , -1, -1, -1, -1
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
    If ObjExist(RetNumber, False) Then
        'are you sure you want to overwrite?
        If MsgBox("Object " & RetNumber & " exists! Are you sure you wish to overwrite?", vbOKCancel, "Save as?") = vbOK Then
            Editor_SaveOBJ RetNumber
        Else
            MsgBox "Save failed!", vbOKOnly, "Save failed!"
        End If
    Else
        'Save
        Editor_SaveOBJ RetNumber
    End If

End Sub

Private Sub SaveLbl_Click()

    'Confirm
    If MsgBox("Are you sure you wish to save changes to Object " & OpenIndex & "?", vbYesNo) = vbNo Then Exit Sub
    
    'Save
    Editor_SaveOBJ OpenIndex

End Sub


Private Sub StatTxt_Change(Index As Integer)
Dim i As Long
On Error GoTo ErrOut

    i = Val(StatTxt(Index).Text)

    Exit Sub
    
ErrOut:

    StatTxt(Index).Text = "0"

End Sub

Private Sub NameTxt_Change()
Dim i As String
On Error GoTo ErrOut
    i = Trim$(NameTxt.Text)
    Exit Sub
ErrOut:
Return
End Sub

Private Sub GrhTxt_Change()
Dim i As Long
On Error GoTo ErrOut
    
    'i = Val(GrhTxt.Text)
    If Val(GrhTxt.Text) <= 0 Then ObjGrh.GrhIndex = 0 Else Engine_Init_Grh ObjGrh, Val(GrhTxt.Text)

    Exit Sub

ErrOut:

    GrhTxt.Text = "0"
    ObjGrh.GrhIndex = 0

End Sub

Private Sub PriceTxt_Change()
    Dim i As Long
    On Error GoTo ErrOut
        i = Val(PriceTxt.Text)
        Exit Sub
ErrOut:
    Return
End Sub

Private Sub WeaponTypeTxt_Change()
Dim i As Byte
On Error GoTo ErrOut

    i = Val(WeaponTypeTxt.Text)

    Exit Sub
    
ErrOut:

    WeaponTypeTxt.Text = "0"
    
End Sub

Private Sub WeaponTypeLbl_Click()
Dim s As String

    s = "Weapon types:" & vbCrLf
    s = s & "0: Hands" & vbCrLf
    s = s & "1: Staff" & vbCrLf
    s = s & "2: Dagger" & vbCrLf
    s = s & "3: Sword" & vbCrLf
    s = s & "4: Throwing"
    MsgBox s, vbOKOnly
End Sub

