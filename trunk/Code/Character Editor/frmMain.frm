VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "vbGORE Character File Editor"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12705
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":17D2A
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   847
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   56
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   265
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   55
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   264
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   54
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   263
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   53
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   262
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   52
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   261
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   51
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   260
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   50
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   259
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   49
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   244
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   48
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   243
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   47
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   242
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   46
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   241
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   45
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   240
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   44
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   239
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   43
      Left            =   10680
      MaxLength       =   10
      TabIndex        =   238
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   42
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   223
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   41
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   222
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   40
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   221
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   39
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   220
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   38
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   219
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   37
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   218
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   36
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   217
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   35
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   202
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   34
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   201
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   33
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   200
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   32
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   199
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   31
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   198
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   30
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   197
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   29
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   196
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   28
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   181
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   27
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   180
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   26
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   179
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   25
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   178
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   24
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   177
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   23
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   176
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   22
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   175
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   21
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   160
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   20
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   159
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   19
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   158
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   18
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   157
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   17
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   156
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   16
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   155
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   15
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   154
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   14
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   139
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   13
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   138
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   12
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   137
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   11
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   136
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   10
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   135
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   9
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   134
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   8
      Left            =   2880
      MaxLength       =   10
      TabIndex        =   133
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   7
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   118
      Text            =   "0"
      Top             =   5760
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   6
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   117
      Text            =   "0"
      Top             =   5520
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   5
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   116
      Text            =   "0"
      Top             =   5280
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   4
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   115
      Text            =   "0"
      Top             =   5040
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   3
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   114
      Text            =   "0"
      Top             =   4800
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   2
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   113
      Text            =   "0"
      Top             =   4560
      Width           =   210
   End
   Begin VB.TextBox EquiptedTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   112
      Text            =   "0"
      Top             =   4320
      Width           =   210
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   56
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   258
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   55
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   257
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   54
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   256
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   53
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   255
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   52
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   254
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   51
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   253
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   50
      Left            =   11760
      MaxLength       =   10
      TabIndex        =   252
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   49
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   237
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   48
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   236
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   47
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   235
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   46
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   234
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   45
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   233
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   44
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   232
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   43
      Left            =   10200
      MaxLength       =   10
      TabIndex        =   231
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   42
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   216
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   41
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   215
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   40
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   214
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   39
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   213
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   38
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   212
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   37
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   211
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   36
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   210
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   35
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   195
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   34
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   194
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   33
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   193
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   32
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   192
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   31
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   191
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   30
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   190
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   29
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   189
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   28
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   174
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   27
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   173
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   26
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   172
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Index           =   25
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   171
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   170
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   169
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   5520
      MaxLength       =   10
      TabIndex        =   168
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   153
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   152
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   151
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   150
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   149
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   148
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   147
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   132
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   131
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   130
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   129
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   128
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   127
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   126
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   111
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   110
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   109
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      MaxLength       =   10
      TabIndex        =   108
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   107
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   840
      MaxLength       =   10
      TabIndex        =   106
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      MaxLength       =   10
      TabIndex        =   105
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   34
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   89
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   35
      Left            =   8640
      MaxLength       =   1
      TabIndex        =   90
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   36
      Left            =   9000
      MaxLength       =   1
      TabIndex        =   91
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   37
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   92
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   38
      Left            =   9720
      MaxLength       =   1
      TabIndex        =   93
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   39
      Left            =   10080
      MaxLength       =   1
      TabIndex        =   94
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   40
      Left            =   10440
      MaxLength       =   1
      TabIndex        =   95
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommDlg 
      Left            =   11640
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   33
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   88
      Text            =   "0"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   32
      Left            =   10440
      MaxLength       =   1
      TabIndex        =   87
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   31
      Left            =   10080
      MaxLength       =   1
      TabIndex        =   86
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   30
      Left            =   9720
      MaxLength       =   1
      TabIndex        =   85
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   29
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   84
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   28
      Left            =   9000
      MaxLength       =   1
      TabIndex        =   83
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   8640
      MaxLength       =   1
      TabIndex        =   82
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   81
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   80
      Text            =   "0"
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   10440
      MaxLength       =   1
      TabIndex        =   79
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   10080
      MaxLength       =   1
      TabIndex        =   78
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   9720
      MaxLength       =   1
      TabIndex        =   77
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   76
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   9000
      MaxLength       =   1
      TabIndex        =   75
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   8640
      MaxLength       =   1
      TabIndex        =   74
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   73
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   72
      Text            =   "0"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   10440
      MaxLength       =   1
      TabIndex        =   71
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   10080
      MaxLength       =   1
      TabIndex        =   70
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   9720
      MaxLength       =   1
      TabIndex        =   69
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   68
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   9000
      MaxLength       =   1
      TabIndex        =   67
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   8640
      MaxLength       =   1
      TabIndex        =   66
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   65
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   64
      Text            =   "0"
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   10440
      MaxLength       =   1
      TabIndex        =   63
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   10080
      MaxLength       =   1
      TabIndex        =   62
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   9720
      MaxLength       =   1
      TabIndex        =   61
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   9360
      MaxLength       =   1
      TabIndex        =   60
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   9000
      MaxLength       =   1
      TabIndex        =   59
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   8640
      MaxLength       =   1
      TabIndex        =   58
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   57
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox KnownSkillTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   56
      Text            =   "0"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   42
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   55
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   41
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   54
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox WeaponSlotTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12000
      TabIndex        =   97
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox ArmorSlotTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12000
      TabIndex        =   96
      Text            =   "0"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   56
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   251
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   55
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   250
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   49
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   230
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   48
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   229
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   42
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   209
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   41
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   208
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   35
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   188
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   34
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   187
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   28
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   167
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   166
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   146
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   145
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   125
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   124
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   360
      MaxLength       =   10
      TabIndex        =   104
      Text            =   "0"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   360
      MaxLength       =   10
      TabIndex        =   103
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   54
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   249
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   53
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   248
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   52
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   247
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   51
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   246
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   50
      Left            =   11280
      MaxLength       =   10
      TabIndex        =   245
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   47
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   228
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   46
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   227
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   45
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   226
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   44
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   225
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   43
      Left            =   9720
      MaxLength       =   10
      TabIndex        =   224
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   40
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   207
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   39
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   206
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   38
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   205
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   37
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   204
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   36
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   203
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   33
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   186
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   32
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   185
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   31
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   184
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   30
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   183
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   29
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   182
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   165
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   164
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   163
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   162
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   161
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   144
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   143
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   142
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   141
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   140
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   123
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   122
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   121
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   120
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   119
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   360
      MaxLength       =   10
      TabIndex        =   102
      Text            =   "0"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   360
      MaxLength       =   10
      TabIndex        =   101
      Text            =   "0"
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   360
      MaxLength       =   10
      TabIndex        =   100
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   360
      MaxLength       =   10
      TabIndex        =   99
      Text            =   "0"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox InventoryTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   360
      MaxLength       =   10
      TabIndex        =   98
      Text            =   "0"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox QuestTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11520
      TabIndex        =   12
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox CompletedQuestsTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Text            =   "0"
      Top             =   1680
      Width           =   12255
   End
   Begin VB.TextBox HeadHeadingTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8400
      TabIndex        =   10
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox HeadingTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6240
      TabIndex        =   9
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox WeaponTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4680
      TabIndex        =   8
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox BodyTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox HeadTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox HairTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox YTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10680
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox XTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10320
      TabIndex        =   3
      Text            =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox MapTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9120
      TabIndex        =   2
      Text            =   "0"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox GoldTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6240
      TabIndex        =   1
      Text            =   "0"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox PassTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Text            =   "0"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   40
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   53
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   39
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   52
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   38
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   51
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   37
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   50
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   36
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   49
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   35
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   48
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   34
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   47
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   33
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   46
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   32
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   45
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   31
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   44
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   30
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   43
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   29
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   42
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   28
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   41
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   27
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   40
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   26
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   39
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   25
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   38
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   37
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   36
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   35
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   34
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   33
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   32
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   31
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   30
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   29
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   28
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   27
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   26
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   25
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   23
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   22
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   21
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   240
      MaxLength       =   10
      TabIndex        =   20
      Text            =   "0"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   240
      MaxLength       =   10
      TabIndex        =   19
      Text            =   "0"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   240
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "0"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   240
      MaxLength       =   10
      TabIndex        =   17
      Text            =   "0"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   240
      MaxLength       =   10
      TabIndex        =   16
      Text            =   "0"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      MaxLength       =   10
      TabIndex        =   15
      Text            =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox StatTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      MaxLength       =   10
      TabIndex        =   14
      Text            =   "0"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox DescTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   10935
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   266
      Text            =   "0"
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label SaveCmd 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Left            =   11880
      TabIndex        =   286
      Top             =   840
      Width           =   555
   End
   Begin VB.Label LoadCmd 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Left            =   11880
      TabIndex        =   285
      Top             =   600
      Width           =   540
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Known Skills:"
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
      Index           =   18
      Left            =   7920
      TabIndex        =   284
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipted Weapon Slot:"
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
      Index           =   16
      Left            =   9585
      TabIndex        =   283
      Top             =   4080
      Width           =   2400
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipted Armor Slot:"
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
      Index           =   15
      Left            =   9840
      TabIndex        =   282
      Top             =   3840
      Width           =   2145
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Quest:"
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
      Index           =   14
      Left            =   9960
      TabIndex        =   281
      Top             =   1440
      Width           =   1470
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Completed Quests:"
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
      Index           =   13
      Left            =   240
      TabIndex        =   280
      Top             =   1440
      Width           =   1980
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Head Heading:"
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
      Left            =   6720
      TabIndex        =   279
      Top             =   960
      Width           =   1590
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Heading:"
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
      Index           =   11
      Left            =   5160
      TabIndex        =   278
      Top             =   960
      Width           =   960
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon:"
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
      Left            =   3720
      TabIndex        =   277
      Top             =   960
      Width           =   945
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
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
      Left            =   2520
      TabIndex        =   276
      Top             =   960
      Width           =   615
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Head:"
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
      Left            =   1320
      TabIndex        =   275
      Top             =   960
      Width           =   645
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hair:"
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
      Left            =   240
      TabIndex        =   274
      Top             =   960
      Width           =   510
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pos (Map-X-Y):"
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
      Left            =   7440
      TabIndex        =   273
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Gold:"
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
      Index           =   5
      Left            =   5640
      TabIndex        =   272
      Top             =   720
      Width           =   570
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pass:"
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
      Index           =   4
      Left            =   3840
      TabIndex        =   271
      Top             =   720
      Width           =   600
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory (ItemID/Amount/IsEquipted):"
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
      Index           =   3
      Left            =   240
      TabIndex        =   270
      Top             =   4080
      Width           =   3915
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   269
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Index           =   1
      Left            =   240
      TabIndex        =   268
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   267
      Top             =   720
      Width           =   690
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
                        Unload Me
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub AmountTxt_Change(Index As Integer)

    If Index > MAX_INVENTORY_SLOTS Then Exit Sub
    UserChar.Object(Index).Amount = AmountTxt(Index).Text

End Sub

Private Sub AmountTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the tooltip

    AmountTxt(Index).ToolTipText = "Index: " & Index

End Sub

Private Sub BodyTxt_Change()

    UserChar.Char.Body = BodyTxt.Text

End Sub

Private Sub CompletedQuestsTxt_Change()

    UserChar.CompletedQuests = CompletedQuestsTxt.Text

End Sub

Private Sub DescTxt_Change()

    UserChar.Desc = DescTxt.Text

End Sub

Private Sub EquiptedTxt_Change(Index As Integer)

    If Index > MAX_INVENTORY_SLOTS Then Exit Sub
    UserChar.Object(Index).Equipped = EquiptedTxt(Index).Text

End Sub

Private Sub EquiptedTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the tooltip

    EquiptedTxt(Index).ToolTipText = "Index: " & Index

End Sub

Private Sub GoldTxt_Change()

    UserChar.Gold = GoldTxt.Text

End Sub

Private Sub HairTxt_Change()

    UserChar.Char.Hair = HairTxt.Text

End Sub

Private Sub HeadHeadingTxt_Change()

    UserChar.Char.HeadHeading = HeadHeadingTxt.Text

End Sub

Private Sub HeadingTxt_Change()

    UserChar.Char.Heading = HeadingTxt.Text

End Sub

Private Sub HeadTxt_Change()

    UserChar.Char.Head = HeadTxt.Text

End Sub

Private Sub InventoryTxt_Change(Index As Integer)

    If Index > MAX_INVENTORY_SLOTS Then Exit Sub
    UserChar.Object(Index).ObjIndex = InventoryTxt(Index).Text

End Sub

Private Sub InventoryTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the tooltip

    InventoryTxt(Index).ToolTipText = "Index: " & Index

End Sub

Private Sub KnownSkillTxt_Change(Index As Integer)

    If Index > NumSkills Then Exit Sub
    UserChar.KnownSkills(Index) = KnownSkillTxt(Index).Text

End Sub

Private Sub KnownSkillTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the tooltip

    KnownSkillTxt(Index).ToolTipText = "Index: " & Index
    KnownSkillTxt(Index).ShowWhatsThis

End Sub

Private Sub LoadCmd_Click()

'Prepare common dialog to show existing .chr files to load

    With CommDlg
        .Filter = "Character Files|*.chr"
        .DialogTitle = "Load character file"
        .FileName = ""
        .Flags = cdlOFNFileMustExist
        .ShowOpen
    End With

    'Get the file path
    FilePath = CommDlg.FileName

    'Check for valid path
    If FilePath = "" Then Exit Sub
    If Right$(FilePath, 4) <> ".chr" Then Exit Sub

    'Open the character file
    LoadUser FilePath

    'Fill in all the information
    FillInInformation

End Sub

Private Sub MapTxt_Change()

    UserChar.Pos.Map = MapTxt.Text

End Sub

Private Sub PassTxt_Change()

    UserChar.Password = PassTxt.Text

End Sub

Private Sub QuestTxt_Change()

    UserChar.Quest = QuestTxt.Text

End Sub

Private Sub SaveCmd_Click()

'Save the changes

    If MsgBox("Are you sure you wish to save changes to the character file?" & vbCrLf & "All changes are final and irreverseable.", vbYesNo) = vbNo Then Exit Sub
    SaveUser FilePath

End Sub

Private Sub StatTxt_Change(Index As Integer)

    If Index > NumStats Then Exit Sub
    UserChar.Stats.BaseStat(Index) = StatTxt(Index).Text

End Sub

Private Sub StatTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the tooltip

    StatTxt(Index).ToolTipText = "Index: " & Index

End Sub

Private Sub WeaponTxt_Change()

    UserChar.Char.Weapon = WeaponTxt.Text

End Sub

Private Sub XTxt_Change()

    UserChar.Pos.X = XTxt.Text

End Sub

Private Sub YTxt_Change()

    UserChar.Pos.Y = YTxt.Text

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:44)  Decl: 1  Code: 216  Total: 217 Lines
':) CommentOnly: 14 (6.5%)  Commented: 1 (0.5%)  Empty: 87 (40.1%)  Max Logic Depth: 2
