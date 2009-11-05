VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Exits / Warps"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmExit.frx":0000
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   119
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton SetOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Set an exit when clicking on the map"
      Top             =   1200
      Width           =   615
   End
   Begin VB.OptionButton EraseOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Erase"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "Erase an already placed exit when clicking on the map"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox YTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Text            =   "20"
      ToolTipText     =   "Y co-ordinate which the user will warp to when stepping on the tile"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox XTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Text            =   "20"
      ToolTipText     =   "X co-ordinate which the user will warp to when stepping on the tile"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox MapTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Text            =   "1"
      ToolTipText     =   "Map which the user will warp to when stepping on the tile"
      Top             =   720
      Width           =   855
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
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   960
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
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   195
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Map:"
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
      TabIndex        =   5
      Top             =   720
      Width           =   435
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetOpt_Click()

    SetOpt.Value = True
    EraseOpt.Value = False

End Sub

Private Sub EraseOpt_Click()

    SetOpt.Value = False
    EraseOpt.Value = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Data2Path & "MapEditor.ini", "EXIT", "X", Me.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "EXIT", "Y", Me.Top
    HideFrmExit

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
