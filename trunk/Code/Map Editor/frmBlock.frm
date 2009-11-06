VERSION 5.00
Begin VB.Form frmBlock 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Blocks"
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBlock.frx":0000
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox SetAttackChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set attack"
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
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Set the attacking block value on click"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox SetWalkChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set walk"
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
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Set the walking block values on click"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox BlockAttackChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "Can not attack through the tile"
      Top             =   2370
      Width           =   200
   End
   Begin VB.CheckBox BlockAllChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   630
      TabIndex        =   4
      ToolTipText     =   "Block/Unblock all directions"
      Top             =   1680
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   390
      TabIndex        =   3
      ToolTipText     =   "Block movement going West"
      Top             =   1680
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   630
      TabIndex        =   2
      ToolTipText     =   "Block movement going South"
      Top             =   1920
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   870
      TabIndex        =   1
      ToolTipText     =   "Block movement going East"
      Top             =   1680
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   630
      TabIndex        =   0
      ToolTipText     =   "Block movement going North"
      Top             =   1440
      Width           =   200
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No Attack"
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
      TabIndex        =   6
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No Walk"
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
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   750
   End
End
Attribute VB_Name = "frmBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BlockAllChk_Click()
Dim i As Byte

    'Change all blocks
    For i = 1 To 4
        BlockChk(i).Value = BlockAllChk.Value
    Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Data2Path & "MapEditor.ini", "BLOCK", "X", Me.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "BLOCK", "Y", Me.Top
    HideFrmBlock

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

Private Sub SetAttackChk_Click()

    BlockAttackChk.Enabled = (SetAttackChk.Value = 1)

End Sub

Private Sub SetWalkChk_Click()
Dim i As Long

    BlockAllChk.Enabled = (SetWalkChk.Value = 1)
    For i = 1 To 4
        BlockChk(i).Enabled = (SetWalkChk.Value = 1)
    Next i

End Sub
