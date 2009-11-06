VERSION 5.00
Begin VB.Form frmBlock 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tile Blocks"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox SetAttackChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Set the attacking block value on click"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox SetWalkChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Set the walking block values on click"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox BlockAttackChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "Can not attack through the tile"
      Top             =   1890
      Width           =   200
   End
   Begin VB.CheckBox BlockAllChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   630
      TabIndex        =   4
      ToolTipText     =   "Block/Unblock all directions"
      Top             =   1200
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   390
      TabIndex        =   3
      ToolTipText     =   "Block movement going West"
      Top             =   1200
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   630
      TabIndex        =   2
      ToolTipText     =   "Block movement going South"
      Top             =   1440
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   870
      TabIndex        =   1
      ToolTipText     =   "Block movement going East"
      Top             =   1200
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   630
      TabIndex        =   0
      ToolTipText     =   "Block movement going North"
      Top             =   960
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1920
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   720
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

Private Sub BlockAllChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Disable / enable all directions at once."

End Sub

Private Sub BlockAttackChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Enable / disable attacking with ranged attacks over this tile."

End Sub

Private Sub BlockChk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case 1: SetInfo "Disable walking north onto this tile."
        Case 2: SetInfo "Disable walking east onto this tile."
        Case 3: SetInfo "Disable walking south onto this tile."
        Case 4: SetInfo "Disable walking west onto this tile."
    End Select
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Enable / disable attacking with ranged attacks over this tile."

End Sub

Private Sub SetAttackChk_Click()

    BlockAttackChk.Enabled = (SetAttackChk.Value = 1)

End Sub

Private Sub SetAttackChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Enables / disables modifying if tiles can be attacked over with ranged attacks."

End Sub

Private Sub SetWalkChk_Click()
Dim i As Long

    BlockAllChk.Enabled = (SetWalkChk.Value = 1)
    For i = 1 To 4
        BlockChk(i).Enabled = (SetWalkChk.Value = 1)
    Next i

End Sub

Private Sub SetWalkChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Enables / disables modifying if tiles can be walked on."

End Sub
