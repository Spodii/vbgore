VERSION 5.00
Begin VB.Form frmExit 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Exits / Warps"
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Begin MapEditor.cForm cForm 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "Exits"
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.OptionButton SetOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Set an exit when clicking on the map"
      Top             =   840
      Width           =   615
   End
   Begin VB.OptionButton EraseOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Erase"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Erase an already placed exit when clicking on the map"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox YTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "20"
      ToolTipText     =   "Y co-ordinate which the user will warp to when stepping on the tile"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox XTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "20"
      ToolTipText     =   "X co-ordinate which the user will warp to when stepping on the tile"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox MapTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "1"
      ToolTipText     =   "Map which the user will warp to when stepping on the tile"
      Top             =   120
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
      Left            =   840
      TabIndex        =   7
      Top             =   480
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
      Left            =   120
      TabIndex        =   6
      Top             =   480
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    cForm.LoadSkin Me
    Skin_Set Me
    Me.Refresh

End Sub

Private Sub MapTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

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
    Var_Write Data2Path & "MapEditor.ini", "EXIT", "X", Me.Left
    Var_Write Data2Path & "MapEditor.ini", "EXIT", "Y", Me.Top
    HideFrmExit

End Sub

Private Sub XTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub YTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
