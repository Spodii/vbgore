VERSION 5.00
Begin VB.Form frmBlock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Blocks"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBlock.frx":0000
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox BlockAllChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   510
      TabIndex        =   4
      ToolTipText     =   "Block/Unblock all directions"
      Top             =   840
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   270
      TabIndex        =   3
      ToolTipText     =   "Block movement going West"
      Top             =   840
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   510
      TabIndex        =   2
      ToolTipText     =   "Block movement going South"
      Top             =   1080
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   750
      TabIndex        =   1
      ToolTipText     =   "Block movement going East"
      Top             =   840
      Width           =   200
   End
   Begin VB.CheckBox BlockChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   0
      ToolTipText     =   "Block movement going North"
      Top             =   600
      Width           =   200
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
