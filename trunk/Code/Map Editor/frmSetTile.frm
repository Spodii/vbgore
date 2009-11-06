VERSION 5.00
Begin VB.Form frmSetTile 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Set Tile"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   192
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ShadowTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "0"
      ToolTipText     =   "1 = Sets Shadow, 0 = Removes Shadow"
      Top             =   1950
      Width           =   255
   End
   Begin VB.CheckBox LightChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Light"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Set light layer 1"
      Top             =   600
      Width           =   720
   End
   Begin VB.CheckBox ShadowChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shadow"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Set layer 4"
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox LayerChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Grh"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Set graphic layer 1"
      Top             =   600
      Width           =   600
   End
   Begin VB.PictureBox LightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   9
      ToolTipText     =   "Preview of the light for the layer"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Graphic index of the layer"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   6
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Right corner"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Right corner"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   360
      MaxLength       =   11
      TabIndex        =   3
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Top-Left corner"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox LightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   360
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "-1"
      ToolTipText     =   "Light placed in the Bottom-Left corner"
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image LayerPic 
      Height          =   480
      Index           =   6
      Left            =   2400
      Top             =   0
      Width           =   480
   End
   Begin VB.Image LayerPic 
      Height          =   480
      Index           =   5
      Left            =   1920
      Top             =   0
      Width           =   480
   End
   Begin VB.Image LayerPic 
      Height          =   480
      Index           =   4
      Left            =   1440
      Top             =   0
      Width           =   480
   End
   Begin VB.Image LayerPic 
      Height          =   480
      Index           =   3
      Left            =   960
      Top             =   0
      Width           =   480
   End
   Begin VB.Image LayerPic 
      Height          =   480
      Index           =   2
      Left            =   480
      Top             =   0
      Width           =   480
   End
   Begin VB.Image LayerPic 
      Height          =   480
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shadow:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   1950
      Width           =   750
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
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1590
      Width           =   375
   End
End
Attribute VB_Name = "frmSetTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub GrhTxt_Change()
Dim i As Integer
On Error GoTo ErrOut

    i = Val(GrhTxt.Text)
    
    'Check for valid range
    If Val(GrhTxt.Text) < 0 Then GrhTxt.Text = "0"
    If Val(GrhTxt.Text) > UBound(GrhData) Then Exit Sub

    Engine_Init_Grh PreviewGrh, Val(GrhTxt.Text)

    DrawPreview
    
    Exit Sub
    
ErrOut:

End Sub

Private Sub GrhTxt_KeyPress(KeyAscii As Integer)

    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    
End Sub

Private Sub GrhTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Grh that will be placed on the layer."

End Sub

Private Sub LayerChk_Click()

    DrawPreview

End Sub

Private Sub LayerChk_KeyPress(KeyAscii As Integer)

    SetInfo "Enables / disables graphic placing and modifying."

End Sub

Private Sub LayerPic_Click(Index As Integer)

    SetLayer Index

End Sub

Private Sub LayerPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Click to select placing graphics on layer " & Index & "."

End Sub

Private Sub LightChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Enables / disables tile light modifying."

End Sub

Private Sub LightPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Preview of what the light will look like for the layer."

End Sub

Private Sub LightTxt_Change(Index As Integer)
Dim i As Long
Dim b As Long

    On Error GoTo ErrOut
    
    i = Val(LightTxt(Index).Text)
    
    Exit Sub

ErrOut:
    
    'Set b as 1 by default (light value is positive)
    b = 1
    
    'If the light value is negative, set b to -1
    If Len(LightTxt(Index).Text) > 1 Then
        If Left$(LightTxt(Index).Text, 1) = "-" Then
            b = -1
        End If
    End If
    
    'Set the value to negative or positive accordingly, then move 1 value
    ' closer to 0 (for negative, add, for positive, subtract) to keep in range
    LightTxt(Index).Text = b * (2 ^ 31) - b

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

Private Sub LightTxt_LostFocus(Index As Integer)
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
    
    If Not Engine_ValidateDevice Then Exit Sub
    
    'Draw the light preview
    For i = 1 To 6
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(255, 0, 0, 0), 1#, 0
        D3DDevice.BeginScene
            Engine_Render_Rectangle 0, 0, 25, 25, 1, 1, 1, 1, 1, 1, 0, 0, Val(LightTxt(1).Text), Val(LightTxt(2).Text), Val(LightTxt(3).Text), Val(LightTxt(4).Text)
        D3DDevice.EndScene
        D3DDevice.Present TempRect, TempRect, frmSetTile.LightPic.hWnd, ByVal 0
    Next i

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
    
    SetInfo "Sets the light value in the tile's " & s & " corner."

End Sub

Private Sub ShadowChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Enables / disables tile graphic shadow value modifying."

End Sub

Private Sub ShadowTxt_Change()
Dim i As Long
On Error GoTo ErrOut

    i = Val(ShadowTxt.Text)
    If i > 0 Then ShadowTxt.Text = 1
    If i < 0 Then ShadowTxt.Text = 0
    
    Exit Sub
    
ErrOut:

End Sub

Private Sub ShadowTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub ShadowTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Sets the tile to cast a shadow (1 = enables shadow, 0 = disables shadow)."

End Sub
