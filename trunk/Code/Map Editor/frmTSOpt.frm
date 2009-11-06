VERSION 5.00
Begin VB.Form frmTSOpt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Tile Select Options"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   245
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MapEditor.cButton SaveBtn 
      Height          =   375
      Left            =   1200
      TabIndex        =   16
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Save Changes"
   End
   Begin MapEditor.cForm cForm 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "Tile Selection Options"
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Misc (Hidden)"
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
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Misc (Displayed)"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Inside Objects"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Outside Objects"
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
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Buildings"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Vegetation"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Inside Tiles"
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
      TabIndex        =   8
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Outside Tiles"
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
      TabIndex        =   6
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox HeightTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Height of the previewed tile"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox WidthTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Width of the previewed tile"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox StartTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The starting number to view the tiles from"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Show Options:"
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
      TabIndex        =   7
      Top             =   960
      Width           =   1245
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Height:"
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
      TabIndex        =   5
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Width:"
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
      TabIndex        =   4
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Number:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmTSOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelLbl_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    HideFrmTSOpt
    ShowFrmTileSelect stBoxID

End Sub

Private Sub Form_Load()

    cForm.LoadSkin Me
    Skin_Set Me

End Sub

Private Sub HeightTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub SaveBtn_Click()

    tsTileHeight = Val(HeightTxt.Text)
    tsTileWidth = Val(WidthTxt.Text)
    tsStart = Val(StartTxt.Text)

    Var_Write Data2Path & "MapEditor.ini", "TSOPT", "W", WidthTxt.Text
    Var_Write Data2Path & "MapEditor.ini", "TSOPT", "H", HeightTxt.Text
    Var_Write Data2Path & "MapEditor.ini", "TSOPT", "S", StartTxt.Text
    
    tsWidth = CLng(frmTileSelect.ScaleWidth / tsTileWidth)  'Use clng to make sure we round down
    tsHeight = CLng(frmTileSelect.ScaleHeight / tsTileHeight)
    ReDim PreviewGrhList(tsWidth * tsHeight)    'Resize our array accordingly to fit all our Grhs
    Engine_SetTileSelectionArray
    
    HideFrmTSOpt
    ShowFrmTileSelect stBoxID

End Sub

Private Sub StartTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub WidthTxt_KeyPress(KeyAscii As Integer)
    If GetAsyncKeyState(vbKeyControl) = 0 Then
        If IsNumeric(Chr$(KeyAscii)) = False Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
