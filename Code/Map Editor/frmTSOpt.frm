VERSION 5.00
Begin VB.Form frmTSOpt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tile Select Options"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   253
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox CatChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox HeightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Height of the previewed tile"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox WidthTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "Width of the previewed tile"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox StartTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "0"
      ToolTipText     =   "The starting number to view the tiles from"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label SaveBtn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save Changes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   15
      Top             =   3480
      Width           =   1245
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1200
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   840
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
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
      ForeColor       =   &H80000008&
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    HideFrmTSOpt
    ShowFrmTileSelect stBoxID

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
    If tsTileWidth <= 0 Then tsTileWidth = 32
    If tsTileHeight <= 0 Then tsTileHeight = 32
    If tsStart <= 0 Then tsStart = 1
    If tsTileWidth > 1024 Then tsTileWidth = 1024
    If tsTileHeight > 1024 Then tsTileHeight = 1024

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
