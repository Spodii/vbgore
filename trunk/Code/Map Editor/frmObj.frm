VERSION 5.00
Begin VB.Form frmObj 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Objects"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmObj.frx":0000
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   181
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      ToolTipText     =   "The amount of objects that will be placed down"
      Top             =   3375
      Width           =   975
   End
   Begin VB.OptionButton SetOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set OBJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "If the selected object will be placed on the map when a tile is clicked"
      Top             =   3600
      Width           =   975
   End
   Begin VB.OptionButton EraseOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Erase OBJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "If an existing object is erased from the map when a tile is clicked"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ListBox OBJList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "The object to be placed down - these objects do NOT respawn. Once they are picked up, they are gone forever"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      Left            =   840
      TabIndex        =   4
      Top             =   3375
      Width           =   705
   End
End
Attribute VB_Name = "frmObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AmountTxt_Change()

    If IsNumeric(AmountTxt.Text) = False Then AmountTxt.Text = "0"
    If Val(AmountTxt.Text) < 0 Then AmountTxt.Text = "0"

End Sub

Private Sub EraseOpt_Click()

    SetOpt.Value = False
    EraseOpt.Value = True

End Sub

Private Sub Form_Load()
Dim NumOBJs As Integer
Dim OBJ As Long
Dim OBJs() As ObjData
Dim FileNum As Byte
Dim LoopC As Integer

    'Set the default option to Set
    SetOpt.Value = True
    EraseOpt.Value = False
    
    'Set the NPCs array
    FileNum = FreeFile
    DB_RS.Open "SELECT id FROM objects ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    NumOBJs = DB_RS(0)
    DB_RS.Close
    ReDim OBJs(1 To NumOBJs)
    OBJList.Clear

    'Load all the names
    DB_RS.Open "SELECT name FROM objects", DB_Conn, adOpenStatic, adLockOptimistic
    Do While DB_RS.EOF = False
        OBJList.AddItem Trim$(DB_RS!Name)
        DB_RS.MoveNext
    Loop
    DB_RS.Close
    
    'Select the first slot
    OBJList.ListIndex = 0

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Engine_Var_Write Data2Path & "MapEditor.ini", "OBJ", "X", Me.Left
    Engine_Var_Write Data2Path & "MapEditor.ini", "OBJ", "Y", Me.Top
    HideFrmObj

End Sub

Private Sub SetOpt_Click()

    SetOpt.Value = True
    EraseOpt.Value = False

End Sub
