VERSION 5.00
Begin VB.Form frmNPCs 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "NPCs"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   182
   ShowInTaskbar   =   0   'False
   Begin MapEditor.cForm cForm 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      Caption         =   "NPCs"
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.OptionButton EraseOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Erase NPC"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "If existing NPCs will be removed when the map is clicked"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton SetOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Set NPC"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "If a NPC will be set when the map is clicked"
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox NPCList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmNPCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EraseOpt_Click()

    SetOpt.Value = False
    EraseOpt.Value = True

End Sub

Private Sub Form_Load()
Dim NumNPCs As Integer
Dim NPC As Long
Dim NPCs() As NPC
Dim FileNum As Byte
Dim LoopC As Integer

    cForm.LoadSkin Me
    Skin_Set Me
    Me.Refresh
    
    'Set the default option to Set
    SetOpt.Value = True
    EraseOpt.Value = False
    
    'Get the number of NPCs (Sort by id, descending, only get 1 entry, only return id)
    DB_RS.Open "SELECT id FROM npcs ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    NumNPCs = DB_RS(0)
    DB_RS.Close
    
    'Clear the npc list
    NPCList.Clear
    
    'Check for a valid number of NPCs
    If NumNPCs <= 0 Then Exit Sub
    
    'Load the NPCs
    DB_RS.Open "SELECT name FROM npcs", DB_Conn, adOpenStatic, adLockOptimistic
    Do While DB_RS.EOF = False
        NPCList.AddItem Trim$(DB_RS!Name)
        DB_RS.MoveNext
    Loop
    DB_RS.Close
    
    'Select the first slot
    NPCList.ListIndex = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = 1
    Var_Write Data2Path & "MapEditor.ini", "NPCS", "X", Me.Left
    Var_Write Data2Path & "MapEditor.ini", "NPCS", "Y", Me.Top
    HideFrmNPCs

End Sub

Private Sub NPCList_Click()

    'Change to Set mode
    SetOpt_Click

End Sub

Private Sub SetOpt_Click()

    SetOpt.Value = True
    EraseOpt.Value = False

End Sub
