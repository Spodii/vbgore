VERSION 5.00
Begin VB.Form frmNPCs 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " NPCs"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton EraseOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Erase NPC"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "If existing NPCs will be removed when the map is clicked"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton SetOpt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set NPC"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "If a NPC will be set when the map is clicked"
      Top             =   3000
      Width           =   975
   End
   Begin VB.ListBox NPCList 
      Appearance      =   0  'Flat
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

Private Sub EraseOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Click to toggle erasing NPCs."

End Sub

Private Sub Form_Load()
Dim NumNPCs As Integer
Dim NPC As Long
Dim NPCs() As NPC
Dim FileNum As Byte
Dim LoopC As Integer

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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub NPCList_Click()

    'Change to Set mode
    SetOpt_Click

End Sub

Private Sub NPCList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Select which NPC you want to place when using Set NPC."

End Sub

Private Sub SetOpt_Click()

    SetOpt.Value = True
    EraseOpt.Value = False

End Sub

Private Sub SetOpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Click to toggle placing NPCs. The NPC placed is the one selected in the NPCs list box."

End Sub
