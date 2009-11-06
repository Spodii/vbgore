VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Optimization Report"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   334
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MapEditor.cButton FixBtn 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Fixes the selected entry only"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Fix Selected"
   End
   Begin MapEditor.cForm cForm 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      MaximizeBtn     =   0   'False
      MinimizeBtn     =   0   'False
      CaptionTop      =   0
      AllowResizing   =   0   'False
   End
   Begin VB.ListBox OptList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2970
      ItemData        =   "frmReport.frx":0000
      Left            =   120
      List            =   "frmReport.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin MapEditor.cButton SimBtn 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Fixes all problems similar to the selected entry"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Fix Similar"
   End
   Begin MapEditor.cButton AllBtn 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Fix all the problems in the list"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Fix All"
   End
   Begin MapEditor.cButton DeleteBtn 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      ToolTipText     =   "Delete the selected problem from the list - this will NOT fix or remove the problem, just hide it from the list"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Delete"
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UpdateReport()
Dim i As Long

    'Update the report list
    OptList.Clear
    If UBound(MapOpt) = 0 Then Exit Sub
    For i = 1 To UBound(MapOpt)
        Select Case MapOpt(i).Type
            Case None
                OptList.AddItem "Nothing (Fixed/Deleted)"
            Case ObjOnBlocked
                OptList.AddItem "Obj On Blocked - X: " & MapOpt(i).tX & " Y: " & MapOpt(i).tY
            Case NPCOnBlocked
                OptList.AddItem "NPC On Blocked - X: " & MapOpt(i).tX & " Y: " & MapOpt(i).tY
            Case DuplicateGrhLayers
                OptList.AddItem "Duplicate Grhs - X: " & MapOpt(i).tX & " Y: " & MapOpt(i).tY & " L1: " & MapOpt(i).Layer & " L2: " & MapOpt(i).Layer2
            Case EmptyLight
                OptList.AddItem "   Empty Light - X: " & MapOpt(i).tX & " Y: " & MapOpt(i).tY & " L: " & MapOpt(i).Layer
        End Select
    Next i
    
End Sub

Private Sub AllBtn_Click()
Dim i As Long

    'Click this button and all your problems shall vanish!
    For i = 1 To UBound(MapOpt)
        FixProblem i
    Next i
    UpdateReport

End Sub

Private Sub DeleteBtn_Click()
    
    MapOpt(OptList.ListIndex + 1).Type = None
    
    UpdateReport
    
End Sub

Private Sub FixBtn_Click()

    'Fix the selected problem
    FixProblem OptList.ListIndex + 1
    
    'Show results
    UpdateReport

End Sub

Private Sub Form_Load()

    cForm.LoadSkin Me
    Skin_Set Me
    
    'Show report
    UpdateReport
    
End Sub

Private Sub FixProblem(ByVal Index As Long)
Dim RetVal As VbMsgBoxResult
Dim i As Long

    On Error GoTo ErrOut
    
    'Check for a valid range
    If LBound(MapOpt) > Index Then Exit Sub
    If UBound(MapOpt) < Index Then Exit Sub

    'Auto-fix the problem the best we can
    Select Case MapOpt(Index).Type
    
        Case ObjOnBlocked
            'RetVal = MsgBox("Problem: Object placed on a blocked tile" & vbCrLf & "Yes: Remove object" & vbCrLf & "No: Remove block (only works if not in border)" & vbCrLf & "Cancel: Do nothing", vbYesNoCancel)
            'Remove the object
            'If RetVal = vbYes Then
                For i = 1 To LastObj
                    If OBJList(i).Pos.X = MapOpt(Index).tX Then
                        If OBJList(i).Pos.Y = MapOpt(Index).tY Then Engine_OBJ_Erase i
                    End If
                Next i
                MapData(MapOpt(Index).tX, MapOpt(Index).tY).ObjInfo.ObjIndex = 0
                MapData(MapOpt(Index).tX, MapOpt(Index).tY).ObjInfo.Amount = 0
                MapOpt(Index).Type = None   'Problem fixed :)
            'Remove the block
            'ElseIf RetVal = vbNo Then
            '    MapData(MapOpt(Index).tX, MapOpt(Index).tY).Blocked = 0
            '    MapOpt(Index).Type = None   'Problem fixed :)
            'End If
            
        Case NPCOnBlocked
            'RetVal = MsgBox("Problem: NPC placed on a blocked tile" & vbCrLf & "Yes: Remove NPC" & vbCrLf & "No: Remove block (only works if not in border)" & vbCrLf & "Cancel: Do nothing", vbYesNoCancel)
            'Remove the NPC
            'If RetVal = vbYes Then
                Engine_Char_Erase MapData(MapOpt(Index).tX, MapOpt(Index).tY).NPCIndex
                MapOpt(Index).Type = None   'Problem fixed :)
            'Remove the block
            'ElseIf RetVal = vbNo Then
            '    MapData(MapOpt(Index).tX, MapOpt(Index).tY).Blocked = 0
            '    MapOpt(Index).Type = None   'Problem fixed :)
            'End If
            
        Case EmptyLight
            'Remove the lights
            For i = (MapOpt(Index).Layer - 1) * 4 + 1 To (MapOpt(Index).Layer - 1) * 4 + 4
                MapData(MapOpt(Index).tX, MapOpt(Index).tY).Light(i) = -1
            Next i
            MapOpt(Index).Type = None   'Problem fixed :)
            
        Case DuplicateGrhLayers
            'Remove the lowest layer
            MapData(MapOpt(Index).tX, MapOpt(Index).tY).Graphic(MapOpt(Index).Layer).GrhIndex = 0
            MapOpt(Index).Type = None   'Problem fixed :)
            
    End Select

    Exit Sub
    
ErrOut:

    'Problem :(
    MsgBox "Error auto-fixing index " & Index, vbOKOnly
    
End Sub

Private Sub SimBtn_Click()
Dim ProblemType As MapOptType
Dim i As Long
    
    On Error GoTo ErrOut

    'Confirm
    If MsgBox("Fix all problems similar to the selected problem?") = vbNo Then Exit Sub
    
    'Fix problems similar to the selected
    ProblemType = MapOpt(OptList.ListIndex + 1).Type
    For i = 1 To UBound(MapOpt)
        If MapOpt(i).Type = ProblemType Then
            FixProblem i
        End If
    Next i
    UpdateReport
    
    Exit Sub
    
ErrOut:

    MsgBox "Error fixing all similar types to " & OptList.ListIndex + 1 & "! Stopped at " & i & ".", vbOKOnly

End Sub
