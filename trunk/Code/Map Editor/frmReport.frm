VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Optimization Report"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmReport.frx":0000
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ItemData        =   "frmReport.frx":40E6E
      Left            =   120
      List            =   "frmReport.frx":40E70
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label SimLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Similar"
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
      Left            =   1800
      TabIndex        =   4
      ToolTipText     =   "Fix all problems similar to the selected effect"
      Top             =   3720
      Width           =   870
   End
   Begin VB.Label AllLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fix All"
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
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "Fix all the problems in the list"
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label DeleteLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
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
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Delete the problem from the list - this will NOT fix or remove the problem, just hide it from the list"
      Top             =   3720
      Width           =   570
   End
   Begin VB.Label FixLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fix Selected"
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
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Fix only the selected problem"
      Top             =   3720
      Width           =   1065
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

Private Sub AllLbl_Click()
Dim i As Long

    'Click this button and all your problems shall vanish!
    For i = 1 To UBound(MapOpt)
        FixProblem i
    Next i
    UpdateReport

End Sub

Private Sub FixLbl_Click()

    'Fix the selected problem
    FixProblem OptList.ListIndex + 1
    
    'Show results
    UpdateReport

End Sub

Private Sub Form_Load()

    'Show report
    UpdateReport

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

Private Sub SimLbl_Click()
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
