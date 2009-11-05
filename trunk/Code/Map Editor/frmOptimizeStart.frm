VERSION 5.00
Begin VB.Form frmOptimizeStart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Optimize Map"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   178
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox DuplicateGrhChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Duplicate Grh Layers"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Check for the same graphic placed on multiple layers"
      Top             =   600
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox BlockedNPCChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "NPCs On Blocked Tiles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Check for NPCs placed on tiles that are blocked"
      Top             =   840
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox EmptyLightsChk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Lighted Empty Layers"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Check for lights placed on layers that do not have a graphic"
      Top             =   360
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Label OptBtn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Begin Optimization Check"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   2190
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Optimization Checks:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmOptimizeStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddToOpt(ByVal OptType As MapOptType, ByVal tX As Byte, ByVal tY As Byte, Optional ByVal Layer As Byte = 0, Optional ByVal Layer2 As Byte = 0)
Dim Index As Long

    'Add an item to the optimization list
    Index = UBound(MapOpt) + 1
    ReDim Preserve MapOpt(0 To Index)
    MapOpt(Index).tX = tX
    MapOpt(Index).tY = tY
    MapOpt(Index).Layer = Layer
    MapOpt(Index).Layer2 = Layer2
    MapOpt(Index).Type = OptType

End Sub

Private Sub BlockedNPCChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Check for NPCs placed on tiles that are blocked."

End Sub

Private Sub DuplicateGrhChk_Click()

    SetInfo "Check for the same graphic placed on multiple layers."

End Sub

Private Sub EmptyLightsChk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Check for lights placed on layers that do not have a grh."

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub OptBtn_Click()
Dim X As Long
Dim Y As Long
Dim i As Long
Dim j As Long

    'Set up the list
    ReDim MapOpt(0)

    'Check for objects on blocked tiles
    'If BlockedObjChk.Value = 1 Then
    '    OptBtn.Caption = "Checking Obj tiles..."
    '    OptBtn.Refresh
    '    For X = 1 To MapInfo.Width
    '        For Y = 1 To MapInfo.Height
    '            If MapData(X, Y).ObjInfo.ObjIndex > 0 Then
    '                'Check if the obj is in the border
    '                If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    '                    AddToOpt ObjOnBlocked, X, Y
    '                'Check if on a blocked tile
    '                Else
    '                    If MapData(X, Y).Blocked Then AddToOpt ObjOnBlocked, X, Y
    '                End If
    '            End If
    '        Next Y
    '    Next X
    'End If
    
    'Check for NPCs on blocked tiles
    If BlockedNPCChk.Value = 1 Then
        OptBtn.Caption = "Checking NPC tiles..."
        OptBtn.Refresh
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
                If MapData(X, Y).NPCIndex > 0 Then
                    'Check if the NPC is in the border
                    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
                        AddToOpt NPCOnBlocked, X, Y
                    'Check if on a blocked tile
                    Else
                        If MapData(X, Y).Blocked Then AddToOpt NPCOnBlocked, X, Y
                    End If
                End If
            Next Y
        Next X
    End If
    
    'Check for empty lights
    If EmptyLightsChk.Value = 1 Then
        OptBtn.Caption = "Checking lights..."
        OptBtn.Refresh
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
                For i = 2 To 6  'Loop through layers
                    If MapData(X, Y).Graphic(i).GrhIndex = 0 Then   'Check for empty layer
                        'Check the 4 lights that correspond to the layer
                        If MapData(X, Y).Light((i - 1) * 4 + 1) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        ElseIf MapData(X, Y).Light((i - 1) * 4 + 2) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        ElseIf MapData(X, Y).Light((i - 1) * 4 + 3) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        ElseIf MapData(X, Y).Light((i - 1) * 4 + 4) <> -1 Then
                            AddToOpt EmptyLight, X, Y, i
                        End If
                    End If
                Next i
            Next Y
        Next X
    End If
    
    'Check for duplicate grh layers (same tile with the same grh on two or more layers)
    If EmptyLightsChk.Value = 1 Then
        OptBtn.Caption = "Checking grh layers..."
        OptBtn.Refresh
        For X = 1 To MapInfo.Width
            For Y = 1 To MapInfo.Height
                For i = 1 To 6  'Loop through base layers
                    If MapData(X, Y).Graphic(i).GrhIndex > 0 Then   'We dont care if the grh = 0
                        For j = i + 1 To 6  'Loop through comparison layers
                            If MapData(X, Y).Graphic(i).GrhIndex = MapData(X, Y).Graphic(j).GrhIndex Then
                                AddToOpt DuplicateGrhLayers, X, Y, i, j
                            End If
                        Next j
                    End If
                Next i
            Next Y
        Next X
    End If
    
    'Restore the label
    OptBtn.Caption = "Begin Optimization Check"
    OptBtn.Refresh
    
    'Show report
    Me.Visible = False
    frmReport.Visible = True
    frmReport.Show
    frmReport.SetFocus

End Sub

Private Sub OptBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Begins the optimization check routine."

End Sub
