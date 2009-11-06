VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Grh Categorizer"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   347
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton NextBtn 
      Caption         =   "Next"
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.Timer RenderTmr 
      Interval        =   50
      Left            =   4320
      Top             =   600
   End
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
      Left            =   2160
      TabIndex        =   8
      Top             =   4560
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
      Left            =   2160
      TabIndex        =   7
      Top             =   4320
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
      Left            =   2160
      TabIndex        =   6
      Top             =   4080
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
      Left            =   2160
      TabIndex        =   5
      Top             =   3840
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
      TabIndex        =   4
      Top             =   4560
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
      TabIndex        =   3
      Top             =   4320
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
      TabIndex        =   2
      Top             =   4080
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
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.PictureBox PreviewPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Back 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invert Background"
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
      Left            =   3360
      TabIndex        =   11
      Top             =   45
      Width           =   1590
   End
   Begin VB.Label InfoLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Info:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   45
      Width           =   315
   End
   Begin VB.Label ValueLbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value: 0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4305
      TabIndex        =   9
      Top             =   3720
      Width           =   585
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The following can be used to get the categorization value:
'    j = "asdfLweeXdfasdf"
'    s = Mid$(j, InStr(1, j, "(") + 1, Len(j) - InStr(1, j, ")") - InStr(1, j, "(") + (InStr(1, j, ")") - InStr(1, j, "(") - 2))
'    MsgBox s

Private Const HighlightColor As Long = 65280
Private BckClr As Long

Private Sub ClearColors()
Dim i As Long

    For i = CatChk.lbound To CatChk.UBound
        CatChk(i).ForeColor = &H80000008
    Next i

End Sub

Private Function GetFlags() As Long
Dim i As Long

    'Update the value
    GetFlags = 0
    For i = CatChk.lbound To CatChk.UBound
        If CatChk(i).Value Then
            GetFlags = GetFlags Or (2 ^ i)
        End If
    Next i

End Function

Private Sub Back_Click()

    If BckClr = 0 Then BckClr = -1 Else BckClr = 0

End Sub

Private Sub UpdateValue()

    'Update the value
    ValueLbl.Caption = "Value: " & GetFlags

End Sub

Private Sub CatChk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ClearColors
    CatChk(Index).ForeColor = HighlightColor

End Sub

Private Sub CatChk_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim s As String

    'Force right-click to set the value
    If Button = vbRightButton Then CatChk(Index).Value = Abs(CatChk(Index).Value - 1)

    'Update the display value
    UpdateValue
    
    'Create the string
    s = Engine_Var_Get(Data2Path & "GrhRaw.txt", "A", "Grh" & CurrGrhNum)
    If InStr(1, s, "(") Then s = Left$(s, InStr(1, s, "(") - 2)

    'Save the value
    If GetFlags = 0 Then
        Engine_Var_Write Data2Path & "GrhRaw.txt", "A", "Grh" & CurrGrhNum, s
    Else
        Engine_Var_Write Data2Path & "GrhRaw.txt", "A", "Grh" & CurrGrhNum, s & "-(" & GetFlags & ")"
    End If
    
    'Move on if a right-click
    If Button = vbRightButton Then NextBtn_Click
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim c As Control
    
    For Each c In Me
        If TypeName(c) = "cButton" Then
            c.Refresh
            c.DrawState = 0
        End If
    Next c
    Set c = Nothing
    
    ClearColors

End Sub

Private Sub NextBtn_Click()
Dim i As Long

    'Get the next grh
    CurrGrhNum = GetNextUncategorizedGrh
    Engine_Init_Grh CurrGrh, CurrGrhNum

    'Clear the ticks
    For i = CatChk.lbound To CatChk.UBound
        CatChk(i).Value = 0
    Next i

End Sub

Private Sub NextBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ClearColors
    
End Sub

Private Sub PreviewPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ClearColors

End Sub

Private Sub RenderTmr_Timer()

    If CurrGrh.GrhIndex > 0 Then
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, BckClr, 1#, 0
        D3DDevice.BeginScene
            Engine_Render_Grh CurrGrh, 0, 0, 0, 1, True
        D3DDevice.EndScene
        D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    End If
    
End Sub
