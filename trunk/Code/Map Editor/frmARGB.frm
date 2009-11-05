VERSION 5.00
Begin VB.Form frmARGB 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ARGB <-> Long"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox BTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Text            =   "255"
      ToolTipText     =   "Blue value of the light (0 to 255)"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox GTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "255"
      ToolTipText     =   "Green value of the light (0 to 255)"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox RTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "255"
      ToolTipText     =   "Red value of the light (0 to 255)"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox ATxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "255"
      ToolTipText     =   "Alpha value of the light (0 to 255)"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox LongTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "-1"
      ToolTipText     =   "Long value of the ARGB light"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Left            =   75
      TabIndex        =   10
      Top             =   990
      Width           =   135
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
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
      Left            =   1275
      TabIndex        =   9
      Top             =   990
      Width           =   150
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
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
      Left            =   675
      TabIndex        =   8
      Top             =   990
      Width           =   150
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
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
      Left            =   1875
      TabIndex        =   7
      Top             =   990
      Width           =   135
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ARGB value:"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label MiscLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Captured LONG value:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmARGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ATxt_Change()

    ARGBtoLONG

End Sub

Private Sub ATxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Alpha color value (ex: 255 for opaque / solid, 0 for invisible, 128 for 50% transparency)."

End Sub

Private Sub BTxt_Change()

    ARGBtoLONG

End Sub

Private Sub BTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Blue color value."

End Sub

Private Sub Form_Load()

    Me.Top = Val(Var_Get(Data2Path & "MapEditor.ini", Me.Name, "Y"))
    Me.Left = Val(Var_Get(Data2Path & "MapEditor.ini", Me.Name, "X"))

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1
    Me.Visible = False

End Sub

Private Sub GTxt_Change()

    ARGBtoLONG

End Sub

Private Sub GTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Green color value."

End Sub

Private Sub LongTxt_Change()
Dim b(0 To 3) As Byte
Dim l As Long
    
    On Error GoTo ErrOut

    'Create the ARGB values
    l = CLng(LongTxt.Text)
    CopyMemory b(0), l, 4
    
    'Display the values
    BTxt.Text = b(0)
    GTxt.Text = b(1)
    RTxt.Text = b(2)
    ATxt.Text = b(3)
           
Exit Sub

ErrOut:

    MsgBox "Error converting Long value (" & l & ") to ARGB value (" & b(0) & "," & b(1) & "," & b(2) & "," & b(3) & ")"

End Sub

Private Sub ARGBtoLONG()

    LongTxt.Text = D3DColorARGB(Val(ATxt.Text), Val(RTxt.Text), Val(GTxt.Text), Val(BTxt.Text))

End Sub

Private Sub LongTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "The LONG value created by the entered ARGB value. This is the value that the engine uses."

End Sub

Private Sub RTxt_Change()
    
    ARGBtoLONG
    
End Sub

Private Sub RTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo "Red color value."

End Sub
