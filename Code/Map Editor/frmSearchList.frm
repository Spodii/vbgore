VERSION 5.00
Begin VB.Form frmSearchList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search: """""
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox SearchLst 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmSearchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Me.Width = Val(Var_Get(Data2Path & "MapEditor.ini", Me.Name, "W"))
    Me.Height = Val(Var_Get(Data2Path & "MapEditor.ini", Me.Name, "H"))
    Me.Left = Val(Var_Get(Data2Path & "MapEditor.ini", Me.Name, "X"))
    Me.Top = Val(Var_Get(Data2Path & "MapEditor.ini", Me.Name, "Y"))

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Var_Write Data2Path & "MapEditor.ini", Me.Name, "W", Me.Width
    Var_Write Data2Path & "MapEditor.ini", Me.Name, "H", Me.Height
    Var_Write Data2Path & "MapEditor.ini", Me.Name, "X", Me.Left
    Var_Write Data2Path & "MapEditor.ini", Me.Name, "Y", Me.Top
    Form_Resize
    
End Sub

Private Sub Form_Resize()

    SearchLst.Width = Me.ScaleWidth
    SearchLst.Height = Me.ScaleHeight + 5

End Sub

Private Sub SearchLst_Click()

    If LoadTextureToForm(frmSearchTexture, DescResults(SearchLst.ListIndex + 1)) = 0 Then Exit Sub
    SearchTextureFileNum = DescResults(SearchLst.ListIndex + 1)

End Sub

Private Sub SearchLst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    SetInfo vbNullString
    
End Sub
