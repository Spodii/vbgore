VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Grid-Created Grh List"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmView.frx":17D2A
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

    'Resize the GrhTxt
    GrhTxt.Left = 25
    GrhTxt.Top = 25
    GrhTxt.Width = Me.Width - 200
    GrhTxt.Height = Me.Height - 450

End Sub
