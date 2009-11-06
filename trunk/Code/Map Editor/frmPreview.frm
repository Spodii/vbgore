VERSION 5.00
Begin VB.Form frmPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Grh Preview"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   1740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   116
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    SetInfo "Preview of your currently selected grh."

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If IsUnloading = 0 Then Cancel = 1

End Sub
