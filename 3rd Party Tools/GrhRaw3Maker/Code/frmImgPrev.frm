VERSION 5.00
Begin VB.Form frmImgPrev 
   Caption         =   "Image Viewer"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   Icon            =   "frmImgPrev.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScroll 
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.PictureBox picImage 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.VScrollBar VScroll 
      Height          =   3615
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   5295
   End
End
Attribute VB_Name = "frmImgPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub form_load()
    picImage.Picture = LoadPicture(frmMain.CommDlg.FileName)
    If picImage.Width > picScroll.Width Then
        HScroll.Max = picImage.Width
        HScroll.Visible = True
    End If
    If picImage.Height > picScroll.Height Then
        VScroll.Max = picImage.Height
        VScroll.Visible = True
    End If
End Sub

Private Sub HScroll_Scroll()
   picImage.Left = -HScroll.Value
End Sub

Private Sub VScroll_Change()
    VScroll_Scroll
End Sub

Private Sub VScroll_Scroll()
    picImage.Top = -VScroll.Value
End Sub
