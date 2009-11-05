VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Free Number Finder"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4260
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label GrhValuesLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   600
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4050
      WordWrap        =   -1  'True
   End
   Begin VB.Label GrhFilesLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   600
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   4050
      WordWrap        =   -1  'True
   End
   Begin VB.Label CalculateLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recalculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free Grh Values (Grh1.raw):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Free Grh Files (XXX.png):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2610
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CheckNum As Long = 15
Private GrhValues(1 To CheckNum) As Long
Private GrhFiles(1 To CheckNum) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Private Sub CalculateLbl_Click()
Dim Ret As String
Dim i As Long
Dim c As Long

    'Calculate the free grh values
    c = 1
    i = 0
    Do While c <= CheckNum
        i = i + 1
        If Var_Get(Data2Path & "Grh1.raw", "A", "Grh" & i) = "" Then
            GrhValues(c) = i
            c = c + 1
        End If
    Loop
    
    'Calculate the free grh files
    c = 1
    i = 0
    Do While c <= CheckNum
        i = i + 1
        If Server_FileExist(GrhPath & i & ".PNG", vbNormal) = False Then
            GrhFiles(c) = i
            c = c + 1
        End If
    Loop
    
    'Display
    Ret = ""
    For i = 1 To CheckNum
        Ret = Ret & GrhFiles(i)
        If i <> CheckNum Then Ret = Ret & ", "
    Next i
    GrhFilesLbl.Caption = Ret
    
    Ret = ""
    For i = 1 To CheckNum
        Ret = Ret & GrhValues(i)
        If i <> CheckNum Then Ret = Ret & ", "
    Next i
    GrhValuesLbl.Caption = Ret
    

End Sub

Private Sub Form_Load()

    InitFilePaths
    CalculateLbl_Click

End Sub

Private Function Server_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
On Error GoTo ErrOut

    If Dir$(File, FileType) <> "" Then Server_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Server_FileExist = False

End Function

Private Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = vbNullString

    sSpaces = Space$(1000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    Var_Get = RTrim$(sSpaces)
    Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Private Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub