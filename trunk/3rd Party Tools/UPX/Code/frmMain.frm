VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "UPX - EXE Compressor"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LogTxt 
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmMain.frx":0000
      Top             =   1080
      Width           =   5415
   End
   Begin VB.CommandButton RunCmd 
      Caption         =   "Run!"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox DirTxt 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Log:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Source Directory:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Log(ByVal LogText As String)

    LogTxt.Text = LogTxt.Text & vbNewLine & LogText

End Sub

Private Sub Form_Load()
Dim s() As String
Dim rs As String
Dim i As Long

    LogTxt.Text = vbNullString
    
    s = Split(App.Path & "\", "\")
    For i = 0 To UBound(s) - 3
        rs = rs & s(i) & "\"
    Next i
    DirTxt.Text = rs
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Me.Visible = False

End Sub

Private Sub LogTxt_Change()

    LogTxt.SelStart = Len(LogTxt.Text)

End Sub

Private Sub RunCmd_Click()
Dim s() As String
Dim Files() As String
Dim i As Long
Dim j As Long

    Files() = AllFilesInFolders(DirTxt.Text, True)
    Log UBound(Files) + 1 & " files found total"
    For i = 0 To UBound(Files)
        If Me.Visible = False Then Exit For
        If ValidateImageSuffix(Files(i)) Then
            s = Split(Files(i), "\")
            Log "Optimizing file " & s(UBound(s))
            DoEvents
            OptimizeEXE Files(i)
            j = j + 1
        End If
    Next i
    If Me.Visible = False Then
        Unload Me
        End
    End If
    Log "Optimizations complete!"
    Log "A total of " & j & " files were processed."

End Sub

Private Function ValidateImageSuffix(ByVal File As String) As Boolean

    File = LCase$(File)
    If Right$(File, 4) = ".exe" Then
        ValidateImageSuffix = True
    Else
        ValidateImageSuffix = False
    End If

End Function

Private Sub OptimizeEXE(ByVal File As String)

    CommandLine App.Path & "\upxconsole.exe --brute -qqq " & Chr$(34) & File & Chr$(34)

End Sub

Private Sub CommandLine(ByVal CommandLineString As String)
Dim Start As STARTUPINFO
Dim Proc As PROCESS_INFORMATION

    Start.dwFlags = &H1
    Start.wShowWindow = 0
    CreateProcessA 0&, CommandLineString, 0&, 0&, False, &H20&, 0&, 0&, Start, Proc
    Do While WaitForSingleObject(Proc.hProcess, 0) = 258
        DoEvents
        Sleep 1
    Loop

End Sub
