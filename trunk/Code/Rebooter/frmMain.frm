VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "vbGORE Rebooter"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer RestartTmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton RunCmd 
      Caption         =   "Start Rebooter"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox PathTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label BrowseLbl 
      AutoSize        =   -1  'True
      Caption         =   "..."
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
      Left            =   2700
      TabIndex        =   3
      Top             =   450
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Application path:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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

Private Proc As PROCESS_INFORMATION

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub BrowseLbl_Click()
    
    'Set up the browse dialog
    CD.Filter = "Executeables|*.exe"
    CD.DialogTitle = "Select server"
    CD.FileName = "GameServer.exe"
    CD.InitDir = App.Path
    CD.Flags = cdlOFNFileMustExist
    
    'Open the dialog
    CD.ShowOpen
    
    'When the dialog is closed, grab the path retrieved
    PathTxt.Text = Trim$(CD.FileName)
    
    'Move to the right side of the text
    PathTxt.SelStart = Len(PathTxt.Text)

End Sub

Private Sub Form_Load()

    'Set the default path
    PathTxt.Text = App.Path & "\GameServer.exe"
    
    'Move to the right side of the text
    PathTxt.SelStart = Len(PathTxt.Text)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Restore from tray
    If X = LeftDown Then
        If Me.WindowState = 1 Then
            Me.WindowState = 0
            Me.Show
            TrayDelete
        End If
    End If

End Sub

Private Sub Form_Resize()

    'Check if minimized
    If WindowState = vbMinimized Then
        
        'Put into the tray
        TrayAdd Me, Me.Caption, MouseMove
        Me.Hide
        
    End If

End Sub

Private Sub RestartTmr_Timer()
Dim Start As STARTUPINFO

    'Check the process status
    If WaitForSingleObject(Proc.hProcess, 0) <> 258 Then
        
        'Process has ended, restart it
        Start.cb = Len(Start)
        CreateProcessA 0&, PathTxt.Text, 0&, 0&, 1&, &H20&, 0&, 0&, Start, Proc
        
    End If

End Sub

Private Sub RunCmd_Click()

    If RunCmd.Caption = "Start Rebooter" Then

        'Check if the file exists
        If Engine_FileExist(PathTxt.Text, vbNormal) = False Then
            MsgBox "File does not exist!", vbOKOnly Or vbCritical
            Exit Sub
        End If
    
        'Enable the restart timer
        RestartTmr.Enabled = True
        
        'Change the caption
        RunCmd.Caption = "Stop Rebooter"
        
    Else
    
        'Turn off the timer
        RestartTmr.Enabled = False
        
        'Change the caption
        RunCmd.Caption = "Start Rebooter"
        
    End If
    
End Sub

Private Function Engine_FileExist(ByVal File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function
