VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Processor - Bulk file compressing and encrypting"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LogList 
      Height          =   3375
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   61
      TabStop         =   0   'False
      Text            =   "frmMain.frx":17D2A
      Top             =   5280
      Width           =   8535
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   9
      Left            =   1080
      TabIndex        =   47
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   42
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   37
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   6
      Left            =   1080
      TabIndex        =   32
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   27
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox SubTxt 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encryption / Compression"
      Height          =   5055
      Left            =   120
      TabIndex        =   52
      Top             =   120
      Width           =   8535
      Begin VB.TextBox SubTxt 
         Height          =   285
         Index           =   2
         Left            =   960
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox SubTxt 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox SubTxt 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox KillSourceChk 
         Caption         =   "Delete Source"
         Height          =   195
         Left            =   7080
         TabIndex        =   51
         ToolTipText     =   "Checking this box will delete a file after compressing or encrypting it, leaving only the compressed / encrypted copy"
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   9
         Left            =   7680
         TabIndex        =   50
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   46
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   9
         Left            =   2280
         TabIndex        =   48
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   9
         Left            =   4560
         TabIndex        =   49
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   8
         Left            =   7680
         TabIndex        =   45
         Top             =   3840
         Width           =   735
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   41
         Top             =   3840
         Width           =   735
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   8
         Left            =   2280
         TabIndex        =   43
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   8
         Left            =   4560
         TabIndex        =   44
         Top             =   3840
         Width           =   3015
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   7
         Left            =   7680
         TabIndex        =   40
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   3480
         Width           =   735
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   7
         Left            =   2280
         TabIndex        =   38
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   7
         Left            =   4560
         TabIndex        =   39
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   0
         Left            =   7680
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   1
         Left            =   7680
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   2
         Left            =   7680
         TabIndex        =   15
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   3
         Left            =   7680
         TabIndex        =   20
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   4
         Left            =   7680
         TabIndex        =   25
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   5
         Left            =   7680
         TabIndex        =   30
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox AExTxt 
         Height          =   285
         Index           =   6
         Left            =   7680
         TabIndex        =   35
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   6
         Left            =   4560
         TabIndex        =   34
         Top             =   3120
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   6
         Left            =   2280
         TabIndex        =   33
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   5
         Left            =   4560
         TabIndex        =   29
         Top             =   2760
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   5
         Left            =   2280
         TabIndex        =   28
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   26
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   4
         Left            =   4560
         TabIndex        =   24
         Top             =   2400
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   4
         Left            =   2280
         TabIndex        =   23
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   19
         Top             =   2040
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   3
         Left            =   2280
         TabIndex        =   18
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   2
         Left            =   4560
         TabIndex        =   14
         Top             =   1680
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   2
         Left            =   2280
         TabIndex        =   13
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   9
         Top             =   1320
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox KeyTxt 
         Height          =   285
         Index           =   0
         Left            =   4560
         TabIndex        =   4
         Top             =   960
         Width           =   3015
      End
      Begin VB.ComboBox AlgCmb 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox ExTxt 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton DecompBtn 
         Caption         =   "Decompress / Decrypt"
         Height          =   255
         Left            =   2040
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton CompBtn 
         Caption         =   "Compress / Encrypt"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox DirTxt 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Root folder path to use when searching for all the files with the defined extensions"
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Folder:"
         Height          =   195
         Index           =   5
         Left            =   960
         TabIndex        =   60
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Unaltered Extension:"
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Altered Extension:"
         Height          =   435
         Index           =   4
         Left            =   7680
         TabIndex        =   58
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Key:"
         Height          =   195
         Index           =   3
         Left            =   4560
         TabIndex        =   57
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Algorithm:"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   56
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folder path:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   270
         Width           =   840
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'HIGHLY recommended you change the text of this key!
Private Const SaveKey As String = "fdkl@$09123kjdflqaqweo!013"

Private Const SEP As String = "-----------------------------------------------------------------------------------"
Private Const DONE As String = "All files have been processed! ^_^"

Private Const NumAlgorithms As Long = 13
Private Enum Algorithms
    eNone = 0
    eRLE = 1
    eRLE_Looped = 2
    eLZMA = 3
    ePAQ8l = 4
    eDeflate64 = 5
    eMonkeyAudio = 6
    eXOR = 7
    eRC4 = 8
    eGOST = 9
    eSkipjack = 10
    eTwofish = 11
    eBlowfish = 12
    eCryptAPI = 13
End Enum
#If False Then
Private eNone, eRLE, eRLE_Looped, eLZMA, ePAQ8l, eDeflate64, eMonkeyAudio, eXOR, eRC4, eGOST, eSkipjack, eTwofish, eBlowfish, eCryptAPI
#End If

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Private Sub Log(ByVal LogText As String)

    'Add a line to the log
    If Len(LogList.Text) = 0 Then
        LogList.Text = LogList.Text & LogText
    Else
        LogList.Text = LogList.Text & vbNewLine & LogText
    End If
    LogList.Refresh

End Sub

Private Sub AExTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    AExTxt(Index).ToolTipText = "Extension of the file after being run through the defined algorithm"

End Sub

Private Sub AExTxt_GotFocus(Index As Integer)

    AExTxt(Index).BackColor = &H80000013

End Sub

Private Sub AExTxt_LostFocus(Index As Integer)

    AExTxt(Index).BackColor = &H80000005

End Sub

Private Sub AlgCmb_GotFocus(Index As Integer)

    AlgCmb(Index).BackColor = &H80000013

End Sub

Private Sub AlgCmb_LostFocus(Index As Integer)

    AlgCmb(Index).BackColor = &H80000005

End Sub

Private Sub CompBtn_Click()

    If MsgBox("Are you sure you wish to run the compression / encryption on the following files?", vbYesNo) = vbNo Then Exit Sub

    RunMainLoop False

End Sub

Private Sub LockEditing(ByVal eEnabled As Boolean)
Dim i As Long

    For i = 0 To AlgCmb.UBound
        ExTxt(i).Enabled = eEnabled
        SubTxt(i).Enabled = eEnabled
        AlgCmb(i).Enabled = eEnabled
        AExTxt(i).Enabled = eEnabled
        KeyTxt(i).Enabled = eEnabled
    Next i
    DirTxt.Enabled = eEnabled
    KillSourceChk.Enabled = eEnabled
    CompBtn.Enabled = eEnabled
    DecompBtn.Enabled = eEnabled

End Sub

Private Sub RunMainLoop(ByVal Reverse As Boolean)
Dim LastPercent As Single
Dim StartTime As Long
Dim FileData() As Byte
Dim FileList() As String
Dim ListLoop As Long
Dim FileLoop As Long
Dim s As String

    LockEditing False

    'Loop through all the algorithms in the list
    For ListLoop = 0 To AlgCmb.UBound
    
        'Check if the algorithm is used
        If AlgCmb(ListLoop).ListIndex <> eNone Then
        
            'Get the file list
            If Not Reverse Then
                GetFiles DirTxt.Text & SubTxt(ListLoop).Text, ExTxt(ListLoop).Text, FileList()
            Else
                GetFiles DirTxt.Text & SubTxt(ListLoop).Text, AExTxt(ListLoop).Text, FileList()
            End If
            
            'Check for information in the array
            Err.Number = 0
            On Error Resume Next
            s = UBound(FileList())
            If Err.Number <> 0 Or s = vbNullString Or FileList(0) = vbNullString Then
                On Error GoTo 0
                Log "No files found for routine " & ListLoop & "! Please make sure a valid extension and path was entered."
            Else
                On Error GoTo 0
                Log SEP
                Log "Running routine " & ListLoop & " with algorithm " & AlgCmb(ListLoop).List(AlgCmb(ListLoop).ListIndex) & " on " & UBound(FileList) + 1 & " files"
                Log "Processing files..."
                
                LastPercent = -1
                DoEvents
                StartTime = timeGetTime
                
                'Loop through all the files found
                For FileLoop = 0 To UBound(FileList)
                
                    'Update the completion percentage
                    If FileLoop = UBound(FileList) Then
                        LastPercent = 1
                        LogList.Text = LogList.Text & " " & Int(LastPercent * 100) & "%"
                        LogList.Refresh
                        DoEvents
                    Else
                        If (FileLoop / (UBound(FileList))) > LastPercent + 0.1 Then
                            LastPercent = FileLoop / (UBound(FileList))
                            LogList.Text = LogList.Text & " " & Int(LastPercent * 100) & "%"
                            LogList.Refresh
                            DoEvents
                        End If
                    End If
                
                    'Load the file into a byte array in memory
                    GetBytesFromFile FileList(FileLoop), FileData()
                    
                    'Run the appropriate algorithm on the file
                    RunAlgorithm AlgCmb(ListLoop).ListIndex, FileData(), Reverse, KeyTxt(ListLoop).Text
                    
                    'Save the file
                    If Not Reverse Then
                        SaveFile ChangeFileExtension(FileList(FileLoop), AExTxt(ListLoop).Text), FileData()
                    Else
                        SaveFile ChangeFileExtension(FileList(FileLoop), ExTxt(ListLoop).Text), FileData()
                    End If
                    
                    'Delete the source file only if dest <> source
                    If KillSourceChk.Value = 1 Then
                        If Dir$(FileList(FileLoop), vbNormal) <> vbNullString Then
                            If Not Reverse Then
                                If ChangeFileExtension(FileList(FileLoop), AExTxt(ListLoop).Text) <> FileList(FileLoop) Then
                                    Kill FileList(FileLoop)
                                End If
                            Else
                                If ChangeFileExtension(FileList(FileLoop), ExTxt(ListLoop).Text) <> FileList(FileLoop) Then
                                    Kill FileList(FileLoop)
                                End If
                            End If
                        End If
                    End If
                
                Next FileLoop
                
                Log "List " & ListLoop & " completed in " & Round((timeGetTime - StartTime) / 1000, 2) & " seconds."
            
            End If
        
        End If
        
NextListLoop:
    
    Next ListLoop
    
    Log SEP
    Log SEP
    Log DONE
    Log SEP
    Log SEP
    
    LockEditing True

End Sub

Private Function ChangeFileExtension(ByVal FileName As String, ByVal NewExtension As String) As String
Dim s() As String
Dim L As String
Dim i As Long

    If Len(NewExtension) = 0 Then
        ChangeFileExtension = FileName
        Exit Function
    End If

    'Chop off the old extension
    s = Split(FileName, ".")
    For i = 0 To UBound(s) - 1
        L = L & s(i) & "."
    Next i
    
    'Add the new extension
    ChangeFileExtension = L & NewExtension

End Function

Private Sub SaveFile(ByVal File As String, ByRef Bytes() As Byte)
Dim FileNum As Byte

    'Make sure the file doesn't already exist
    If Dir$(File, vbNormal) <> vbNullString Then Kill File

    'Get the next free file index to use
    FileNum = FreeFile
    
    'Open the file
    Open File For Binary Access Write As #FileNum
    
        'Put all the information into the file
        Put #FileNum, , Bytes()
        
    Close #FileNum

End Sub

Private Sub GetBytesFromFile(ByVal File As String, ByRef Bytes() As Byte)
Dim FileNum As Byte

    'Get the next free file index to use
    FileNum = FreeFile
    
    'Open the file
    Open File For Binary Access Read As #FileNum
    
        'Resize the bytes array by the size of the file
        ReDim Bytes(0 To LOF(FileNum) - 1)
        
        'Get all the information from the file
        Get #FileNum, , Bytes()
        
    Close #FileNum

End Sub

Private Sub RunAlgorithm(ByVal Algorithm As Algorithms, ByRef Bytes() As Byte, ByVal Reverse As Boolean, Optional ByVal Key As String = vbNullString)

    'Runs the specified algorithm on an array of bytes
    Select Case Algorithm
        Case eRLE
            If Not Reverse Then
                Compression_Compress_RLE Bytes(), False
            Else
                Compression_DeCompress_RLE Bytes()
            End If
        Case eRLE_Looped
            If Not Reverse Then
                Compression_Compress_RLELoop Bytes()
            Else
                Compression_DeCompress_RLELoop Bytes()
            End If
        Case eLZMA
            If Not Reverse Then
                Compression_Compress_LZMA Bytes()
            Else
                Compression_DeCompress_LZMA Bytes()
            End If
        Case ePAQ8l
            If Not Reverse Then
                Compression_Compress_PAQ8l Bytes()
            Else
                Compression_DeCompress_PAQ8l Bytes()
            End If
        Case eDeflate64
            If Not Reverse Then
                Compression_Compress_Deflate64 Bytes()
            Else
                Compression_DeCompress_Deflate64 Bytes()
            End If
        Case eMonkeyAudio
            If Not Reverse Then
                Compression_Compress_MonkeyAudio Bytes()
            Else
                Compression_DeCompress_MonkeyAudio Bytes()
            End If
        Case eXOR
            If Not Reverse Then
                Encryption_XOR_EncryptByte Bytes(), Key
            Else
                Encryption_XOR_DecryptByte Bytes(), Key
            End If
        Case eRC4
            If Not Reverse Then
                Encryption_RC4_EncryptByte Bytes(), Key
            Else
                Encryption_RC4_DecryptByte Bytes(), Key
            End If
        Case eBlowfish
            If Not Reverse Then
                Encryption_Blowfish_EncryptByte Bytes(), Key
            Else
                Encryption_Blowfish_DecryptByte Bytes(), Key
            End If
        Case eCryptAPI
            If Not Reverse Then
                Encryption_CryptAPI_EncryptByte Bytes(), Key
            Else
                Encryption_CryptAPI_DecryptByte Bytes(), Key
            End If
        Case eGOST
            If Not Reverse Then
                Encryption_Gost_EncryptByte Bytes(), Key
            Else
                Encryption_Gost_DecryptByte Bytes(), Key
            End If
        Case eSkipjack
            If Not Reverse Then
                Encryption_Skipjack_EncryptByte Bytes(), Key
            Else
                Encryption_Skipjack_DecryptByte Bytes(), Key
            End If
        Case eTwofish
            If Not Reverse Then
                Encryption_Twofish_EncryptByte Bytes(), Key
            Else
                Encryption_Twofish_DecryptByte Bytes(), Key
            End If
    End Select
    
End Sub

Private Sub DecompBtn_Click()

    If MsgBox("Are you sure you wish to run the decompression / decryption on the following files?", vbYesNo) = vbNo Then Exit Sub

    RunMainLoop True

End Sub

Private Sub DirTxt_GotFocus()

    DirTxt.BackColor = &H80000013

End Sub

Private Sub DirTxt_LostFocus()

    DirTxt.BackColor = &H80000005

End Sub

Private Sub ExTxt_GotFocus(Index As Integer)

    ExTxt(Index).BackColor = &H80000013

End Sub

Private Sub ExTxt_LostFocus(Index As Integer)

    ExTxt(Index).BackColor = &H80000005

End Sub

Private Sub ExTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ExTxt(Index).ToolTipText = "Extension of the file before being run through the defined algorithm"
    
End Sub

Private Sub Form_Load()
Dim i As Long
Dim j As Long
    timeBeginPeriod 1
    
    InitFilePaths
    
    LogList.Text = vbNullString
    
    'Set the default file path
    DirTxt.Text = App.Path & "\"
    DirTxt.SelStart = Len(DirTxt.Text)

    'Add the algorithms
    For i = 0 To AlgCmb.UBound
        With AlgCmb(i)
            For j = 0 To NumAlgorithms
                Select Case j
                    Case eNone
                        .AddItem "Unused"
                    Case eRLE
                        .AddItem "RLE (Comp)"
                    Case eRLE_Looped
                        .AddItem "RLE Looped (Comp)"
                    Case eLZMA
                        .AddItem "LZMA (Comp)"
                    Case eDeflate64
                        .AddItem "Deflate64 (Comp)"
                    Case eMonkeyAudio
                        .AddItem "MAC (wav only) (Comp)"
                    Case eXOR
                        .AddItem "XOR (Enc)"
                    Case eRC4
                        .AddItem "RC4 (Enc)"
                    Case eBlowfish
                        .AddItem "Blowfish (Enc)"
                    Case eCryptAPI
                        .AddItem "CryptAPI (Enc)"
                    Case eGOST
                        .AddItem "GOST (Enc)"
                    Case eSkipjack
                        .AddItem "Skipjack (Enc)"
                    Case eTwofish
                        .AddItem "Twofish (Enc)"
                    Case ePAQ8l
                        .AddItem "PAQ8l (Comp)"
                End Select
            Next j
        End With
        AlgCmb(i).ListIndex = 0
    Next
    
    'Load the configuration
    LoadConfig
    
End Sub

Private Sub SaveConfig()
Dim f As String
Dim i As Long

    'Set the file path
    f = Data2Path & "FileProcessor.ini"
    
    Encryption_XOR_DecryptFile f, f, SaveKey

    'Saves the configuration
    For i = 0 To AlgCmb.UBound
        Var_Write f, CStr(i), "BeforeExtension", ExTxt(i).Text
        Var_Write f, CStr(i), "AfterExtension", AExTxt(i).Text
        Var_Write f, CStr(i), "Algorithm", CStr(AlgCmb(i).ListIndex)
        Var_Write f, CStr(i), "Key", KeyTxt(i).Text
        Var_Write f, CStr(i), "SubFolder", SubTxt(i).Text
    Next i
    Var_Write f, "MISC", "DeleteChk", KillSourceChk.Value
    
    Encryption_XOR_EncryptFile f, f, SaveKey
    
End Sub

Private Sub LoadConfig()
Dim f As String
Dim i As Long

    'Set the file path
    f = Data2Path & "FileProcessor.ini"
    
    Encryption_XOR_DecryptFile f, f, SaveKey
    
    'Load the saved configuration
    For i = 0 To AlgCmb.UBound
        ExTxt(i).Text = Var_Get(f, CStr(i), "BeforeExtension")
        AExTxt(i).Text = Var_Get(f, CStr(i), "AfterExtension")
        AlgCmb(i).ListIndex = Val(Var_Get(f, CStr(i), "Algorithm"))
        KeyTxt(i).Text = Var_Get(f, CStr(i), "Key")
        SubTxt(i).Text = Var_Get(f, CStr(i), "SubFolder")
    Next i
    KillSourceChk.Value = Val(Var_Get(f, "MISC", "DeleteChk"))
    
    Encryption_XOR_EncryptFile f, f, SaveKey

End Sub

Private Sub GetFiles(ByVal FilePath As String, ByVal Extension As String, ByRef FileList() As String)
Dim TempList() As String
Dim i As Long
Dim K As Long
Dim L As Long
Dim e As Long
Dim s() As String

    On Error Resume Next

    'Start L at -1 so our first array element is 0
    L = -1
    
    'Make sure we were passed an empty array
    Erase FileList
    
    If Len(Extension) = 0 Or Extension = "*" Then
        
        'No extension, grab them all
        FileList = AllFilesInFolders(FilePath, True)
        
    Else
        
        'Get the list of all the files
        TempList = AllFilesInFolders(FilePath, True)
        
        'Loop throgh all the files and check for the extensions
        s = Split(Extension, ",")
        K = UBound(TempList)
        For i = 0 To K
        
            For e = 0 To UBound(s)
        
                'Check the extension
                If UCase$(Right$(TempList(i), Len(s(e)))) = UCase$(s(e)) Then
                
                    'The extension was a match, add to the final list
                    L = L + 1
                    ReDim Preserve FileList(L)
                    FileList(L) = TempList(i)
                    GoTo NextI
                    
                End If
                
            Next e
            
NextI:
            
        Next i
    
    End If
    
    On Error GoTo 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveConfig

End Sub

Private Sub KeyTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    KeyTxt(Index).ToolTipText = "Key used for the algorithm if the alorithm specified is an encryption"

End Sub

Private Sub KeyTxt_GotFocus(Index As Integer)

    KeyTxt(Index).BackColor = &H80000013

End Sub

Private Sub KeyTxt_LostFocus(Index As Integer)

    KeyTxt(Index).BackColor = &H80000005

End Sub

Private Sub LogLIst_Change()

    LogList.SelStart = Len(LogList.Text)

End Sub

Private Sub SubTxt_GotFocus(Index As Integer)

    SubTxt(Index).BackColor = &H80000013

End Sub

Private Sub SubTxt_LostFocus(Index As Integer)

    SubTxt(Index).BackColor = &H80000005

End Sub

Private Sub SubTxt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    SubTxt(Index).ToolTipText = "If a sub-folder is specified, only that folder and its sub folders are used"

End Sub

Private Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
Dim sSpaces As String

    sSpaces = Space$(1000)
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), File
    Var_Get = RTrim$(sSpaces)
    Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
    
End Function

Private Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

    'Writes a var to a text file
    writeprivateprofilestring Main, Var, Value, File

End Sub
