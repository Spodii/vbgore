Attribute VB_Name = "General"
Option Explicit

'Local socket ID
Public LocalID As Long
Public Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    PStatus As Double
    LastAmount As Double
End Type
Public Info As da

Public RecFileName As String       'File name (path) that we are recieving
Public RecFileSize As Long         'Official file size recieved from server
Public RecFileHash As String * 32  'Hash received from the server

Public WriteFileNum As Byte    'File number being written to

Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub GOREsock_Connection(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_Close(inSox As Long)

    frmMain.ConnectCmd.Enabled = True

End Sub

Public Sub GOREsock_DataArrival(inSox As Long, inData() As Byte)
Dim TempArray() As String
Dim TempStr As String
Dim AppPath As String
Dim FileNum As Byte
Dim rData As String
Dim i As Integer
Dim j As Long
Dim b() As Byte

    'Turn the data into a string
    rData = StrConv(inData(), vbUnicode)

    'Check for a send request
    If Left$(rData, 19) = "****sendrequest****" Then

        'Retrieve the file name, size and time
        TempArray = Split(rData, "|")
        RecFileName = TempArray(1)
        RecFileSize = TempArray(2)
        RecFileHash = TempArray(3)
        frmMain.FileLbl.Caption = RecFileName & " (" & RecFileSize & ")"

        'If the application path is "x:\", set it to "x:" or else we will have "x:\\" (where x is any letter)
        AppPath = App.Path
        If Len(AppPath) = 3 Then AppPath = Left$(AppPath, 2)

        'Close old file if open
        If WriteFileNum > 0 Then Close #WriteFileNum
        WriteFileNum = 0

        'Create directory if not yet made
        MakeSureDirectoryPathExists AppPath & RecFileName & ".compressed"

        'Open file to write to
        WriteFileNum = FreeFile
        Open AppPath & RecFileName & ".compressed" For Binary Access Write As #WriteFileNum

        'Check if the file write time is same
        If Engine_FileExist(AppPath & RecFileName, vbNormal) Then
            
            'If the file does exist, compair it to the recieved MD5 hash to see if it is the same file
            TempStr = MD5_File(App.Path & RecFileName)
            If TempStr = RecFileHash Then
            
                'The file is the same, do not update
                b() = StrConv("****no****", vbFromUnicode)
                GOREsock_SendData LocalID, b()
                Exit Sub
                
            End If

        End If

        'File is different, request the update
        b() = StrConv("****ok****", vbFromUnicode)
        GOREsock_SendData LocalID, b()
        frmMain.StatusLbl.Caption = "Downloading..."
        Exit Sub

    End If

    'End of file reached
    If Right$(rData, 17) = "****ENDOFFILE****" Then
        TempStr = Left$(rData, Len(rData) - 17) 'Crop out the ENDOFFILE to recieve the last bit of data
        If Len(TempStr) Then Put #WriteFileNum, , TempStr   'Write the last data
        Close #WriteFileNum     'Close the file since we're done
        Compression_DeCompress App.Path & RecFileName & ".compressed", App.Path & RecFileName, LZW 'Take the compressed file and decompress it
        'Kill App.Path & RecFileName & ".compressed" 'Kill the compressed file
        frmMain.PercentLbl.Caption = "0%"
        Exit Sub
    End If

    'Done downloading
    If Right$(rData, 12) = "****DONE****" Then
        GOREsock_Shut LocalID
        GOREsock_ShutDown
        frmMain.StatusLbl.Caption = "Download Successful!"
        frmMain.FileLbl.Caption = ""
        frmMain.PercentLbl.Caption = "100%"
        
        'Make sure the file is closed
        On Error Resume Next
        Close #WriteFileNum
        On Error GoTo 0
        
        'Clean up the compressed files
        TempArray = AllFilesInFolders(App.Path, True)
        For j = 1 To UBound(TempArray)
            If Right$(TempArray(j), 11) = ".compressed" Then Kill TempArray(j)
        Next j
        
        'Load the client
        If MsgBox("The update has been completed! Do you wish to run the client now?", vbYesNo) = vbYes Then
            ShellExecute frmMain.hWnd, vbNullString, App.Path & "\GameClient.exe", "-sdf@041jkdf0)21`~", vbNullString, 1
        End If
        
        'Unload the updater
        GOREsock_Shut LocalID
        DoEvents
        GOREsock_UnHook
        DoEvents
        
        'Initiate the closedown (gives the socket time to unload)
        frmMain.CloseTimer.Enabled = True
        
        Exit Sub
    End If

    'If not the above, then we are *hopefully* recieving the file data
    Put #WriteFileNum, , rData

End Sub

Public Sub GOREsock_RecvProgress(inSox As Long, bytesRecv As Long, bytesRemaining As Long)

    'Update recieved percentage
    frmMain.PercentLbl.Caption = (Round(bytesRecv / (bytesRecv + bytesRemaining), 2) * 100) & "%"

End Sub

Sub GOREsock_Connecting(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_SendProgress(inSox As Long, bytesSent As Long, bytesRemaining As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Sub GOREsock_SendComplete(inSox As Long)

'*********************************************
'Empty procedure
'*********************************************

End Sub

Public Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean

'Checks if a file exists

    Engine_FileExist = (Dir$(File, FileType) <> "")

End Function
