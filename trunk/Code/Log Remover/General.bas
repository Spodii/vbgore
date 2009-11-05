Attribute VB_Name = "General"
Option Explicit

Sub Main()
Dim LinesRemoved As Long
Dim FileLine() As String
Dim TempSplit() As String
Dim FilePaths() As String
Dim NumFiles As Byte
Dim FileNum As Byte
Dim s As String
Dim i As Long
Dim j As Long

    'Confirm
    If MsgBox("Are you sure you wish to remove all comments from the server?", vbYesNo) = vbNo Then End
    If MsgBox("The process is about to begin." & vbCrLf & "Please make sure you back up all your code before pressing OK!", vbOKCancel) = vbCancel Then End
    If MsgBox("This is your last chance. Are you POSITIVE you wish to remove all the logs from the server forever?" & vbCrLf & "Press Cancel NOW to abort, or else they will be gone forever!", vbOKCancel) = vbCancel Then End
    
    'Gather all the code files
    FileNum = FreeFile
    Open App.Path & "\GameServer.vbp" For Input As #FileNum
    
        'Loop until we reach the end
        Do While Not EOF(FileNum)
        
            'Get the line
            Line Input #FileNum, s
            
            'Check for a module, class module or form
            If UCase$(Left$(s, 7)) = "MODULE=" Or UCase$(Left$(s, 6)) = "CLASS=" Then
            
                'Split the file path from the rest and store the full file path
                TempSplit() = Split(s, "; ")
                NumFiles = NumFiles + 1
                ReDim Preserve FilePaths(1 To NumFiles)
                FilePaths(NumFiles) = App.Path & "\" & TempSplit(UBound(TempSplit))
                
            ElseIf UCase$(Left$(s, 5)) = "FORM=" Then
            
                'Forms are written differently, so load accordingly
                TempSplit() = Split(s, "=")
                NumFiles = NumFiles + 1
                ReDim Preserve FilePaths(1 To NumFiles)
                FilePaths(NumFiles) = App.Path & "\" & TempSplit(UBound(TempSplit))
                
            End If
            
        Loop
        
    Close #FileNum
    
    'Loop through the code files
    For i = 1 To NumFiles
        
        'Open the file and get all the lines
        j = 0
        FileNum = FreeFile
        Open FilePaths(i) For Input As #FileNum
            Do While EOF(FileNum) = False
                Line Input #FileNum, s
                
                'Ignore //\\LOGLINE//\\ lines
                If Not Right$(s, 15) = "//\\LOGLINE//\\" Then
                    j = j + 1
                    ReDim Preserve FileLine(1 To j)
                    FileLine(j) = s & vbCrLf
                Else
                    LinesRemoved = LinesRemoved + 1
                End If

            Loop
        Close #FileNum
        
        'Kill the file to remove the old content
        Kill FilePaths(i)
        
        'Write the information back into the file
        FileNum = FreeFile
        Open FilePaths(i) For Binary As #FileNum
            For j = 1 To UBound(FileLine())
                Put #FileNum, , FileLine(j)
            Next j
        Close #FileNum
        
    Next i
    
    MsgBox "All //\\LOGLINE//\\ lines have been successfully removed!" & vbCrLf & "A total of " & LinesRemoved & " were removed.", vbOKOnly
    
    End
    
End Sub
