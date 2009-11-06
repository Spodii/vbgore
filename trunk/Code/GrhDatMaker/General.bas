Attribute VB_Name = "General"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Private Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    FileExist = (Dir$(File, FileType) <> "")

End Function

Private Function IsPowerof2(ByVal Number As Long) As Boolean
Dim i As Long
Dim j As Long

    For i = 1 To 12
        j = (2 ^ i)
        If Number = j Then
            IsPowerof2 = True
        Else
            If j > Number Then Exit Function
        End If
    Next i

End Function

Private Sub Main()

'*****************************************************************
'Loads GrhRaw.txt, parses and outputs Grh.dat
'*****************************************************************
Dim ImageSize As CImageInfo
Dim FileList() As String
Dim GrhBuffer() As Long  'Holds all the entered Grh values
Dim TempSplit() As String
Dim sX As Integer
Dim sY As Integer
Dim PixelWidth As Integer
Dim PixelHeight As Integer
Dim FileNum As Long
Dim NumFrames As Byte
Dim Frames() As Long
Dim Speed As Single
Dim LastGrh As Long
Dim Grh As Long
Dim Frame As Long
Dim ln As String
Dim TempLine As String
Dim NumRawFile As Long
Dim LastFile As Long
Dim Lines As Long

    InitFilePaths
    
    'Check for valid file sizes
    Set ImageSize = New CImageInfo
    FileList() = AllFilesInFolders(GrhPath, False)
    For FileNum = 0 To UBound(FileList)
        If Right$(FileList(FileNum), 4) = ".png" Then
            ImageSize.ReadImageInfo FileList(FileNum)
            If IsPowerof2(ImageSize.Width) = False Or IsPowerof2(ImageSize.Height) = False Then
                TempSplit = Split(FileList(FileNum))
                TempLine = TempLine & TempSplit(UBound(TempSplit))
            End If
        End If
    Next FileNum
    If TempLine <> vbNullString Then
        MsgBox "The following image files were found to not have sizes by powers of 2!" & vbNewLine & _
            "Leaving images not in powers of 2 will cause lots of graphical errors!" & vbNewLine & vbNewLine & TempLine & _
            vbNewLine & vbNewLine & "Grh.dat will still be made, but graphics may not appear correctly in-game.", vbOKOnly Or vbCritical
    End If
    Set ImageSize = Nothing
    TempLine = vbNullString
    FileNum = 0
    sX = 0
    sY = 0

    'Delete any old file
    If FileExist(DataPath & "grh.dat", vbNormal) = True Then Kill DataPath & "grh.dat"

    'Open new file
    Open DataPath & "grh.dat" For Binary As #1
    Seek #1, 1

    Open Data2Path & "GrhRaw.txt" For Input As #2

    'Do a loop to check for repeat numbers
    While Not EOF(2)
        DoEvents
        Line Input #2, TempLine
        If LCase$(Left$(TempLine, 3)) = "grh" Then
            TempLine = Right$(TempLine, Len(TempLine) - 3)
            If InStr(1, TempLine, "=") <= 0 Then GoTo ErrorHandler
            Grh = CLng(Left$(TempLine, InStr(1, TempLine, "=", vbTextCompare) - 1))
            Lines = Lines + 1
            ReDim Preserve GrhBuffer(1 To Lines)
            GrhBuffer(Lines) = Grh
        End If

    Wend

    'Check for duplicate entries (slow, but whatever - this tool doesn't need to be fast)
    For sX = 1 To Lines
        For sY = sX + 1 To Lines
            If GrhBuffer(sX) = GrhBuffer(sY) Then

                'Notify of duplicate
                If MsgBox("Duplcates entries of Grh" & GrhBuffer(sX) & " found. Do you wish to continue compiling Grh.dat?" & vbCrLf _
                          & "Duplicate Grh numbers can lead to graphical display failures and artifacts.", vbYesNo) = vbNo Then
                    Exit Sub
                End If

            End If
        Next sY
    Next sX

    Close #2

    Open Data2Path & "GrhRaw.txt" For Input As #2

    'Clear variables
    sX = 0
    sY = 0
    Lines = 0
    Grh = 0

    While Not EOF(2)
        DoEvents
        Line Input #2, TempLine
        If LCase$(Left$(TempLine, 3)) = "grh" Then
            TempLine = Right$(TempLine, Len(TempLine) - 3)

            If InStr(1, TempLine, "=") <= 0 Then GoTo ErrorHandler

            Grh = CLng(Left$(TempLine, InStr(1, TempLine, "=", vbTextCompare) - 1))
            If Grh > LastGrh Then LastGrh = Grh
            
            ln = Right$(TempLine, Len(TempLine) - Len(CStr(Grh)) - 1)

            If ln <> vbNullString Then
            
                'Split the string
                TempSplit() = Split(ln, "-")
            
                'Get number of frames and check
                NumFrames = Val(TempSplit(0))
                If NumFrames <= 0 Then GoTo ErrorHandler

                'Put grh number
                Put #1, , Grh
                
                'Put number of frames
                Put #1, , NumFrames

                If NumFrames > 1 Then
                    ReDim Frames(1 To NumFrames)
                    
                    'Read a animation GRH set
                    For Frame = 1 To NumFrames
                    
                        'Check and put each frame
                        Frames(Frame) = Val(TempSplit(Frame))
                        If Frames(Frame) <= 0 Or Frames(Frame) > LastGrh Then GoTo ErrorHandler
                        Put #1, , Frames(Frame)
            
                    Next Frame
                    
                    'Check and put speed
                    Speed = CSng(TempSplit(NumFrames + 1))
                    If Speed = 0 Then GoTo ErrorHandler
                    Put #1, , Speed
                    
                Else
                    'check and put normal GRH data
                    FileNum = Val(TempSplit(1))
                    If FileNum <= 0 Then GoTo ErrorHandler
                    If FileNum > LastFile Then LastFile = FileNum

                    Put #1, , FileNum

                    sX = Val(TempSplit(2))
                    If sX < 0 Then GoTo ErrorHandler
                    Put #1, , sX

                    sY = Val(TempSplit(3))
                    If sY < 0 Then GoTo ErrorHandler
                    Put #1, , sY

                    PixelWidth = Val(TempSplit(4))
                    If PixelWidth <= 0 Then GoTo ErrorHandler
                    Put #1, , PixelWidth

                    PixelHeight = Val(TempSplit(5))
                    If PixelHeight <= 0 Then GoTo ErrorHandler
                    Put #1, , PixelHeight
                End If
            End If
        End If
    Wend

    Close #2
    Close #1

    WriteVar DataPath & "grh.ini", "INIT", "NumGrhFiles", CStr(LastFile)
    WriteVar DataPath & "grh.ini", "INIT", "NumGrhs", CStr(LastGrh)

    'Display finish box
    MsgBox "Finished.", vbOKOnly

    'Unload
    Erase GrhBuffer
    End
    
Exit Sub
ErrorHandler:

    Close #2
    Close #1
    MsgBox "Error on Grh" & Grh & "!", vbOKOnly Or vbCritical

End Sub

Private Sub WriteVar(File As String, Main As String, Var As String, value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    WritePrivateProfileString Main, Var, value, File

End Sub

