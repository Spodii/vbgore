Attribute VB_Name = "General"
Option Explicit

Public FileNumber As Long
Public LastGrhLstIndex As Long
Public LastOldGrhLstIndex As Long

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Sub Main()

    InitManifest
    InitFilePaths
    Load frmMain
    frmMain.Show

End Sub

Public Sub LoadOldGrhs()
Dim FileNum As Byte
Dim ln As String
Dim s() As String
Dim i As Long

    'Clear the old list
    frmMain.OldGrhLst.Clear

    'Open the GrhRaw file
    FileNum = FreeFile
    Open Data2Path & "GrhRaw.txt" For Input As #FileNum
    
        'Loop through the whole file
        Do While Not EOF(FileNum)
    
            'Get the line
            Line Input #FileNum, ln
            
            'Check if it is a Grh line
            If UCase$(Left$(ln, 3)) = "GRH" Then
            
                'Grab the file number
                s() = Split(ln, "-")
                If Val(Right$(s(0), 1)) = 1 Then
                    If Val(s(1)) = FileNumber Then
                    
                        'Write the line
                        frmMain.OldGrhLst.AddItem ln
    
                    End If
                End If
                
            End If
            
        Loop
        
    Close #FileNum
  
End Sub

Public Sub RefreshImage(Optional ByVal MakeNew As Boolean = True)
Dim Index As Long
Dim b(0 To 2) As Byte
Dim IsRemoved As Boolean
Dim l As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim Rows As Long

'Values are blended together through the following routine:
'(r2/g2/b2 is the primary color, a is the alpha, precalculated for speed)
'r1 = (r1 * a) + (r2 - (r2 * (a / 255)))
'g1 = (g1 * a) + (g2 - (g2 * (a / 255)))
'b1 = (b1 * a) + (b2 - (b2 * (a / 255)))
Const Alpha As Single = (255 - 125) / 255    '125 is the alpha value of the grid
Const Add As Single = (255 - (255 * Alpha))

    'Copy the image from the backbuffer to the front buffer
    frmMain.PreviewPic.Cls
    frmMain.PreviewPic.Picture = frmMain.BackBufferPic.Picture
    
    'Check if we are drawing the grid
    If frmMain.GridChk.Value Then

        'Find the number of rows
        Rows = Val(frmMain.RowsTxt.Text)
        If Rows = -1 Then Rows = Val(frmMain.MaxRowsTxt.Text)
        If Rows <= 0 Then Exit Sub

        'Loop through the draw area
        '//!! Swap the For x / For y
        For x = Val(frmMain.StartXTxt.Text) To Val(frmMain.TexWidthTxt.Text) Step Val(frmMain.GridWidthTxt.Text)
            For y = Val(frmMain.StartYTxt.Text) To Val(frmMain.TexHeightTxt.Text) Step Val(frmMain.GridHeightTxt.Text)
                Index = ((x - Val(frmMain.StartXTxt.Text)) \ Val(frmMain.GridWidthTxt.Text)) + _
                    ((((y - Val(frmMain.StartYTxt.Text)) \ Val(frmMain.GridHeightTxt.Text))) * Rows)
                If x < Val(frmMain.TexWidthTxt.Text) Then
                    If y < Val(frmMain.TexHeightTxt.Text) Then
                    
                        'Check if it is removed
                        IsRemoved = (frmMain.GrhLst.List(Index) = "-removed-")
                        
                        'Draw removed grid item
                        If IsRemoved Then
                            
                            For i = 0 To Val(frmMain.GridWidthTxt.Text) - 1
                                SetPixel frmMain.PreviewPic.hdc, x, y + i, RGB(255, 0, 0)
                                SetPixel frmMain.PreviewPic.hdc, x + 32, y + i, RGB(255, 0, 0)
                                SetPixel frmMain.PreviewPic.hdc, x + i, y, RGB(255, 0, 0)
                                SetPixel frmMain.PreviewPic.hdc, x + i, y + 32, RGB(255, 0, 0)
                            Next i
                            
                        Else
                            
                            'Draw the vertical lines
                            If x > 0 Then
                                For i = 0 To Val(frmMain.GridWidthTxt.Text) - 1
                                
                                    'Get the image's pixel color in RGB format, modify it to get the alpha and set it
                                    l = GetPixel(frmMain.PreviewPic.hdc, x, y + i)
                                    CopyMemory b(0), l, 3
                                    SetPixel frmMain.PreviewPic.hdc, x, y + i, RGB(b(0) * Alpha + Add, b(1) * Alpha + Add, b(2) * Alpha + Add)
                                    
                                Next i
                            End If
                            
                            'Draw the horizontal lines
                            If y > 0 Then
                                For i = 0 To Val(frmMain.GridHeightTxt.Text) - 1
    
                                    'Get the image's pixel color in RGB format, modify it to get the alpha and set it
                                    l = GetPixel(frmMain.PreviewPic.hdc, x + i, y)
                                    CopyMemory b(0), l, 3
                                    SetPixel frmMain.PreviewPic.hdc, x + i, y, RGB(b(0) * Alpha + Add, b(1) * Alpha + Add, b(2) * Alpha + Add)
                                    
                                Next i
                            End If
                            
                        End If
                    
                    End If
                End If
            Next y
        Next x
        
    End If
    
    'Draw the selected new grh entry
    If frmMain.GrhLst.List(frmMain.GrhLst.ListIndex) <> "-removed-" Then
        If frmMain.GrhLst.ListIndex > -1 Then DrawSelectedGrh frmMain.GrhLst.List(frmMain.GrhLst.ListIndex), vbGreen
        If frmMain.OldGrhLst.ListIndex > -1 Then DrawSelectedGrh frmMain.OldGrhLst.List(frmMain.OldGrhLst.ListIndex), vbYellow
    End If
    LastGrhLstIndex = frmMain.GrhLst.ListIndex
    LastOldGrhLstIndex = frmMain.OldGrhLst.ListIndex
    
    'Make the grh entries
    If MakeNew Then MakeNewGrhs
    
End Sub

Public Sub DrawSelectedGrh(ByVal GrhString As String, ByVal Color As Long)
Dim s() As String
Dim x As Long
Dim y As Long
Dim Width As Long
Dim Height As Long

    'Split up the line and grab the important values
    s = Split(GrhString, "-")
    x = s(2)
    y = s(3)
    Width = s(4)
    Height = s(5)
    
    'Draw the rectangle
    frmMain.PreviewPic.Line (x, y)-(x + Width, y), Color
    frmMain.PreviewPic.Line (x, y)-(x, y + Height), Color
    frmMain.PreviewPic.Line (x + Width, y)-(x + Width, y + Height), Color
    frmMain.PreviewPic.Line (x, y + Height)-(x + Width, y + Height), Color

End Sub

Public Function FileExist(File As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    FileExist = (LenB(Dir$(File, FileType)) <> 0)

End Function

Public Sub LoadTexture(ByVal TexturePath As String)
Dim ImageInfo As CImageInfo
Dim s() As String

    'Display the texture
    frmMain.PreviewPic.Cls
    PngPictureLoad TexturePath, frmMain.BackBufferPic, False
    
    'Get the texture dimensions
    Set ImageInfo = New CImageInfo
    ImageInfo.ReadImageInfo TexturePath
    frmMain.TexWidthTxt.Text = ImageInfo.Width
    frmMain.TexHeightTxt.Text = ImageInfo.Height
    
    'Store the file number
    s = Split(frmMain.TexturePathTxt.Text, "\")
    FileNumber = Left$(s(UBound(s)), Len(s(UBound(s))) - 4)

    'Set the maximum number of rows and columns
    UpdateMaxRowsColumns
    
    'Load the Grhs from the texture that are already in GrhRaw.txt
    LoadOldGrhs
    
End Sub

Public Sub UpdateMaxRowsColumns()
Dim i As Long

    'Set the rows
    If Val(frmMain.GridWidthTxt.Text) <= 0 Then
        frmMain.MaxRowsTxt.Text = 0
    Else
        i = (Val(frmMain.TexWidthTxt.Text) - Val(frmMain.StartXTxt.Text)) \ Val(frmMain.GridWidthTxt.Text)
        If i < 0 Then frmMain.MaxRowsTxt.Text = 0 Else frmMain.MaxRowsTxt.Text = i
    End If
    
    'Set the columns
    If Val(frmMain.GridHeightTxt.Text) <= 0 Then
        frmMain.MaxColumnsTxt.Text = 0
    Else
        i = (Val(frmMain.TexHeightTxt.Text) - Val(frmMain.StartYTxt.Text)) \ Val(frmMain.GridHeightTxt.Text)
        If i < 0 Then frmMain.MaxColumnsTxt.Text = 0 Else frmMain.MaxColumnsTxt.Text = i
    End If
    
    'Draw the grid
    RefreshImage
    
End Sub

Public Sub MakeNewGrhs()
Dim TexWidth As Long
Dim TexHeight As Long
Dim GridWidth As Long
Dim GridHeight As Long
Dim Rows As Long
Dim Columns As Long
Dim x As Long
Dim y As Long
Dim GrhIndex As Long
Dim GrhLine As Long

    'Find the number of rows and columns
    Rows = Val(frmMain.RowsTxt.Text)
    Columns = Val(frmMain.ColumnsTxt.Text)
    If Rows = -1 Then Rows = Val(frmMain.MaxRowsTxt.Text)
    If Rows <= 0 Then Exit Sub
    If Columns = -1 Then Columns = Val(frmMain.MaxColumnsTxt.Text)
    If Columns <= 0 Then Exit Sub
    
    'Get the grid size
    GridWidth = Val(frmMain.GridWidthTxt.Text)
    If GridWidth <= 0 Then Exit Sub
    GridHeight = Val(frmMain.GridHeightTxt.Text)
    If GridHeight <= 0 Then Exit Sub
    
    'Get the texture size
    TexWidth = Val(frmMain.TexWidthTxt.Text)
    TexHeight = Val(frmMain.TexHeightTxt.Text)

    'Set the starting Grh index
    GrhIndex = Val(frmMain.StartGrhTxt.Text)

    'Clear the grh list
    frmMain.GrhLst.Clear
    
    'Hide the grh list (speeds updating up)
    frmMain.GrhLst.Enabled = False
    frmMain.GrhLst.Visible = False

    'Loop through the grid (by pixels)
    For y = Val(frmMain.StartYTxt.Text) To TexHeight Step GridHeight
        For x = Val(frmMain.StartXTxt.Text) To TexWidth Step GridWidth

            'Make sure it is in range
            If x >= 0 Then
                If y >= 0 Then
                    If x < TexWidth Then
                        If y < TexHeight Then
                        
                            'Get the next Grh line
                            GrhLine = GrhLine + 1
                            
                            'Find the next free Grh index
                            Do While Not IsFreeGrh(GrhIndex)
                                GrhIndex = GrhIndex + 1
                            Loop

                            'Add the entry to the list box
                            frmMain.GrhLst.AddItem "Grh" & GrhIndex & "=1-" & FileNumber & "-" & _
                                x & "-" & y & "-" & GridWidth & "-" & GridHeight
                                
                            'Move to the next Grh index
                            GrhIndex = GrhIndex + 1
                                
                        End If
                    End If
                End If
            End If
        
        Next x
    Next y
    
    'Show the grh list
    frmMain.GrhLst.Enabled = True
    frmMain.GrhLst.Visible = True
    
End Sub

Public Function IsFreeGrh(ByVal GrhIndex As Long) As Boolean
Dim c As Long

    'Calculate the free grh values
    IsFreeGrh = (LenB(Var_Get(Data2Path & "GrhRaw.txt", "A", "Grh" & GrhIndex)) = 0)

End Function

Public Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional ByVal DontLog As Byte = 0) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************

    Var_Get = Space$(1000)
    GetPrivateProfileString Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)

End Function

Public Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    WritePrivateProfileString Main, Var, Value, File

End Sub

Public Function BuildGrhString()
Dim i As Long

    'Put together the GrhString text
    For i = 0 To frmMain.GrhLst.ListCount
        If frmMain.GrhLst.List(i) <> "-removed-" Then
            BuildGrhString = BuildGrhString & frmMain.GrhLst.List(i) & vbNewLine
        End If
    Next i
    
    'Trim off the last vbNewLine
    BuildGrhString = Left$(BuildGrhString, Len(BuildGrhString) - Len(vbNewLine))

End Function
