Attribute VB_Name = "ImageInfo"
'########################################'
'   Programmed By Inderpal Singh         '
'   Email: inderpal0@hotmail.com         '
'   Date: April 9, 2002                  '
'   Homepage: http://connect.to/lanserver'
'########################################'

Option Explicit
Private bBuf() As Byte
Private Image_Width As Long
Private Image_Height As Long
Private Image_Type As eImageType
Private Image_FileSize As Long

Public Enum eImageType
    itUNKNOWN = 0
    itGIF = 1
    itJPEG = 2
    itPNG = 3
    itBMP = 4
End Enum

Private Type WordBytes
    byte1 As Byte
    byte2 As Byte
End Type

Private Type DWordBytes
    byte1 As Byte
    byte2 As Byte
End Type

Private Type WordWrapper
    Value As Integer
End Type

Public Property Get ImageWidth() As Long
    ImageWidth = Image_Width
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = Image_Height
End Property

Public Property Get ImageType() As eImageType
    ImageType = Image_Type
End Property

Public Property Get FileSize() As Long
    FileSize = Image_FileSize
End Property

Public Sub ReadImageInfo(sFileName As String)
    
    Dim i As Long
    Dim Size As Integer
    
    Image_Width = 0
    Image_Height = 0
    Image_FileSize = 0
    Image_Type = itUNKNOWN
    Size = FreeFile
    Open sFileName For Binary As Size
    Image_FileSize = LOF(Size)
    ReDim bBuf(Image_FileSize)
    Get #Size, 1, bBuf()
    Close Size
'Check For PNG
    If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
        Image_Type = itPNG
        If Image_Type Then
            Image_Width = BEWord(18)
            Image_Height = BEWord(22)
        End If
    End If
' Check For GIF
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
        Image_Type = itGIF
        Image_Width = LEWord(6)
        Image_Height = LEWord(8)
    End If
' Check For BMP
    If bBuf(0) = 66 And bBuf(1) = 77 Then
        Image_Type = itBMP
        Image_Width = LEWord(18)
        Image_Height = LEWord(22)
    End If
' Check For JPEG
    If Image_Type = itUNKNOWN Then
        Dim lPos As Long
        Do
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 And bBuf(lPos + 2) = &HFF) _
            Or (lPos >= Image_FileSize - 10) Then Exit Do
            lPos = lPos + 1
        Loop
        lPos = lPos + 2
        If lPos >= Image_FileSize - 10 Then Exit Sub
        Do
            Do
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= Image_FileSize - 10 Then Exit Sub
            Loop
            lPos = lPos + 1
            If (bBuf(lPos) >= &HC0) And (bBuf(lPos) <= &HC3) Then Exit Do
            lPos = lPos + BEWord(lPos + 1)
            If lPos >= Image_FileSize - 10 Then Exit Sub
        Loop
        Image_Type = itJPEG
        Image_Height = BEWord(lPos + 4)
        Image_Width = BEWord(lPos + 6)
    End If
    ReDim bBuf(0)
End Sub

Private Function LEWord(position As Long) As Long
    Dim x1 As WordBytes
    Dim x2 As WordWrapper
    x1.byte1 = bBuf(position)
    x1.byte2 = bBuf(position + 1)
    LSet x2 = x1
    LEWord = x2.Value
End Function

Private Function BEWord(position As Long) As Long
    Dim x1 As WordBytes
    Dim x2 As WordWrapper
    x1.byte1 = bBuf(position + 1)
    x1.byte2 = bBuf(position)
    LSet x2 = x1
    BEWord = x2.Value
End Function


