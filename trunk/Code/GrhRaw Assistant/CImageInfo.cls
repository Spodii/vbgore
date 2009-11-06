VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' CImageInfo
'
' Author: David Crowell
' davidc@qtm.net
' http://www.qtm.net/~davidc
'
' Released to the public domain
' use however you wish
'
' CImageInfo will get the image type ,dimensions, and
' color depth from JPG, PNG, BMP, and GIF files.
'
' version date: June 16, 1999
'
' http://www.wotsit.org is a good source of
' file format information.  This code would not have been
' possible without the files I found there.
'

' read-only properties

Private Const BUFFERSIZE As Long = 65535

Public Enum eImageType
    itUNKNOWN = 0
    itGIF = 1
    itJPEG = 2
    itPNG = 3
    itBMP = 4
End Enum

Private m_Width As Long
Private m_Height As Long
Private m_Depth As Byte
Private m_ImageType As eImageType

Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get Depth() As Byte
    Depth = m_Depth
End Property

Public Property Get ImageType() As eImageType
    ImageType = m_ImageType
End Property

Public Sub ReadImageInfo(sFileName As String)
' This is the sub to call to retrieve information on a file.
    
    ' Byte array buffer to store part of the file
    Dim bBuf(BUFFERSIZE) As Byte
    ' Open file number
    Dim iFN As Integer
    
    ' Set all properties to default values
    m_Width = 0
    m_Height = 0
    m_Depth = 0
    m_ImageType = itUNKNOWN
    
    ' here we will load the first part of a file into a byte
'array the amount of the file stored here depends on
'the BUFFERSIZE constant
    iFN = FreeFile
    Open sFileName For Binary As iFN
    Get #iFN, 1, bBuf()
    Close iFN
    
    If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
    ' this is a PNG file
    
        m_ImageType = itPNG
        
        ' get bit depth
        Select Case bBuf(25)
            Case 0
            ' greyscale
                m_Depth = bBuf(24)
                
            Case 2
            ' RGB encoded
                m_Depth = bBuf(24) * 3
                
            Case 3
            ' Palette based, 8 bpp
                m_Depth = 8
                
            Case 4
            ' greyscale with alpha
                m_Depth = bBuf(24) * 2
                
            Case 6
            ' RGB encoded with alpha
                m_Depth = bBuf(24) * 4
                
            Case Else
            ' This value is outside of it's normal range, so
            'we'll assume
            ' that this is not a valid file
                m_ImageType = itUNKNOWN
                
        End Select
        
        If m_ImageType Then
        ' if the image is valid then
        
            ' get the width
            m_Width = Mult(bBuf(19), bBuf(18))
            
            ' get the height
            m_Height = Mult(bBuf(23), bBuf(22))
        End If
        
    End If
    
    If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
    ' this is a GIF file
        
        m_ImageType = itGIF
        
        ' get the width
        m_Width = Mult(bBuf(6), bBuf(7))
        
        ' get the height
        m_Height = Mult(bBuf(8), bBuf(9))
        
        ' get bit depth
        m_Depth = (bBuf(10) And 7) + 1
    End If
    
    If bBuf(0) = 66 And bBuf(1) = 77 Then
    ' this is a BMP file
    
        m_ImageType = itBMP
        
        ' get the width
        m_Width = Mult(bBuf(18), bBuf(19))
        
        ' get the height
        m_Height = Mult(bBuf(22), bBuf(23))
        
        ' get bit depth
        m_Depth = bBuf(28)
    End If

    If m_ImageType = itUNKNOWN Then
    ' if the file is not one of the above type then
    ' check to see if it is a JPEG file
        Dim lPos As Long
        
        Do
        ' loop through looking for the byte sequence FF,D8,FF
        ' which marks the begining of a JPEG file
        ' lPos will be left at the postion of the start
            If (bBuf(lPos) = &HFF And bBuf(lPos + 1) = &HD8 _
                 And bBuf(lPos + 2) = &HFF) _
                 Or (lPos >= BUFFERSIZE - 10) Then Exit Do
            
            ' move our pointer up
            lPos = lPos + 1
            
        ' and continue
        Loop
        
        lPos = lPos + 2
        If lPos >= BUFFERSIZE - 10 Then Exit Sub
        
        
        Do
' loop through the markers until we find the one
'starting with FF,C0 which is the block containing the
'image information
        
            Do
            ' loop until we find the beginning of the next marker
                If bBuf(lPos) = &HFF And bBuf(lPos + 1) _
               <> &HFF Then Exit Do
                lPos = lPos + 1
                If lPos >= BUFFERSIZE - 10 Then Exit Sub
            Loop
            
            ' move pointer up
            lPos = lPos + 1
            
            Select Case bBuf(lPos)
                Case &HC0 To &HC3, &HC5 To &HC7, &HC9 To &HCB, _
                &HCD To &HCF
                ' we found the right block
                    Exit Do
            End Select
            
            ' otherwise keep looking
            lPos = lPos + Mult(bBuf(lPos + 2), bBuf(lPos + 1))
            
            ' check for end of buffer
            If lPos >= BUFFERSIZE - 10 Then Exit Sub
            
        Loop
        
        ' If we've gotten this far it is a JPEG and we are ready
        ' to grab the information.
        
        m_ImageType = itJPEG
        
        ' get the height
        m_Height = Mult(bBuf(lPos + 5), bBuf(lPos + 4))
        
        ' get the width
        m_Width = Mult(bBuf(lPos + 7), bBuf(lPos + 6))
        
        ' get the color depth
        m_Depth = bBuf(lPos + 8) * 8
        
    End If
    
End Sub
Private Function Mult(lsb As Byte, msb As Byte) As Long
    Mult = lsb + (msb * CLng(256))
End Function

