Attribute VB_Name = "MD5"
Option Explicit

Private Type MD5_CONTEXT
    State(3) As Long
    Count(1) As Long
    Buffer(63) As Byte
End Type

Private Declare Sub MD5Init Lib "md5.dll" (lpContext As MD5_CONTEXT)
Private Declare Sub MD5Update Lib "md5.dll" (lpContext As MD5_CONTEXT, lpBuffer As Any, ByVal hInputLength As Long)
Private Declare Sub MD5Final Lib "md5.dll" (lpDigest As Any, lpContext As MD5_CONTEXT)

Public Function MD5_File(FileName As String) As String
Dim udtContext As MD5_CONTEXT
Dim bytDigest(15) As Byte
Dim bytData(63) As Byte
Dim FileNum As Byte

    'Get the free file
    FileNum = FreeFile

    'Open the file
    Open FileName For Binary Access Read As #FileNum
    
        'Init the MD5
        MD5Init udtContext
        
        'Loop through the file
        Do While Not EOF(FileNum)
            
            'Get the data
            Get #FileNum, , bytData
            
            'Convert to MD5 if possible
            If Loc(FileNum) < LOF(FileNum) Then MD5Update udtContext, bytData(0), 64
            
        Loop
        
        'Update the last part to MD5
        MD5Update udtContext, bytData(0), LOF(FileNum) Mod 64
        
    Close #FileNum
    
    'Call the finishing routine
    MD5Final bytDigest(0), udtContext
    
    'Return the result
    MD5_File = MD5_DigestToString(bytDigest)

End Function

Public Function MD5_String(SourceString As String) As String
Dim udtContext As MD5_CONTEXT
Dim bytDigest(15) As Byte
Dim bytData() As Byte
    
    'Init the MD5
    MD5Init udtContext
    
    'Turn the string into a byte array
    bytData = MD5_StringToArray(SourceString)
    
    'Do the MD5 operations on the byte array
    MD5Update udtContext, bytData(0), Len(SourceString)
    MD5Final bytDigest(0), udtContext
    
    'Return the byte array to string
    MD5_String = MD5_DigestToString(bytDigest)

End Function

Private Function MD5_DigestToString(Digest() As Byte) As String
Dim lngI As Long

    For lngI = 0 To UBound(Digest)
    
        'Pad with a "0" character if the HEX byte is less than 16 decimal
        If Digest(lngI) < 16 Then
            MD5_DigestToString = MD5_DigestToString & "0" & Hex$(Digest(lngI))
        Else
            MD5_DigestToString = MD5_DigestToString & LCase$(Hex$(Digest(lngI)))
        End If
        
    Next lngI

End Function

Private Function MD5_StringToArray(InString As String) As Byte()
    
    'Convert a string to a byte array
    If LenB(InString) = 0 Then
        ReDim MD5_StringToArray(0)
    Else
        MD5_StringToArray = StrConv(InString, vbFromUnicode)
    End If
    
End Function


