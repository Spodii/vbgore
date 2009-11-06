Attribute VB_Name = "Encryptions"
Option Explicit

'Credits goes to Fredrik Qvarfort for writing the algorithms in Visual Basic!

'***** Packet encryption options *****
Public Const PacketEncTypeNone As Byte = 0  'Use no encryption
Public Const PacketEncTypeRC4 As Byte = 1   'Use RC4 encryption
Public Const PacketEncTypeXOR As Byte = 2   'Use XOR encryption
Public Const PacketEncTypeServerIn As Byte = PacketEncTypeXOR   'Encryption for server in (or client out) packets
Public Const PacketEncTypeServerOut As Byte = PacketEncTypeNone 'Encryption for server out (or client in) packets

'These are only used if the PacketEncType is not PacketEncTypeNone
Public Const PacketEncKey1 As String = "al123vcAM !$@(2!@_#;241234vzxv!@$(*_DSZVc2123"  'First encryption key (any string works)
Public Const PacketEncKey2 As String = "t123409-nsad DS:!$N$MN!U_AKLJ!1240naga!@$)ZZV"  'Second encryption key (any string works)
Public Const PacketEncSeed As Long = 214    'The number to start from (any random value works)
Public Const PacketEncKeys As Byte = 40     'Number of packet encryption keys
Public PacketKeys() As String   'The array of keys generated to encrypt packets

'***** RC4 *****
Private m_sBoxRC4(0 To 255) As Integer

'***** SIMPLE XOR *****
Private m_XORKey() As Byte
Private m_XORKeyLen As Long
Private m_XORKeyValue As String

'***** MISC *****

'Key-dependant
Private m_KeyS As String

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Sub GenerateEncryptionKeys()

'*****************************************************************
'Generates a series of unique keys based off the parameters
'It is recommended you change this routine a bit for better safety for public games
'*****************************************************************
Dim Seed As Long
Dim Key1 As String
Dim Key2 As String
Dim B2() As Byte
Dim b() As Byte
Dim i As Long
Dim j As Long
Dim k As Long

    'Set the start values
    Seed = PacketEncSeed
    Key1 = PacketEncKey1
    Key2 = PacketEncKey2
    
    'Set the number of keys
    ReDim PacketKeys(0 To PacketEncKeys - 1)
    
    'Crop down the keys if needed
    If Len(Key2) > 32 Then Key2 = Left$(Key2, 32)
    If Len(Key1) > 32 Then Key1 = Left$(Key1, 32)
    
    'Loop through the needed keys
    For i = 0 To PacketEncKeys - 1
    
        'Generate a new seed
        Seed = (i * Seed) - 1
    
        'Jumble up the keys through XOR randomization
        b = StrConv(Key1, vbFromUnicode) 'Convert to a byte array
        B2 = StrConv(Key2, vbFromUnicode)
        For j = 0 To Len(Key1) - 1
            Seed = Seed + j + 1         'Modify the seed based on the placement of the character
            Do While Seed > 255         'Keep it in the byte range
                Seed = Seed - 255
            Loop
            b(j) = b(j) Xor Seed        'XOR the character by the seed
            B2(j) = B2(j) Xor CByte(Seed \ 2)
        Next j
        Key1 = StrConv(b, vbUnicode)     'Convert back to a string
        Key2 = StrConv(B2, vbUnicode)
            
        'Jumble up the keys through encryption
        Key2 = Encryption_RC4_EncryptString(Key2, Key1)
        Key1 = Encryption_RC4_EncryptString(Key1, Key2)
        
        'Store the key
        PacketKeys(i) = Key1
        
    Next i

End Sub

Private Function Encryption_Misc_FileExist(FileName As String) As Boolean

    On Error GoTo NotExist

    Call FileLen(FileName)
    Encryption_Misc_FileExist = True
    
    On Error GoTo 0

NotExist:

End Function


Public Sub Encryption_RC4_DecryptByte(ByteArray() As Byte, Optional Key As String)

'The same routine is used for encryption as well
'decryption so why not reuse some code and make
'this class smaller (that is it it wasn't for all
'those damn comments ;))

    Call Encryption_RC4_EncryptByte(ByteArray(), Key)

End Sub

Public Sub Encryption_RC4_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_RC4_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_RC4_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_RC4_DecryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the data into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Decrypt the byte array
    Call Encryption_RC4_DecryptByte(ByteArray(), Key)

    'Convert the byte array back into a string
    Encryption_RC4_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub Encryption_RC4_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim i As Long
Dim j As Long
Dim Temp As Byte
Dim Offset As Long
Dim OrigLen As Long
Dim CipherLen As Long
Dim sBox(0 To 255) As Integer

'Set the new key (optional)

    If (Len(Key) > 0) Then Encryption_RC4_SetKey Key

    'Create a local copy of the sboxes, this
    'is much more elegant than recreating
    'before encrypting/decrypting anything
    Call CopyMem(sBox(0), m_sBoxRC4(0), 512)

    'Get the size of the source array
    OrigLen = UBound(ByteArray) + 1
    CipherLen = OrigLen

    'Encrypt the data
    For Offset = 0 To (OrigLen - 1)
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        Temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = Temp
        ByteArray(Offset) = ByteArray(Offset) Xor (sBox((sBox(i) + sBox(j)) Mod 256))
    Next

End Sub

Public Sub Encryption_RC4_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_RC4_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_RC4_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_RC4_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the data into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call Encryption_RC4_EncryptByte(ByteArray(), Key)

    'Convert the byte array back into a string
    Encryption_RC4_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub Encryption_RC4_SetKey(New_Value As String)

Dim a As Long
Dim b As Long
Dim Temp As Byte
Dim Key() As Byte
Dim KeyLen As Long

'Do nothing if the key is buffered

    If (m_KeyS = New_Value) Then Exit Sub

    'Set the new key
    m_KeyS = New_Value

    'Save the password in a byte array
    Key() = StrConv(m_KeyS, vbFromUnicode)
    KeyLen = Len(m_KeyS)

    'Initialize s-boxes
    For a = 0 To 255
        m_sBoxRC4(a) = a
    Next a
    For a = 0 To 255
        b = (b + m_sBoxRC4(a) + Key(a Mod KeyLen)) Mod 256
        Temp = m_sBoxRC4(a)
        m_sBoxRC4(a) = m_sBoxRC4(b)
        m_sBoxRC4(b) = Temp
    Next

End Sub

Public Sub Encryption_XOR_DecryptByte(ByteArray() As Byte, Optional Key As String)

'The same routine is used for encryption
'as well as decryption so why not reuse
'some code and make this class smaller
'(that is if it wasn't for all those damn
'comments ;))

    Call Encryption_XOR_EncryptByte(ByteArray(), Key)

End Sub

Public Sub Encryption_XOR_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_XOR_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_XOR_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_XOR_DecryptString(Text As String, Optional Key As String) As String

Dim a As Long
Dim ByteLen As Long
Dim ByteArray() As Byte

'Convert the source string into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call Encryption_XOR_DecryptByte(ByteArray(), Key)

    'Return the encrypted data as a string
    Encryption_XOR_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub Encryption_XOR_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim Offset As Long
Dim ByteLen As Long
Dim ResultLen As Long

'Set the new key if one was provided

    If (Len(Key) > 0) Then Encryption_XOR_SetKey Key

    'Get the size of the source array
    ByteLen = UBound(ByteArray) + 1
    ResultLen = ByteLen

    'Loop thru the data encrypting it with simply XOR´ing with the key
    For Offset = 0 To (ByteLen - 1)
        ByteArray(Offset) = ByteArray(Offset) Xor m_XORKey(Offset Mod m_XORKeyLen)
    Next

End Sub

Public Sub Encryption_XOR_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_XOR_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_XOR_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_XOR_EncryptString(Text As String, Optional Key As String) As String

Dim a As Long
Dim ByteLen As Long
Dim ByteArray() As Byte

'Convert the source string into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call Encryption_XOR_EncryptByte(ByteArray(), Key)

    'Return the encrypted data as a string
    Encryption_XOR_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub Encryption_XOR_SetKey(New_Value As String)

'Do nothing if the key is buffered

    If (m_XORKeyValue = New_Value) Then Exit Sub

    'Set the new key and convert it to a
    'byte array for faster accessing later
    m_XORKeyValue = New_Value
    m_XORKeyLen = Len(New_Value)
    m_XORKey() = StrConv(m_XORKeyValue, vbFromUnicode)

End Sub
