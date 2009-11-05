Attribute VB_Name = "EncryptionsEX"
Option Explicit

'Credits goes to Fredrik Qvarfort for writing the algorithms in Visual Basic!

'***** Packet encryption options *****
Public Const PacketEncTypeNone As Byte = 0
Public Const PacketEncTypeRC4 As Byte = 1
Public Const PacketEncTypeXOR As Byte = 2
Public Const PacketEncType As Byte = PacketEncTypeNone
Public Const PacketEncKey As String = "L234)Zlka;2341DFLJK"

'***** BLOWFISH *****
'Constant for number of rounds
Private Const ROUNDS = 16

'Keydependant p-boxes and s-boxes
Private m_pBox(0 To ROUNDS + 1) As Long
Private m_sBoxBF(0 To 3, 0 To 255) As Long

'***** CRYPTAPI *****
Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const KEY_CONTAINER As String = "Metallica"
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_CLASS_HASH As Long = 32768
Private Const ALG_TYPE_ANY As Long = 0
Private Const ALG_TYPE_STREAM As Long = 2048
Private Const ALG_SID_RC4 As Long = 1
Private Const ALG_SID_MD5 As Long = 3
Private Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const ENCRYPT_ALGORITHM As Long = CALG_RC4
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long

'***** GOST *****
Private K(1 To 8) As Long
Private k87(0 To 255) As Byte
Private k65(0 To 255) As Byte
Private k43(0 To 255) As Byte
Private k21(0 To 255) As Byte
Private sBox(0 To 7, 0 To 255) As Byte

'***** RC4 *****
Private m_sBoxRC4(0 To 255) As Integer

'***** SIMPLE XOR *****
Private m_XORKey() As Byte
Private m_XORKeyLen As Long
Private m_XORKeyValue As String

'***** SKIPJACK *****
'To store a buffered key
Private m_SJKeyValue As String

'Key-dependant data
Private m_SJF(0 To 255) As Byte
Private m_SJKey(0 To 127) As Byte

'***** TWOFISH *****
Public Enum TWOFISHKEYLENGTH
    TWOFISH_256 = 256
    TWOFISH_196 = 196
    TWOFISH_128 = 128
    TWOFISH_64 = 64
End Enum
#If False Then
Private TWOFISH_256, TWOFISH_196, TWOFISH_128, TWOFISH_64
#End If

Private Const ROUNDSTF = 16
Private Const BLOCK_SIZETF = 16

Private Const INPUT_WHITEN = 0
Private Const OUTPUT_WHITEN = INPUT_WHITEN + BLOCK_SIZETF / 4
Private Const ROUND_SUBKEYS = OUTPUT_WHITEN + BLOCK_SIZETF / 4

Private Const GF256_FDBK_2 = &H169 / 2
Private Const GF256_FDBK_4 = &H169 / 4

Private MDS(0 To 3, 0 To 255) As Long
Private P(0 To 1, 0 To 255) As Byte

'Key-dependant data
Private sBoxTF(0 To 1023) As Long
Private sKeyTF() As Long

'***** MISC *****
'To be able to run optimized code (addition without the slow UnsignedAdd procedure we
'need to know if we are running in compiled mode or in the IDE)
Private m_RunningCompiled As Boolean

'Store buffered key
Private m_KeyValue As String

'Key-dependant
Private m_Key(0 To 47, 1 To 16) As Byte
Private m_KeyS As String
Private m_Encryption_Misc_InitHex As Boolean
Private m_ByteToHex(0 To 255, 0 To 1) As Byte
Private m_HexToByte(48 To 70, 48 To 70) As Byte

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Sub Encryption_Twofish_Init()

Dim i As Long
Dim j As Long
Dim m1(0 To 1) As Long
Dim mX(0 To 1) As Long
Dim mY(0 To 1) As Long

'We need to check if we are running in compiled
'(EXE) mode or in the IDE, this will allow us to
'use optimized code with unsigned integers in
'compiled mode without any overflow errors when
'running the code in the IDE

    On Local Error Resume Next
        m_RunningCompiled = ((2147483647 + 1) < 0)

        'Initialize P(0,..) array
        P(0, 0) = &HA9
        P(0, 1) = &H67
        P(0, 2) = &HB3
        P(0, 3) = &HE8
        P(0, 4) = &H4
        P(0, 5) = &HFD
        P(0, 6) = &HA3
        P(0, 7) = &H76
        P(0, 8) = &H9A
        P(0, 9) = &H92
        P(0, 10) = &H80
        P(0, 11) = &H78
        P(0, 12) = &HE4
        P(0, 13) = &HDD
        P(0, 14) = &HD1
        P(0, 15) = &H38
        P(0, 16) = &HD
        P(0, 17) = &HC6
        P(0, 18) = &H35
        P(0, 19) = &H98
        P(0, 20) = &H18
        P(0, 21) = &HF7
        P(0, 22) = &HEC
        P(0, 23) = &H6C
        P(0, 24) = &H43
        P(0, 25) = &H75
        P(0, 26) = &H37
        P(0, 27) = &H26
        P(0, 28) = &HFA
        P(0, 29) = &H13
        P(0, 30) = &H94
        P(0, 31) = &H48
        P(0, 32) = &HF2
        P(0, 33) = &HD0
        P(0, 34) = &H8B
        P(0, 35) = &H30
        P(0, 36) = &H84
        P(0, 37) = &H54
        P(0, 38) = &HDF
        P(0, 39) = &H23
        P(0, 40) = &H19
        P(0, 41) = &H5B
        P(0, 42) = &H3D
        P(0, 43) = &H59
        P(0, 44) = &HF3
        P(0, 45) = &HAE
        P(0, 46) = &HA2
        P(0, 47) = &H82
        P(0, 48) = &H63
        P(0, 49) = &H1
        P(0, 50) = &H83
        P(0, 51) = &H2E
        P(0, 52) = &HD9
        P(0, 53) = &H51
        P(0, 54) = &H9B
        P(0, 55) = &H7C
        P(0, 56) = &HA6
        P(0, 57) = &HEB
        P(0, 58) = &HA5
        P(0, 59) = &HBE
        P(0, 60) = &H16
        P(0, 61) = &HC
        P(0, 62) = &HE3
        P(0, 63) = &H61
        P(0, 64) = &HC0
        P(0, 65) = &H8C
        P(0, 66) = &H3A
        P(0, 67) = &HF5
        P(0, 68) = &H73
        P(0, 69) = &H2C
        P(0, 70) = &H25
        P(0, 71) = &HB
        P(0, 72) = &HBB
        P(0, 73) = &H4E
        P(0, 74) = &H89
        P(0, 75) = &H6B
        P(0, 76) = &H53
        P(0, 77) = &H6A
        P(0, 78) = &HB4
        P(0, 79) = &HF1
        P(0, 80) = &HE1
        P(0, 81) = &HE6
        P(0, 82) = &HBD
        P(0, 83) = &H45
        P(0, 84) = &HE2
        P(0, 85) = &HF4
        P(0, 86) = &HB6
        P(0, 87) = &H66
        P(0, 88) = &HCC
        P(0, 89) = &H95
        P(0, 90) = &H3
        P(0, 91) = &H56
        P(0, 92) = &HD4
        P(0, 93) = &H1C
        P(0, 94) = &H1E
        P(0, 95) = &HD7
        P(0, 96) = &HFB
        P(0, 97) = &HC3
        P(0, 98) = &H8E
        P(0, 99) = &HB5
        P(0, 100) = &HE9
        P(0, 101) = &HCF
        P(0, 102) = &HBF
        P(0, 103) = &HBA
        P(0, 104) = &HEA
        P(0, 105) = &H77
        P(0, 106) = &H39
        P(0, 107) = &HAF
        P(0, 108) = &H33
        P(0, 109) = &HC9
        P(0, 110) = &H62
        P(0, 111) = &H71
        P(0, 112) = &H81
        P(0, 113) = &H79
        P(0, 114) = &H9
        P(0, 115) = &HAD
        P(0, 116) = &H24
        P(0, 117) = &HCD
        P(0, 118) = &HF9
        P(0, 119) = &HD8
        P(0, 120) = &HE5
        P(0, 121) = &HC5
        P(0, 122) = &HB9
        P(0, 123) = &H4D
        P(0, 124) = &H44
        P(0, 125) = &H8
        P(0, 126) = &H86
        P(0, 127) = &HE7
        P(0, 128) = &HA1
        P(0, 129) = &H1D
        P(0, 130) = &HAA
        P(0, 131) = &HED
        P(0, 132) = &H6
        P(0, 133) = &H70
        P(0, 134) = &HB2
        P(0, 135) = &HD2
        P(0, 136) = &H41
        P(0, 137) = &H7B
        P(0, 138) = &HA0
        P(0, 139) = &H11
        P(0, 140) = &H31
        P(0, 141) = &HC2
        P(0, 142) = &H27
        P(0, 143) = &H90
        P(0, 144) = &H20
        P(0, 145) = &HF6
        P(0, 146) = &H60
        P(0, 147) = &HFF
        P(0, 148) = &H96
        P(0, 149) = &H5C
        P(0, 150) = &HB1
        P(0, 151) = &HAB
        P(0, 152) = &H9E
        P(0, 153) = &H9C
        P(0, 154) = &H52
        P(0, 155) = &H1B
        P(0, 156) = &H5F
        P(0, 157) = &H93
        P(0, 158) = &HA
        P(0, 159) = &HEF
        P(0, 160) = &H91
        P(0, 161) = &H85
        P(0, 162) = &H49
        P(0, 163) = &HEE
        P(0, 164) = &H2D
        P(0, 165) = &H4F
        P(0, 166) = &H8F
        P(0, 167) = &H3B
        P(0, 168) = &H47
        P(0, 169) = &H87
        P(0, 170) = &H6D
        P(0, 171) = &H46
        P(0, 172) = &HD6
        P(0, 173) = &H3E
        P(0, 174) = &H69
        P(0, 175) = &H64
        P(0, 176) = &H2A
        P(0, 177) = &HCE
        P(0, 178) = &HCB
        P(0, 179) = &H2F
        P(0, 180) = &HFC
        P(0, 181) = &H97
        P(0, 182) = &H5
        P(0, 183) = &H7A
        P(0, 184) = &HAC
        P(0, 185) = &H7F
        P(0, 186) = &HD5
        P(0, 187) = &H1A
        P(0, 188) = &H4B
        P(0, 189) = &HE
        P(0, 190) = &HA7
        P(0, 191) = &H5A
        P(0, 192) = &H28
        P(0, 193) = &H14
        P(0, 194) = &H3F
        P(0, 195) = &H29
        P(0, 196) = &H88
        P(0, 197) = &H3C
        P(0, 198) = &H4C
        P(0, 199) = &H2
        P(0, 200) = &HB8
        P(0, 201) = &HDA
        P(0, 202) = &HB0
        P(0, 203) = &H17
        P(0, 204) = &H55
        P(0, 205) = &H1F
        P(0, 206) = &H8A
        P(0, 207) = &H7D
        P(0, 208) = &H57
        P(0, 209) = &HC7
        P(0, 210) = &H8D
        P(0, 211) = &H74
        P(0, 212) = &HB7
        P(0, 213) = &HC4
        P(0, 214) = &H9F
        P(0, 215) = &H72
        P(0, 216) = &H7E
        P(0, 217) = &H15
        P(0, 218) = &H22
        P(0, 219) = &H12
        P(0, 220) = &H58
        P(0, 221) = &H7
        P(0, 222) = &H99
        P(0, 223) = &H34
        P(0, 224) = &H6E
        P(0, 225) = &H50
        P(0, 226) = &HDE
        P(0, 227) = &H68
        P(0, 228) = &H65
        P(0, 229) = &HBC
        P(0, 230) = &HDB
        P(0, 231) = &HF8
        P(0, 232) = &HC8
        P(0, 233) = &HA8
        P(0, 234) = &H2B
        P(0, 235) = &H40
        P(0, 236) = &HDC
        P(0, 237) = &HFE
        P(0, 238) = &H32
        P(0, 239) = &HA4
        P(0, 240) = &HCA
        P(0, 241) = &H10
        P(0, 242) = &H21
        P(0, 243) = &HF0
        P(0, 244) = &HD3
        P(0, 245) = &H5D
        P(0, 246) = &HF
        P(0, 247) = &H0
        P(0, 248) = &H6F
        P(0, 249) = &H9D
        P(0, 250) = &H36
        P(0, 251) = &H42
        P(0, 252) = &H4A
        P(0, 253) = &H5E
        P(0, 254) = &HC1
        P(0, 255) = &HE0

        'Initialize P(1,..) array
        P(1, 0) = &H75
        P(1, 1) = &HF3
        P(1, 2) = &HC6
        P(1, 3) = &HF4
        P(1, 4) = &HDB
        P(1, 5) = &H7B
        P(1, 6) = &HFB
        P(1, 7) = &HC8
        P(1, 8) = &H4A
        P(1, 9) = &HD3
        P(1, 10) = &HE6
        P(1, 11) = &H6B
        P(1, 12) = &H45
        P(1, 13) = &H7D
        P(1, 14) = &HE8
        P(1, 15) = &H4B
        P(1, 16) = &HD6
        P(1, 17) = &H32
        P(1, 18) = &HD8
        P(1, 19) = &HFD
        P(1, 20) = &H37
        P(1, 21) = &H71
        P(1, 22) = &HF1
        P(1, 23) = &HE1
        P(1, 24) = &H30
        P(1, 25) = &HF
        P(1, 26) = &HF8
        P(1, 27) = &H1B
        P(1, 28) = &H87
        P(1, 29) = &HFA
        P(1, 30) = &H6
        P(1, 31) = &H3F
        P(1, 32) = &H5E
        P(1, 33) = &HBA
        P(1, 34) = &HAE
        P(1, 35) = &H5B
        P(1, 36) = &H8A
        P(1, 37) = &H0
        P(1, 38) = &HBC
        P(1, 39) = &H9D
        P(1, 40) = &H6D
        P(1, 41) = &HC1
        P(1, 42) = &HB1
        P(1, 43) = &HE
        P(1, 44) = &H80
        P(1, 45) = &H5D
        P(1, 46) = &HD2
        P(1, 47) = &HD5
        P(1, 48) = &HA0
        P(1, 49) = &H84
        P(1, 50) = &H7
        P(1, 51) = &H14
        P(1, 52) = &HB5
        P(1, 53) = &H90
        P(1, 54) = &H2C
        P(1, 55) = &HA3
        P(1, 56) = &HB2
        P(1, 57) = &H73
        P(1, 58) = &H4C
        P(1, 59) = &H54
        P(1, 60) = &H92
        P(1, 61) = &H74
        P(1, 62) = &H36
        P(1, 63) = &H51
        P(1, 64) = &H38
        P(1, 65) = &HB0
        P(1, 66) = &HBD
        P(1, 67) = &H5A
        P(1, 68) = &HFC
        P(1, 69) = &H60
        P(1, 70) = &H62
        P(1, 71) = &H96
        P(1, 72) = &H6C
        P(1, 73) = &H42
        P(1, 74) = &HF7
        P(1, 75) = &H10
        P(1, 76) = &H7C
        P(1, 77) = &H28
        P(1, 78) = &H27
        P(1, 79) = &H8C
        P(1, 80) = &H13
        P(1, 81) = &H95
        P(1, 82) = &H9C
        P(1, 83) = &HC7
        P(1, 84) = &H24
        P(1, 85) = &H46
        P(1, 86) = &H3B
        P(1, 87) = &H70
        P(1, 88) = &HCA
        P(1, 89) = &HE3
        P(1, 90) = &H85
        P(1, 91) = &HCB
        P(1, 92) = &H11
        P(1, 93) = &HD0
        P(1, 94) = &H93
        P(1, 95) = &HB8
        P(1, 96) = &HA6
        P(1, 97) = &H83
        P(1, 98) = &H20
        P(1, 99) = &HFF
        P(1, 100) = &H9F
        P(1, 101) = &H77
        P(1, 102) = &HC3
        P(1, 103) = &HCC
        P(1, 104) = &H3
        P(1, 105) = &H6F
        P(1, 106) = &H8
        P(1, 107) = &HBF
        P(1, 108) = &H40
        P(1, 109) = &HE7
        P(1, 110) = &H2B
        P(1, 111) = &HE2
        P(1, 112) = &H79
        P(1, 113) = &HC
        P(1, 114) = &HAA
        P(1, 115) = &H82
        P(1, 116) = &H41
        P(1, 117) = &H3A
        P(1, 118) = &HEA
        P(1, 119) = &HB9
        P(1, 120) = &HE4
        P(1, 121) = &H9A
        P(1, 122) = &HA4
        P(1, 123) = &H97
        P(1, 124) = &H7E
        P(1, 125) = &HDA
        P(1, 126) = &H7A
        P(1, 127) = &H17
        P(1, 128) = &H66
        P(1, 129) = &H94
        P(1, 130) = &HA1
        P(1, 131) = &H1D
        P(1, 132) = &H3D
        P(1, 133) = &HF0
        P(1, 134) = &HDE
        P(1, 135) = &HB3
        P(1, 136) = &HB
        P(1, 137) = &H72
        P(1, 138) = &HA7
        P(1, 139) = &H1C
        P(1, 140) = &HEF
        P(1, 141) = &HD1
        P(1, 142) = &H53
        P(1, 143) = &H3E
        P(1, 144) = &H8F
        P(1, 145) = &H33
        P(1, 146) = &H26
        P(1, 147) = &H5F
        P(1, 148) = &HEC
        P(1, 149) = &H76
        P(1, 150) = &H2A
        P(1, 151) = &H49
        P(1, 152) = &H81
        P(1, 153) = &H88
        P(1, 154) = &HEE
        P(1, 155) = &H21
        P(1, 156) = &HC4
        P(1, 157) = &H1A
        P(1, 158) = &HEB
        P(1, 159) = &HD9
        P(1, 160) = &HC5
        P(1, 161) = &H39
        P(1, 162) = &H99
        P(1, 163) = &HCD
        P(1, 164) = &HAD
        P(1, 165) = &H31
        P(1, 166) = &H8B
        P(1, 167) = &H1
        P(1, 168) = &H18
        P(1, 169) = &H23
        P(1, 170) = &HDD
        P(1, 171) = &H1F
        P(1, 172) = &H4E
        P(1, 173) = &H2D
        P(1, 174) = &HF9
        P(1, 175) = &H48
        P(1, 176) = &H4F
        P(1, 177) = &HF2
        P(1, 178) = &H65
        P(1, 179) = &H8E
        P(1, 180) = &H78
        P(1, 181) = &H5C
        P(1, 182) = &H58
        P(1, 183) = &H19
        P(1, 184) = &H8D
        P(1, 185) = &HE5
        P(1, 186) = &H98
        P(1, 187) = &H57
        P(1, 188) = &H67
        P(1, 189) = &H7F
        P(1, 190) = &H5
        P(1, 191) = &H64
        P(1, 192) = &HAF
        P(1, 193) = &H63
        P(1, 194) = &HB6
        P(1, 195) = &HFE
        P(1, 196) = &HF5
        P(1, 197) = &HB7
        P(1, 198) = &H3C
        P(1, 199) = &HA5
        P(1, 200) = &HCE
        P(1, 201) = &HE9
        P(1, 202) = &H68
        P(1, 203) = &H44
        P(1, 204) = &HE0
        P(1, 205) = &H4D
        P(1, 206) = &H43
        P(1, 207) = &H69
        P(1, 208) = &H29
        P(1, 209) = &H2E
        P(1, 210) = &HAC
        P(1, 211) = &H15
        P(1, 212) = &H59
        P(1, 213) = &HA8
        P(1, 214) = &HA
        P(1, 215) = &H9E
        P(1, 216) = &H6E
        P(1, 217) = &H47
        P(1, 218) = &HDF
        P(1, 219) = &H34
        P(1, 220) = &H35
        P(1, 221) = &H6A
        P(1, 222) = &HCF
        P(1, 223) = &HDC
        P(1, 224) = &H22
        P(1, 225) = &HC9
        P(1, 226) = &HC0
        P(1, 227) = &H9B
        P(1, 228) = &H89
        P(1, 229) = &HD4
        P(1, 230) = &HED
        P(1, 231) = &HAB
        P(1, 232) = &H12
        P(1, 233) = &HA2
        P(1, 234) = &HD
        P(1, 235) = &H52
        P(1, 236) = &HBB
        P(1, 237) = &H2
        P(1, 238) = &H2F
        P(1, 239) = &HA9
        P(1, 240) = &HD7
        P(1, 241) = &H61
        P(1, 242) = &H1E
        P(1, 243) = &HB4
        P(1, 244) = &H50
        P(1, 245) = &H4
        P(1, 246) = &HF6
        P(1, 247) = &HC2
        P(1, 248) = &H16
        P(1, 249) = &H25
        P(1, 250) = &H86
        P(1, 251) = &H56
        P(1, 252) = &H55
        P(1, 253) = &H9
        P(1, 254) = &HBE
        P(1, 255) = &H91

        'Initialize the MDS array
        For i = 0 To 255
            j = P(0, i)
            m1(0) = j
            mX(0) = j Xor Encryption_Twofish_LFSR2(j)
            mY(0) = j Xor Encryption_Twofish_LFSR1(j) Xor Encryption_Twofish_LFSR2(j)

            j = P(1, i)
            m1(1) = j
            mX(1) = j Xor Encryption_Twofish_LFSR2(j)
            mY(1) = j Xor Encryption_Twofish_LFSR1(j) Xor Encryption_Twofish_LFSR2(j)

            MDS(0, i) = (m1(1) Or Encryption_Twofish_lBSL(mX(1), 8) Or Encryption_Twofish_lBSL(mY(1), 16) Or Encryption_Twofish_lBSL(mY(1), 24))
            MDS(1, i) = (mY(0) Or Encryption_Twofish_lBSL(mY(0), 8) Or Encryption_Twofish_lBSL(mX(0), 16) Or Encryption_Twofish_lBSL(m1(0), 24))
            MDS(2, i) = (mX(1) Or Encryption_Twofish_lBSL(mY(1), 8) Or Encryption_Twofish_lBSL(m1(1), 16) Or Encryption_Twofish_lBSL(mY(1), 24))
            MDS(3, i) = (mX(0) Or Encryption_Twofish_lBSL(m1(0), 8) Or Encryption_Twofish_lBSL(mY(0), 16) Or Encryption_Twofish_lBSL(mX(0), 24))
        Next

End Sub

Private Static Sub Encryption_Blowfish_DecryptBlock(Xl As Long, Xr As Long)

Dim i As Long
Dim j As Long
Dim Temp As Long

    Temp = Xr
    Xr = Xl Xor m_pBox(ROUNDS + 1)
    Xl = Temp Xor m_pBox(ROUNDS)

    j = ROUNDS - 2
    For i = 0 To (ROUNDS \ 2 - 1)
        Xl = Xl Xor Encryption_Blowfish_F(Xr)
        Xr = Xr Xor m_pBox(j + 1)
        Xr = Xr Xor Encryption_Blowfish_F(Xl)
        Xl = Xl Xor m_pBox(j)
        j = j - 2
    Next

End Sub

Public Sub Encryption_Blowfish_DecryptByte(ByteArray() As Byte, Optional Key As String)

Dim Offset As Long
Dim OrigLen As Long
Dim LeftWord As Long
Dim RightWord As Long
Dim CipherLen As Long
Dim CipherLeft As Long
Dim CipherRight As Long

'Set the new key if one was provided

    If (Len(Key) > 0) Then Encryption_Blowfish_SetKey Key

    'Get the size of the ciphertext
    CipherLen = UBound(ByteArray) + 1

    'Decrypt the data in 64-bit blocks
    For Offset = 0 To (CipherLen - 1) Step 8
        'Get the next block of ciphertext
        Call Encryption_Misc_GetWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_GetWord(RightWord, ByteArray(), Offset + 4)

        'Decrypt the block
        Call Encryption_Blowfish_DecryptBlock(LeftWord, RightWord)

        'XOR with the previous cipherblock
        LeftWord = LeftWord Xor CipherLeft
        RightWord = RightWord Xor CipherRight

        'Store the current ciphertext to use
        'XOR with the next block plaintext
        Call Encryption_Misc_GetWord(CipherLeft, ByteArray(), Offset)
        Call Encryption_Misc_GetWord(CipherRight, ByteArray(), Offset + 4)

        'Store the block
        Call Encryption_Misc_PutWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_PutWord(RightWord, ByteArray(), Offset + 4)
        
    Next

    'Get the size of the original array
    Call CopyMem(OrigLen, ByteArray(8), 4)

    'Make sure OrigLen is a reasonable value,
    'if we used the wrong key the next couple
    'of statements could be dangerous (GPF)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then
        Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Blowfish decryption")
    End If

    'Resize the bytearray to hold only the plaintext
    'and not the extra information added by the
    'encryption routine
    Call CopyMem(ByteArray(0), ByteArray(12), OrigLen)
    ReDim Preserve ByteArray(OrigLen - 1)

End Sub

Public Sub Encryption_Blowfish_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_Blowfish_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Blowfish_DecryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the string to a bytearray

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the array
    Call Encryption_Blowfish_DecryptByte(ByteArray(), Key)

    'Return the encrypted data as a string
    Encryption_Blowfish_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Static Sub Encryption_Blowfish_EncryptBlock(Xl As Long, Xr As Long)

Dim i As Long
Dim j As Long
Dim Temp As Long

    j = 0
    For i = 0 To (ROUNDS \ 2 - 1)
        Xl = Xl Xor m_pBox(j)
        Xr = Xr Xor Encryption_Blowfish_F(Xl)
        Xr = Xr Xor m_pBox(j + 1)
        Xl = Xl Xor Encryption_Blowfish_F(Xr)
        j = j + 2
    Next

    Temp = Xr
    Xr = Xl Xor m_pBox(ROUNDS)
    Xl = Temp Xor m_pBox(ROUNDS + 1)

End Sub

Public Sub Encryption_Blowfish_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim Offset As Long
Dim OrigLen As Long
Dim LeftWord As Long
Dim RightWord As Long
Dim CipherLen As Long
Dim CipherLeft As Long
Dim CipherRight As Long

'Set the new key if one was provided

    If (Len(Key) > 0) Then Encryption_Blowfish_SetKey Key

    'Get the size of the original array
    OrigLen = UBound(ByteArray) + 1

    'First we add 12 bytes (4 bytes for the
    'length and 8 bytes for the seed values
    'for the CBC routine), and the ciphertext
    'must be a multiple of 8 bytes
    CipherLen = OrigLen + 12
    If (CipherLen Mod 8 <> 0) Then
        CipherLen = CipherLen + 8 - (CipherLen Mod 8)
    End If
    ReDim Preserve ByteArray(CipherLen - 1)
    Call CopyMem(ByteArray(12), ByteArray(0), OrigLen)

    'Store the length descriptor in bytes [9-12]
    Call CopyMem(ByteArray(8), OrigLen, 4)

    'Store a block of random data in bytes [1-8],
    'these work as seed values for the CBC routine
    'and is used to produce different ciphertext
    'even when encrypting the same data with the
    'same key)
    Call Randomize
    Call CopyMem(ByteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(ByteArray(4), CLng(2147483647 * Rnd), 4)

    'Encrypt the data in 64-bit blocks
    For Offset = 0 To (CipherLen - 1) Step 8
        'Get the next block of plaintext
        Call Encryption_Misc_GetWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_GetWord(RightWord, ByteArray(), Offset + 4)

        'XOR the plaintext with the previous
        'ciphertext (CBC, Cipher-Block Chaining)
        LeftWord = LeftWord Xor CipherLeft
        RightWord = RightWord Xor CipherRight

        'Encrypt the block
        Call Encryption_Blowfish_EncryptBlock(LeftWord, RightWord)

        'Store the block
        Call Encryption_Misc_PutWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_PutWord(RightWord, ByteArray(), Offset + 4)

        'Store the cipherblock (for CBC)
        CipherLeft = LeftWord
        CipherRight = RightWord

    Next

End Sub

Public Sub Encryption_Blowfish_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_Blowfish_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Blowfish_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the string to a bytearray

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the array
    Call Encryption_Blowfish_EncryptByte(ByteArray(), Key)

    'Return the encrypted data as a string
    Encryption_Blowfish_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Static Function Encryption_Blowfish_F(ByVal X As Long) As Long

Dim xb(0 To 3) As Byte

    Call CopyMem(xb(0), X, 4)
    If (m_RunningCompiled) Then
        Encryption_Blowfish_F = (((m_sBoxBF(0, xb(3)) + m_sBoxBF(1, xb(2))) Xor m_sBoxBF(2, xb(1))) + m_sBoxBF(3, xb(0)))
    Else
        Encryption_Blowfish_F = Encryption_Misc_UnsignedAdd((Encryption_Misc_UnsignedAdd(m_sBoxBF(0, xb(3)), m_sBoxBF(1, xb(2))) Xor m_sBoxBF(2, xb(1))), m_sBoxBF(3, xb(0)))
    End If

End Function

Private Sub Encryption_Blowfish_Init()

'We need to check if we are running in compiled
'(EXE) mode or in the IDE, this will allow us to
'use optimized code with unsigned integers in
'compiled mode without any overflow errors when
'running the code in the IDE

    On Local Error Resume Next
        m_RunningCompiled = ((2147483647 + 1) < 0)

        'Initialize p-boxes
        m_pBox(0) = &H243F6A88
        m_pBox(1) = &H85A308D3
        m_pBox(2) = &H13198A2E
        m_pBox(3) = &H3707344
        m_pBox(4) = &HA4093822
        m_pBox(5) = &H299F31D0
        m_pBox(6) = &H82EFA98
        m_pBox(7) = &HEC4E6C89
        m_pBox(8) = &H452821E6
        m_pBox(9) = &H38D01377
        m_pBox(10) = &HBE5466CF
        m_pBox(11) = &H34E90C6C
        m_pBox(12) = &HC0AC29B7
        m_pBox(13) = &HC97C50DD
        m_pBox(14) = &H3F84D5B5
        m_pBox(15) = &HB5470917
        m_pBox(16) = &H9216D5D9
        m_pBox(17) = &H8979FB1B

        'Initialize s-boxes
        m_sBoxBF(0, 0) = &HD1310BA6
        m_sBoxBF(1, 0) = &H98DFB5AC
        m_sBoxBF(2, 0) = &H2FFD72DB
        m_sBoxBF(3, 0) = &HD01ADFB7
        m_sBoxBF(0, 1) = &HB8E1AFED
        m_sBoxBF(1, 1) = &H6A267E96
        m_sBoxBF(2, 1) = &HBA7C9045
        m_sBoxBF(3, 1) = &HF12C7F99
        m_sBoxBF(0, 2) = &H24A19947
        m_sBoxBF(1, 2) = &HB3916CF7
        m_sBoxBF(2, 2) = &H801F2E2
        m_sBoxBF(3, 2) = &H858EFC16
        m_sBoxBF(0, 3) = &H636920D8
        m_sBoxBF(1, 3) = &H71574E69
        m_sBoxBF(2, 3) = &HA458FEA3
        m_sBoxBF(3, 3) = &HF4933D7E
        m_sBoxBF(0, 4) = &HD95748F
        m_sBoxBF(1, 4) = &H728EB658
        m_sBoxBF(2, 4) = &H718BCD58
        m_sBoxBF(3, 4) = &H82154AEE
        m_sBoxBF(0, 5) = &H7B54A41D
        m_sBoxBF(1, 5) = &HC25A59B5
        m_sBoxBF(2, 5) = &H9C30D539
        m_sBoxBF(3, 5) = &H2AF26013
        m_sBoxBF(0, 6) = &HC5D1B023
        m_sBoxBF(1, 6) = &H286085F0
        m_sBoxBF(2, 6) = &HCA417918
        m_sBoxBF(3, 6) = &HB8DB38EF
        m_sBoxBF(0, 7) = &H8E79DCB0
        m_sBoxBF(1, 7) = &H603A180E
        m_sBoxBF(2, 7) = &H6C9E0E8B
        m_sBoxBF(3, 7) = &HB01E8A3E
        m_sBoxBF(0, 8) = &HD71577C1
        m_sBoxBF(1, 8) = &HBD314B27
        m_sBoxBF(2, 8) = &H78AF2FDA
        m_sBoxBF(3, 8) = &H55605C60
        m_sBoxBF(0, 9) = &HE65525F3
        m_sBoxBF(1, 9) = &HAA55AB94
        m_sBoxBF(2, 9) = &H57489862
        m_sBoxBF(3, 9) = &H63E81440
        m_sBoxBF(0, 10) = &H55CA396A
        m_sBoxBF(1, 10) = &H2AAB10B6
        m_sBoxBF(2, 10) = &HB4CC5C34
        m_sBoxBF(3, 10) = &H1141E8CE
        m_sBoxBF(0, 11) = &HA15486AF
        m_sBoxBF(1, 11) = &H7C72E993
        m_sBoxBF(2, 11) = &HB3EE1411
        m_sBoxBF(3, 11) = &H636FBC2A
        m_sBoxBF(0, 12) = &H2BA9C55D
        m_sBoxBF(1, 12) = &H741831F6
        m_sBoxBF(2, 12) = &HCE5C3E16
        m_sBoxBF(3, 12) = &H9B87931E
        m_sBoxBF(0, 13) = &HAFD6BA33
        m_sBoxBF(1, 13) = &H6C24CF5C
        m_sBoxBF(2, 13) = &H7A325381
        m_sBoxBF(3, 13) = &H28958677
        m_sBoxBF(0, 14) = &H3B8F4898
        m_sBoxBF(1, 14) = &H6B4BB9AF
        m_sBoxBF(2, 14) = &HC4BFE81B
        m_sBoxBF(3, 14) = &H66282193
        m_sBoxBF(0, 15) = &H61D809CC
        m_sBoxBF(1, 15) = &HFB21A991
        m_sBoxBF(2, 15) = &H487CAC60
        m_sBoxBF(3, 15) = &H5DEC8032
        m_sBoxBF(0, 16) = &HEF845D5D
        m_sBoxBF(1, 16) = &HE98575B1
        m_sBoxBF(2, 16) = &HDC262302
        m_sBoxBF(3, 16) = &HEB651B88
        m_sBoxBF(0, 17) = &H23893E81
        m_sBoxBF(1, 17) = &HD396ACC5
        m_sBoxBF(2, 17) = &HF6D6FF3
        m_sBoxBF(3, 17) = &H83F44239
        m_sBoxBF(0, 18) = &H2E0B4482
        m_sBoxBF(1, 18) = &HA4842004
        m_sBoxBF(2, 18) = &H69C8F04A
        m_sBoxBF(3, 18) = &H9E1F9B5E
        m_sBoxBF(0, 19) = &H21C66842
        m_sBoxBF(1, 19) = &HF6E96C9A
        m_sBoxBF(2, 19) = &H670C9C61
        m_sBoxBF(3, 19) = &HABD388F0
        m_sBoxBF(0, 20) = &H6A51A0D2
        m_sBoxBF(1, 20) = &HD8542F68
        m_sBoxBF(2, 20) = &H960FA728
        m_sBoxBF(3, 20) = &HAB5133A3
        m_sBoxBF(0, 21) = &H6EEF0B6C
        m_sBoxBF(1, 21) = &H137A3BE4
        m_sBoxBF(2, 21) = &HBA3BF050
        m_sBoxBF(3, 21) = &H7EFB2A98
        m_sBoxBF(0, 22) = &HA1F1651D
        m_sBoxBF(1, 22) = &H39AF0176
        m_sBoxBF(2, 22) = &H66CA593E
        m_sBoxBF(3, 22) = &H82430E88
        m_sBoxBF(0, 23) = &H8CEE8619
        m_sBoxBF(1, 23) = &H456F9FB4
        m_sBoxBF(2, 23) = &H7D84A5C3
        m_sBoxBF(3, 23) = &H3B8B5EBE
        m_sBoxBF(0, 24) = &HE06F75D8
        m_sBoxBF(1, 24) = &H85C12073
        m_sBoxBF(2, 24) = &H401A449F
        m_sBoxBF(3, 24) = &H56C16AA6
        m_sBoxBF(0, 25) = &H4ED3AA62
        m_sBoxBF(1, 25) = &H363F7706
        m_sBoxBF(2, 25) = &H1BFEDF72
        m_sBoxBF(3, 25) = &H429B023D
        m_sBoxBF(0, 26) = &H37D0D724
        m_sBoxBF(1, 26) = &HD00A1248
        m_sBoxBF(2, 26) = &HDB0FEAD3
        m_sBoxBF(3, 26) = &H49F1C09B
        m_sBoxBF(0, 27) = &H75372C9
        m_sBoxBF(1, 27) = &H80991B7B
        m_sBoxBF(2, 27) = &H25D479D8
        m_sBoxBF(3, 27) = &HF6E8DEF7
        m_sBoxBF(0, 28) = &HE3FE501A
        m_sBoxBF(1, 28) = &HB6794C3B
        m_sBoxBF(2, 28) = &H976CE0BD
        m_sBoxBF(3, 28) = &H4C006BA
        m_sBoxBF(0, 29) = &HC1A94FB6
        m_sBoxBF(1, 29) = &H409F60C4
        m_sBoxBF(2, 29) = &H5E5C9EC2
        m_sBoxBF(3, 29) = &H196A2463
        m_sBoxBF(0, 30) = &H68FB6FAF
        m_sBoxBF(1, 30) = &H3E6C53B5
        m_sBoxBF(2, 30) = &H1339B2EB
        m_sBoxBF(3, 30) = &H3B52EC6F
        m_sBoxBF(0, 31) = &H6DFC511F
        m_sBoxBF(1, 31) = &H9B30952C
        m_sBoxBF(2, 31) = &HCC814544
        m_sBoxBF(3, 31) = &HAF5EBD09
        m_sBoxBF(0, 32) = &HBEE3D004
        m_sBoxBF(1, 32) = &HDE334AFD
        m_sBoxBF(2, 32) = &H660F2807
        m_sBoxBF(3, 32) = &H192E4BB3
        m_sBoxBF(0, 33) = &HC0CBA857
        m_sBoxBF(1, 33) = &H45C8740F
        m_sBoxBF(2, 33) = &HD20B5F39
        m_sBoxBF(3, 33) = &HB9D3FBDB
        m_sBoxBF(0, 34) = &H5579C0BD
        m_sBoxBF(1, 34) = &H1A60320A
        m_sBoxBF(2, 34) = &HD6A100C6
        m_sBoxBF(3, 34) = &H402C7279
        m_sBoxBF(0, 35) = &H679F25FE
        m_sBoxBF(1, 35) = &HFB1FA3CC
        m_sBoxBF(2, 35) = &H8EA5E9F8
        m_sBoxBF(3, 35) = &HDB3222F8
        m_sBoxBF(0, 36) = &H3C7516DF
        m_sBoxBF(1, 36) = &HFD616B15
        m_sBoxBF(2, 36) = &H2F501EC8
        m_sBoxBF(3, 36) = &HAD0552AB
        m_sBoxBF(0, 37) = &H323DB5FA
        m_sBoxBF(1, 37) = &HFD238760
        m_sBoxBF(2, 37) = &H53317B48
        m_sBoxBF(3, 37) = &H3E00DF82
        m_sBoxBF(0, 38) = &H9E5C57BB
        m_sBoxBF(1, 38) = &HCA6F8CA0
        m_sBoxBF(2, 38) = &H1A87562E
        m_sBoxBF(3, 38) = &HDF1769DB
        m_sBoxBF(0, 39) = &HD542A8F6
        m_sBoxBF(1, 39) = &H287EFFC3
        m_sBoxBF(2, 39) = &HAC6732C6
        m_sBoxBF(3, 39) = &H8C4F5573
        m_sBoxBF(0, 40) = &H695B27B0
        m_sBoxBF(1, 40) = &HBBCA58C8
        m_sBoxBF(2, 40) = &HE1FFA35D
        m_sBoxBF(3, 40) = &HB8F011A0
        m_sBoxBF(0, 41) = &H10FA3D98
        m_sBoxBF(1, 41) = &HFD2183B8
        m_sBoxBF(2, 41) = &H4AFCB56C
        m_sBoxBF(3, 41) = &H2DD1D35B
        m_sBoxBF(0, 42) = &H9A53E479
        m_sBoxBF(1, 42) = &HB6F84565
        m_sBoxBF(2, 42) = &HD28E49BC
        m_sBoxBF(3, 42) = &H4BFB9790
        m_sBoxBF(0, 43) = &HE1DDF2DA
        m_sBoxBF(1, 43) = &HA4CB7E33
        m_sBoxBF(2, 43) = &H62FB1341
        m_sBoxBF(3, 43) = &HCEE4C6E8
        m_sBoxBF(0, 44) = &HEF20CADA
        m_sBoxBF(1, 44) = &H36774C01
        m_sBoxBF(2, 44) = &HD07E9EFE
        m_sBoxBF(3, 44) = &H2BF11FB4
        m_sBoxBF(0, 45) = &H95DBDA4D
        m_sBoxBF(1, 45) = &HAE909198
        m_sBoxBF(2, 45) = &HEAAD8E71
        m_sBoxBF(3, 45) = &H6B93D5A0
        m_sBoxBF(0, 46) = &HD08ED1D0
        m_sBoxBF(1, 46) = &HAFC725E0
        m_sBoxBF(2, 46) = &H8E3C5B2F
        m_sBoxBF(3, 46) = &H8E7594B7
        m_sBoxBF(0, 47) = &H8FF6E2FB
        m_sBoxBF(1, 47) = &HF2122B64
        m_sBoxBF(2, 47) = &H8888B812
        m_sBoxBF(3, 47) = &H900DF01C
        m_sBoxBF(0, 48) = &H4FAD5EA0
        m_sBoxBF(1, 48) = &H688FC31C
        m_sBoxBF(2, 48) = &HD1CFF191
        m_sBoxBF(3, 48) = &HB3A8C1AD
        m_sBoxBF(0, 49) = &H2F2F2218
        m_sBoxBF(1, 49) = &HBE0E1777
        m_sBoxBF(2, 49) = &HEA752DFE
        m_sBoxBF(3, 49) = &H8B021FA1
        m_sBoxBF(0, 50) = &HE5A0CC0F
        m_sBoxBF(1, 50) = &HB56F74E8
        m_sBoxBF(2, 50) = &H18ACF3D6
        m_sBoxBF(3, 50) = &HCE89E299
        m_sBoxBF(0, 51) = &HB4A84FE0
        m_sBoxBF(1, 51) = &HFD13E0B7
        m_sBoxBF(2, 51) = &H7CC43B81
        m_sBoxBF(3, 51) = &HD2ADA8D9
        m_sBoxBF(0, 52) = &H165FA266
        m_sBoxBF(1, 52) = &H80957705
        m_sBoxBF(2, 52) = &H93CC7314
        m_sBoxBF(3, 52) = &H211A1477
        m_sBoxBF(0, 53) = &HE6AD2065
        m_sBoxBF(1, 53) = &H77B5FA86
        m_sBoxBF(2, 53) = &HC75442F5
        m_sBoxBF(3, 53) = &HFB9D35CF
        m_sBoxBF(0, 54) = &HEBCDAF0C
        m_sBoxBF(1, 54) = &H7B3E89A0
        m_sBoxBF(2, 54) = &HD6411BD3
        m_sBoxBF(3, 54) = &HAE1E7E49
        m_sBoxBF(0, 55) = &H250E2D
        m_sBoxBF(1, 55) = &H2071B35E
        m_sBoxBF(2, 55) = &H226800BB
        m_sBoxBF(3, 55) = &H57B8E0AF
        m_sBoxBF(0, 56) = &H2464369B
        m_sBoxBF(1, 56) = &HF009B91E
        m_sBoxBF(2, 56) = &H5563911D
        m_sBoxBF(3, 56) = &H59DFA6AA
        m_sBoxBF(0, 57) = &H78C14389
        m_sBoxBF(1, 57) = &HD95A537F
        m_sBoxBF(2, 57) = &H207D5BA2
        m_sBoxBF(3, 57) = &H2E5B9C5
        m_sBoxBF(0, 58) = &H83260376
        m_sBoxBF(1, 58) = &H6295CFA9
        m_sBoxBF(2, 58) = &H11C81968
        m_sBoxBF(3, 58) = &H4E734A41
        m_sBoxBF(0, 59) = &HB3472DCA
        m_sBoxBF(1, 59) = &H7B14A94A
        m_sBoxBF(2, 59) = &H1B510052
        m_sBoxBF(3, 59) = &H9A532915
        m_sBoxBF(0, 60) = &HD60F573F
        m_sBoxBF(1, 60) = &HBC9BC6E4
        m_sBoxBF(2, 60) = &H2B60A476
        m_sBoxBF(3, 60) = &H81E67400
        m_sBoxBF(0, 61) = &H8BA6FB5
        m_sBoxBF(1, 61) = &H571BE91F
        m_sBoxBF(2, 61) = &HF296EC6B
        m_sBoxBF(3, 61) = &H2A0DD915
        m_sBoxBF(0, 62) = &HB6636521
        m_sBoxBF(1, 62) = &HE7B9F9B6
        m_sBoxBF(2, 62) = &HFF34052E
        m_sBoxBF(3, 62) = &HC5855664
        m_sBoxBF(0, 63) = &H53B02D5D
        m_sBoxBF(1, 63) = &HA99F8FA1
        m_sBoxBF(2, 63) = &H8BA4799
        m_sBoxBF(3, 63) = &H6E85076A
        m_sBoxBF(0, 64) = &H4B7A70E9
        m_sBoxBF(1, 64) = &HB5B32944
        m_sBoxBF(2, 64) = &HDB75092E
        m_sBoxBF(3, 64) = &HC4192623
        m_sBoxBF(0, 65) = &HAD6EA6B0
        m_sBoxBF(1, 65) = &H49A7DF7D
        m_sBoxBF(2, 65) = &H9CEE60B8
        m_sBoxBF(3, 65) = &H8FEDB266
        m_sBoxBF(0, 66) = &HECAA8C71
        m_sBoxBF(1, 66) = &H699A17FF
        m_sBoxBF(2, 66) = &H5664526C
        m_sBoxBF(3, 66) = &HC2B19EE1
        m_sBoxBF(0, 67) = &H193602A5
        m_sBoxBF(1, 67) = &H75094C29
        m_sBoxBF(2, 67) = &HA0591340
        m_sBoxBF(3, 67) = &HE4183A3E
        m_sBoxBF(0, 68) = &H3F54989A
        m_sBoxBF(1, 68) = &H5B429D65
        m_sBoxBF(2, 68) = &H6B8FE4D6
        m_sBoxBF(3, 68) = &H99F73FD6
        m_sBoxBF(0, 69) = &HA1D29C07
        m_sBoxBF(1, 69) = &HEFE830F5
        m_sBoxBF(2, 69) = &H4D2D38E6
        m_sBoxBF(3, 69) = &HF0255DC1
        m_sBoxBF(0, 70) = &H4CDD2086
        m_sBoxBF(1, 70) = &H8470EB26
        m_sBoxBF(2, 70) = &H6382E9C6
        m_sBoxBF(3, 70) = &H21ECC5E
        m_sBoxBF(0, 71) = &H9686B3F
        m_sBoxBF(1, 71) = &H3EBAEFC9
        m_sBoxBF(2, 71) = &H3C971814
        m_sBoxBF(3, 71) = &H6B6A70A1
        m_sBoxBF(0, 72) = &H687F3584
        m_sBoxBF(1, 72) = &H52A0E286
        m_sBoxBF(2, 72) = &HB79C5305
        m_sBoxBF(3, 72) = &HAA500737
        m_sBoxBF(0, 73) = &H3E07841C
        m_sBoxBF(1, 73) = &H7FDEAE5C
        m_sBoxBF(2, 73) = &H8E7D44EC
        m_sBoxBF(3, 73) = &H5716F2B8
        m_sBoxBF(0, 74) = &HB03ADA37
        m_sBoxBF(1, 74) = &HF0500C0D
        m_sBoxBF(2, 74) = &HF01C1F04
        m_sBoxBF(3, 74) = &H200B3FF
        m_sBoxBF(0, 75) = &HAE0CF51A
        m_sBoxBF(1, 75) = &H3CB574B2
        m_sBoxBF(2, 75) = &H25837A58
        m_sBoxBF(3, 75) = &HDC0921BD
        m_sBoxBF(0, 76) = &HD19113F9
        m_sBoxBF(1, 76) = &H7CA92FF6
        m_sBoxBF(2, 76) = &H94324773
        m_sBoxBF(3, 76) = &H22F54701
        m_sBoxBF(0, 77) = &H3AE5E581
        m_sBoxBF(1, 77) = &H37C2DADC
        m_sBoxBF(2, 77) = &HC8B57634
        m_sBoxBF(3, 77) = &H9AF3DDA7
        m_sBoxBF(0, 78) = &HA9446146
        m_sBoxBF(1, 78) = &HFD0030E
        m_sBoxBF(2, 78) = &HECC8C73E
        m_sBoxBF(3, 78) = &HA4751E41
        m_sBoxBF(0, 79) = &HE238CD99
        m_sBoxBF(1, 79) = &H3BEA0E2F
        m_sBoxBF(2, 79) = &H3280BBA1
        m_sBoxBF(3, 79) = &H183EB331
        m_sBoxBF(0, 80) = &H4E548B38
        m_sBoxBF(1, 80) = &H4F6DB908
        m_sBoxBF(2, 80) = &H6F420D03
        m_sBoxBF(3, 80) = &HF60A04BF
        m_sBoxBF(0, 81) = &H2CB81290
        m_sBoxBF(1, 81) = &H24977C79
        m_sBoxBF(2, 81) = &H5679B072
        m_sBoxBF(3, 81) = &HBCAF89AF
        m_sBoxBF(0, 82) = &HDE9A771F
        m_sBoxBF(1, 82) = &HD9930810
        m_sBoxBF(2, 82) = &HB38BAE12
        m_sBoxBF(3, 82) = &HDCCF3F2E
        m_sBoxBF(0, 83) = &H5512721F
        m_sBoxBF(1, 83) = &H2E6B7124
        m_sBoxBF(2, 83) = &H501ADDE6
        m_sBoxBF(3, 83) = &H9F84CD87
        m_sBoxBF(0, 84) = &H7A584718
        m_sBoxBF(1, 84) = &H7408DA17
        m_sBoxBF(2, 84) = &HBC9F9ABC
        m_sBoxBF(3, 84) = &HE94B7D8C
        m_sBoxBF(0, 85) = &HEC7AEC3A
        m_sBoxBF(1, 85) = &HDB851DFA
        m_sBoxBF(2, 85) = &H63094366
        m_sBoxBF(3, 85) = &HC464C3D2
        m_sBoxBF(0, 86) = &HEF1C1847
        m_sBoxBF(1, 86) = &H3215D908
        m_sBoxBF(2, 86) = &HDD433B37
        m_sBoxBF(3, 86) = &H24C2BA16
        m_sBoxBF(0, 87) = &H12A14D43
        m_sBoxBF(1, 87) = &H2A65C451
        m_sBoxBF(2, 87) = &H50940002
        m_sBoxBF(3, 87) = &H133AE4DD
        m_sBoxBF(0, 88) = &H71DFF89E
        m_sBoxBF(1, 88) = &H10314E55
        m_sBoxBF(2, 88) = &H81AC77D6
        m_sBoxBF(3, 88) = &H5F11199B
        m_sBoxBF(0, 89) = &H43556F1
        m_sBoxBF(1, 89) = &HD7A3C76B
        m_sBoxBF(2, 89) = &H3C11183B
        m_sBoxBF(3, 89) = &H5924A509
        m_sBoxBF(0, 90) = &HF28FE6ED
        m_sBoxBF(1, 90) = &H97F1FBFA
        m_sBoxBF(2, 90) = &H9EBABF2C
        m_sBoxBF(3, 90) = &H1E153C6E
        m_sBoxBF(0, 91) = &H86E34570
        m_sBoxBF(1, 91) = &HEAE96FB1
        m_sBoxBF(2, 91) = &H860E5E0A
        m_sBoxBF(3, 91) = &H5A3E2AB3
        m_sBoxBF(0, 92) = &H771FE71C
        m_sBoxBF(1, 92) = &H4E3D06FA
        m_sBoxBF(2, 92) = &H2965DCB9
        m_sBoxBF(3, 92) = &H99E71D0F
        m_sBoxBF(0, 93) = &H803E89D6
        m_sBoxBF(1, 93) = &H5266C825
        m_sBoxBF(2, 93) = &H2E4CC978
        m_sBoxBF(3, 93) = &H9C10B36A
        m_sBoxBF(0, 94) = &HC6150EBA
        m_sBoxBF(1, 94) = &H94E2EA78
        m_sBoxBF(2, 94) = &HA5FC3C53
        m_sBoxBF(3, 94) = &H1E0A2DF4
        m_sBoxBF(0, 95) = &HF2F74EA7
        m_sBoxBF(1, 95) = &H361D2B3D
        m_sBoxBF(2, 95) = &H1939260F
        m_sBoxBF(3, 95) = &H19C27960
        m_sBoxBF(0, 96) = &H5223A708
        m_sBoxBF(1, 96) = &HF71312B6
        m_sBoxBF(2, 96) = &HEBADFE6E
        m_sBoxBF(3, 96) = &HEAC31F66
        m_sBoxBF(0, 97) = &HE3BC4595
        m_sBoxBF(1, 97) = &HA67BC883
        m_sBoxBF(2, 97) = &HB17F37D1
        m_sBoxBF(3, 97) = &H18CFF28
        m_sBoxBF(0, 98) = &HC332DDEF
        m_sBoxBF(1, 98) = &HBE6C5AA5
        m_sBoxBF(2, 98) = &H65582185
        m_sBoxBF(3, 98) = &H68AB9802
        m_sBoxBF(0, 99) = &HEECEA50F
        m_sBoxBF(1, 99) = &HDB2F953B
        m_sBoxBF(2, 99) = &H2AEF7DAD
        m_sBoxBF(3, 99) = &H5B6E2F84
        m_sBoxBF(0, 100) = &H1521B628
        m_sBoxBF(1, 100) = &H29076170
        m_sBoxBF(2, 100) = &HECDD4775
        m_sBoxBF(3, 100) = &H619F1510
        m_sBoxBF(0, 101) = &H13CCA830
        m_sBoxBF(1, 101) = &HEB61BD96
        m_sBoxBF(2, 101) = &H334FE1E
        m_sBoxBF(3, 101) = &HAA0363CF
        m_sBoxBF(0, 102) = &HB5735C90
        m_sBoxBF(1, 102) = &H4C70A239
        m_sBoxBF(2, 102) = &HD59E9E0B
        m_sBoxBF(3, 102) = &HCBAADE14
        m_sBoxBF(0, 103) = &HEECC86BC
        m_sBoxBF(1, 103) = &H60622CA7
        m_sBoxBF(2, 103) = &H9CAB5CAB
        m_sBoxBF(3, 103) = &HB2F3846E
        m_sBoxBF(0, 104) = &H648B1EAF
        m_sBoxBF(1, 104) = &H19BDF0CA
        m_sBoxBF(2, 104) = &HA02369B9
        m_sBoxBF(3, 104) = &H655ABB50
        m_sBoxBF(0, 105) = &H40685A32
        m_sBoxBF(1, 105) = &H3C2AB4B3
        m_sBoxBF(2, 105) = &H319EE9D5
        m_sBoxBF(3, 105) = &HC021B8F7
        m_sBoxBF(0, 106) = &H9B540B19
        m_sBoxBF(1, 106) = &H875FA099
        m_sBoxBF(2, 106) = &H95F7997E
        m_sBoxBF(3, 106) = &H623D7DA8
        m_sBoxBF(0, 107) = &HF837889A
        m_sBoxBF(1, 107) = &H97E32D77
        m_sBoxBF(2, 107) = &H11ED935F
        m_sBoxBF(3, 107) = &H16681281
        m_sBoxBF(0, 108) = &HE358829
        m_sBoxBF(1, 108) = &HC7E61FD6
        m_sBoxBF(2, 108) = &H96DEDFA1
        m_sBoxBF(3, 108) = &H7858BA99
        m_sBoxBF(0, 109) = &H57F584A5
        m_sBoxBF(1, 109) = &H1B227263
        m_sBoxBF(2, 109) = &H9B83C3FF
        m_sBoxBF(3, 109) = &H1AC24696
        m_sBoxBF(0, 110) = &HCDB30AEB
        m_sBoxBF(1, 110) = &H532E3054
        m_sBoxBF(2, 110) = &H8FD948E4
        m_sBoxBF(3, 110) = &H6DBC3128
        m_sBoxBF(0, 111) = &H58EBF2EF
        m_sBoxBF(1, 111) = &H34C6FFEA
        m_sBoxBF(2, 111) = &HFE28ED61
        m_sBoxBF(3, 111) = &HEE7C3C73
        m_sBoxBF(0, 112) = &H5D4A14D9
        m_sBoxBF(1, 112) = &HE864B7E3
        m_sBoxBF(2, 112) = &H42105D14
        m_sBoxBF(3, 112) = &H203E13E0
        m_sBoxBF(0, 113) = &H45EEE2B6
        m_sBoxBF(1, 113) = &HA3AAABEA
        m_sBoxBF(2, 113) = &HDB6C4F15
        m_sBoxBF(3, 113) = &HFACB4FD0
        m_sBoxBF(0, 114) = &HC742F442
        m_sBoxBF(1, 114) = &HEF6ABBB5
        m_sBoxBF(2, 114) = &H654F3B1D
        m_sBoxBF(3, 114) = &H41CD2105
        m_sBoxBF(0, 115) = &HD81E799E
        m_sBoxBF(1, 115) = &H86854DC7
        m_sBoxBF(2, 115) = &HE44B476A
        m_sBoxBF(3, 115) = &H3D816250
        m_sBoxBF(0, 116) = &HCF62A1F2
        m_sBoxBF(1, 116) = &H5B8D2646
        m_sBoxBF(2, 116) = &HFC8883A0
        m_sBoxBF(3, 116) = &HC1C7B6A3
        m_sBoxBF(0, 117) = &H7F1524C3
        m_sBoxBF(1, 117) = &H69CB7492
        m_sBoxBF(2, 117) = &H47848A0B
        m_sBoxBF(3, 117) = &H5692B285
        m_sBoxBF(0, 118) = &H95BBF00
        m_sBoxBF(1, 118) = &HAD19489D
        m_sBoxBF(2, 118) = &H1462B174
        m_sBoxBF(3, 118) = &H23820E00
        m_sBoxBF(0, 119) = &H58428D2A
        m_sBoxBF(1, 119) = &HC55F5EA
        m_sBoxBF(2, 119) = &H1DADF43E
        m_sBoxBF(3, 119) = &H233F7061
        m_sBoxBF(0, 120) = &H3372F092
        m_sBoxBF(1, 120) = &H8D937E41
        m_sBoxBF(2, 120) = &HD65FECF1
        m_sBoxBF(3, 120) = &H6C223BDB
        m_sBoxBF(0, 121) = &H7CDE3759
        m_sBoxBF(1, 121) = &HCBEE7460
        m_sBoxBF(2, 121) = &H4085F2A7
        m_sBoxBF(3, 121) = &HCE77326E
        m_sBoxBF(0, 122) = &HA6078084
        m_sBoxBF(1, 122) = &H19F8509E
        m_sBoxBF(2, 122) = &HE8EFD855
        m_sBoxBF(3, 122) = &H61D99735
        m_sBoxBF(0, 123) = &HA969A7AA
        m_sBoxBF(1, 123) = &HC50C06C2
        m_sBoxBF(2, 123) = &H5A04ABFC
        m_sBoxBF(3, 123) = &H800BCADC
        m_sBoxBF(0, 124) = &H9E447A2E
        m_sBoxBF(1, 124) = &HC3453484
        m_sBoxBF(2, 124) = &HFDD56705
        m_sBoxBF(3, 124) = &HE1E9EC9
        m_sBoxBF(0, 125) = &HDB73DBD3
        m_sBoxBF(1, 125) = &H105588CD
        m_sBoxBF(2, 125) = &H675FDA79
        m_sBoxBF(3, 125) = &HE3674340
        m_sBoxBF(0, 126) = &HC5C43465
        m_sBoxBF(1, 126) = &H713E38D8
        m_sBoxBF(2, 126) = &H3D28F89E
        m_sBoxBF(3, 126) = &HF16DFF20
        m_sBoxBF(0, 127) = &H153E21E7
        m_sBoxBF(1, 127) = &H8FB03D4A
        m_sBoxBF(2, 127) = &HE6E39F2B
        m_sBoxBF(3, 127) = &HDB83ADF7
        m_sBoxBF(0, 128) = &HE93D5A68
        m_sBoxBF(1, 128) = &H948140F7
        m_sBoxBF(2, 128) = &HF64C261C
        m_sBoxBF(3, 128) = &H94692934
        m_sBoxBF(0, 129) = &H411520F7
        m_sBoxBF(1, 129) = &H7602D4F7
        m_sBoxBF(2, 129) = &HBCF46B2E
        m_sBoxBF(3, 129) = &HD4A20068
        m_sBoxBF(0, 130) = &HD4082471
        m_sBoxBF(1, 130) = &H3320F46A
        m_sBoxBF(2, 130) = &H43B7D4B7
        m_sBoxBF(3, 130) = &H500061AF
        m_sBoxBF(0, 131) = &H1E39F62E
        m_sBoxBF(1, 131) = &H97244546
        m_sBoxBF(2, 131) = &H14214F74
        m_sBoxBF(3, 131) = &HBF8B8840
        m_sBoxBF(0, 132) = &H4D95FC1D
        m_sBoxBF(1, 132) = &H96B591AF
        m_sBoxBF(2, 132) = &H70F4DDD3
        m_sBoxBF(3, 132) = &H66A02F45
        m_sBoxBF(0, 133) = &HBFBC09EC
        m_sBoxBF(1, 133) = &H3BD9785
        m_sBoxBF(2, 133) = &H7FAC6DD0
        m_sBoxBF(3, 133) = &H31CB8504
        m_sBoxBF(0, 134) = &H96EB27B3
        m_sBoxBF(1, 134) = &H55FD3941
        m_sBoxBF(2, 134) = &HDA2547E6
        m_sBoxBF(3, 134) = &HABCA0A9A
        m_sBoxBF(0, 135) = &H28507825
        m_sBoxBF(1, 135) = &H530429F4
        m_sBoxBF(2, 135) = &HA2C86DA
        m_sBoxBF(3, 135) = &HE9B66DFB
        m_sBoxBF(0, 136) = &H68DC1462
        m_sBoxBF(1, 136) = &HD7486900
        m_sBoxBF(2, 136) = &H680EC0A4
        m_sBoxBF(3, 136) = &H27A18DEE
        m_sBoxBF(0, 137) = &H4F3FFEA2
        m_sBoxBF(1, 137) = &HE887AD8C
        m_sBoxBF(2, 137) = &HB58CE006
        m_sBoxBF(3, 137) = &H7AF4D6B6
        m_sBoxBF(0, 138) = &HAACE1E7C
        m_sBoxBF(1, 138) = &HD3375FEC
        m_sBoxBF(2, 138) = &HCE78A399
        m_sBoxBF(3, 138) = &H406B2A42
        m_sBoxBF(0, 139) = &H20FE9E35
        m_sBoxBF(1, 139) = &HD9F385B9
        m_sBoxBF(2, 139) = &HEE39D7AB
        m_sBoxBF(3, 139) = &H3B124E8B
        m_sBoxBF(0, 140) = &H1DC9FAF7
        m_sBoxBF(1, 140) = &H4B6D1856
        m_sBoxBF(2, 140) = &H26A36631
        m_sBoxBF(3, 140) = &HEAE397B2
        m_sBoxBF(0, 141) = &H3A6EFA74
        m_sBoxBF(1, 141) = &HDD5B4332
        m_sBoxBF(2, 141) = &H6841E7F7
        m_sBoxBF(3, 141) = &HCA7820FB
        m_sBoxBF(0, 142) = &HFB0AF54E
        m_sBoxBF(1, 142) = &HD8FEB397
        m_sBoxBF(2, 142) = &H454056AC
        m_sBoxBF(3, 142) = &HBA489527
        m_sBoxBF(0, 143) = &H55533A3A
        m_sBoxBF(1, 143) = &H20838D87
        m_sBoxBF(2, 143) = &HFE6BA9B7
        m_sBoxBF(3, 143) = &HD096954B
        m_sBoxBF(0, 144) = &H55A867BC
        m_sBoxBF(1, 144) = &HA1159A58
        m_sBoxBF(2, 144) = &HCCA92963
        m_sBoxBF(3, 144) = &H99E1DB33
        m_sBoxBF(0, 145) = &HA62A4A56
        m_sBoxBF(1, 145) = &H3F3125F9
        m_sBoxBF(2, 145) = &H5EF47E1C
        m_sBoxBF(3, 145) = &H9029317C
        m_sBoxBF(0, 146) = &HFDF8E802
        m_sBoxBF(1, 146) = &H4272F70
        m_sBoxBF(2, 146) = &H80BB155C
        m_sBoxBF(3, 146) = &H5282CE3
        m_sBoxBF(0, 147) = &H95C11548
        m_sBoxBF(1, 147) = &HE4C66D22
        m_sBoxBF(2, 147) = &H48C1133F
        m_sBoxBF(3, 147) = &HC70F86DC
        m_sBoxBF(0, 148) = &H7F9C9EE
        m_sBoxBF(1, 148) = &H41041F0F
        m_sBoxBF(2, 148) = &H404779A4
        m_sBoxBF(3, 148) = &H5D886E17
        m_sBoxBF(0, 149) = &H325F51EB
        m_sBoxBF(1, 149) = &HD59BC0D1
        m_sBoxBF(2, 149) = &HF2BCC18F
        m_sBoxBF(3, 149) = &H41113564
        m_sBoxBF(0, 150) = &H257B7834
        m_sBoxBF(1, 150) = &H602A9C60
        m_sBoxBF(2, 150) = &HDFF8E8A3
        m_sBoxBF(3, 150) = &H1F636C1B
        m_sBoxBF(0, 151) = &HE12B4C2
        m_sBoxBF(1, 151) = &H2E1329E
        m_sBoxBF(2, 151) = &HAF664FD1
        m_sBoxBF(3, 151) = &HCAD18115
        m_sBoxBF(0, 152) = &H6B2395E0
        m_sBoxBF(1, 152) = &H333E92E1
        m_sBoxBF(2, 152) = &H3B240B62
        m_sBoxBF(3, 152) = &HEEBEB922
        m_sBoxBF(0, 153) = &H85B2A20E
        m_sBoxBF(1, 153) = &HE6BA0D99
        m_sBoxBF(2, 153) = &HDE720C8C
        m_sBoxBF(3, 153) = &H2DA2F728
        m_sBoxBF(0, 154) = &HD0127845
        m_sBoxBF(1, 154) = &H95B794FD
        m_sBoxBF(2, 154) = &H647D0862
        m_sBoxBF(3, 154) = &HE7CCF5F0
        m_sBoxBF(0, 155) = &H5449A36F
        m_sBoxBF(1, 155) = &H877D48FA
        m_sBoxBF(2, 155) = &HC39DFD27
        m_sBoxBF(3, 155) = &HF33E8D1E
        m_sBoxBF(0, 156) = &HA476341
        m_sBoxBF(1, 156) = &H992EFF74
        m_sBoxBF(2, 156) = &H3A6F6EAB
        m_sBoxBF(3, 156) = &HF4F8FD37
        m_sBoxBF(0, 157) = &HA812DC60
        m_sBoxBF(1, 157) = &HA1EBDDF8
        m_sBoxBF(2, 157) = &H991BE14C
        m_sBoxBF(3, 157) = &HDB6E6B0D
        m_sBoxBF(0, 158) = &HC67B5510
        m_sBoxBF(1, 158) = &H6D672C37
        m_sBoxBF(2, 158) = &H2765D43B
        m_sBoxBF(3, 158) = &HDCD0E804
        m_sBoxBF(0, 159) = &HF1290DC7
        m_sBoxBF(1, 159) = &HCC00FFA3
        m_sBoxBF(2, 159) = &HB5390F92
        m_sBoxBF(3, 159) = &H690FED0B
        m_sBoxBF(0, 160) = &H667B9FFB
        m_sBoxBF(1, 160) = &HCEDB7D9C
        m_sBoxBF(2, 160) = &HA091CF0B
        m_sBoxBF(3, 160) = &HD9155EA3
        m_sBoxBF(0, 161) = &HBB132F88
        m_sBoxBF(1, 161) = &H515BAD24
        m_sBoxBF(2, 161) = &H7B9479BF
        m_sBoxBF(3, 161) = &H763BD6EB
        m_sBoxBF(0, 162) = &H37392EB3
        m_sBoxBF(1, 162) = &HCC115979
        m_sBoxBF(2, 162) = &H8026E297
        m_sBoxBF(3, 162) = &HF42E312D
        m_sBoxBF(0, 163) = &H6842ADA7
        m_sBoxBF(1, 163) = &HC66A2B3B
        m_sBoxBF(2, 163) = &H12754CCC
        m_sBoxBF(3, 163) = &H782EF11C
        m_sBoxBF(0, 164) = &H6A124237
        m_sBoxBF(1, 164) = &HB79251E7
        m_sBoxBF(2, 164) = &H6A1BBE6
        m_sBoxBF(3, 164) = &H4BFB6350
        m_sBoxBF(0, 165) = &H1A6B1018
        m_sBoxBF(1, 165) = &H11CAEDFA
        m_sBoxBF(2, 165) = &H3D25BDD8
        m_sBoxBF(3, 165) = &HE2E1C3C9
        m_sBoxBF(0, 166) = &H44421659
        m_sBoxBF(1, 166) = &HA121386
        m_sBoxBF(2, 166) = &HD90CEC6E
        m_sBoxBF(3, 166) = &HD5ABEA2A
        m_sBoxBF(0, 167) = &H64AF674E
        m_sBoxBF(1, 167) = &HDA86A85F
        m_sBoxBF(2, 167) = &HBEBFE988
        m_sBoxBF(3, 167) = &H64E4C3FE
        m_sBoxBF(0, 168) = &H9DBC8057
        m_sBoxBF(1, 168) = &HF0F7C086
        m_sBoxBF(2, 168) = &H60787BF8
        m_sBoxBF(3, 168) = &H6003604D
        m_sBoxBF(0, 169) = &HD1FD8346
        m_sBoxBF(1, 169) = &HF6381FB0
        m_sBoxBF(2, 169) = &H7745AE04
        m_sBoxBF(3, 169) = &HD736FCCC
        m_sBoxBF(0, 170) = &H83426B33
        m_sBoxBF(1, 170) = &HF01EAB71
        m_sBoxBF(2, 170) = &HB0804187
        m_sBoxBF(3, 170) = &H3C005E5F
        m_sBoxBF(0, 171) = &H77A057BE
        m_sBoxBF(1, 171) = &HBDE8AE24
        m_sBoxBF(2, 171) = &H55464299
        m_sBoxBF(3, 171) = &HBF582E61
        m_sBoxBF(0, 172) = &H4E58F48F
        m_sBoxBF(1, 172) = &HF2DDFDA2
        m_sBoxBF(2, 172) = &HF474EF38
        m_sBoxBF(3, 172) = &H8789BDC2
        m_sBoxBF(0, 173) = &H5366F9C3
        m_sBoxBF(1, 173) = &HC8B38E74
        m_sBoxBF(2, 173) = &HB475F255
        m_sBoxBF(3, 173) = &H46FCD9B9
        m_sBoxBF(0, 174) = &H7AEB2661
        m_sBoxBF(1, 174) = &H8B1DDF84
        m_sBoxBF(2, 174) = &H846A0E79
        m_sBoxBF(3, 174) = &H915F95E2
        m_sBoxBF(0, 175) = &H466E598E
        m_sBoxBF(1, 175) = &H20B45770
        m_sBoxBF(2, 175) = &H8CD55591
        m_sBoxBF(3, 175) = &HC902DE4C
        m_sBoxBF(0, 176) = &HB90BACE1
        m_sBoxBF(1, 176) = &HBB8205D0
        m_sBoxBF(2, 176) = &H11A86248
        m_sBoxBF(3, 176) = &H7574A99E
        m_sBoxBF(0, 177) = &HB77F19B6
        m_sBoxBF(1, 177) = &HE0A9DC09
        m_sBoxBF(2, 177) = &H662D09A1
        m_sBoxBF(3, 177) = &HC4324633
        m_sBoxBF(0, 178) = &HE85A1F02
        m_sBoxBF(1, 178) = &H9F0BE8C
        m_sBoxBF(2, 178) = &H4A99A025
        m_sBoxBF(3, 178) = &H1D6EFE10
        m_sBoxBF(0, 179) = &H1AB93D1D
        m_sBoxBF(1, 179) = &HBA5A4DF
        m_sBoxBF(2, 179) = &HA186F20F
        m_sBoxBF(3, 179) = &H2868F169
        m_sBoxBF(0, 180) = &HDCB7DA83
        m_sBoxBF(1, 180) = &H573906FE
        m_sBoxBF(2, 180) = &HA1E2CE9B
        m_sBoxBF(3, 180) = &H4FCD7F52
        m_sBoxBF(0, 181) = &H50115E01
        m_sBoxBF(1, 181) = &HA70683FA
        m_sBoxBF(2, 181) = &HA002B5C4
        m_sBoxBF(3, 181) = &HDE6D027
        m_sBoxBF(0, 182) = &H9AF88C27
        m_sBoxBF(1, 182) = &H773F8641
        m_sBoxBF(2, 182) = &HC3604C06
        m_sBoxBF(3, 182) = &H61A806B5
        m_sBoxBF(0, 183) = &HF0177A28
        m_sBoxBF(1, 183) = &HC0F586E0
        m_sBoxBF(2, 183) = &H6058AA
        m_sBoxBF(3, 183) = &H30DC7D62
        m_sBoxBF(0, 184) = &H11E69ED7
        m_sBoxBF(1, 184) = &H2338EA63
        m_sBoxBF(2, 184) = &H53C2DD94
        m_sBoxBF(3, 184) = &HC2C21634
        m_sBoxBF(0, 185) = &HBBCBEE56
        m_sBoxBF(1, 185) = &H90BCB6DE
        m_sBoxBF(2, 185) = &HEBFC7DA1
        m_sBoxBF(3, 185) = &HCE591D76
        m_sBoxBF(0, 186) = &H6F05E409
        m_sBoxBF(1, 186) = &H4B7C0188
        m_sBoxBF(2, 186) = &H39720A3D
        m_sBoxBF(3, 186) = &H7C927C24
        m_sBoxBF(0, 187) = &H86E3725F
        m_sBoxBF(1, 187) = &H724D9DB9
        m_sBoxBF(2, 187) = &H1AC15BB4
        m_sBoxBF(3, 187) = &HD39EB8FC
        m_sBoxBF(0, 188) = &HED545578
        m_sBoxBF(1, 188) = &H8FCA5B5
        m_sBoxBF(2, 188) = &HD83D7CD3
        m_sBoxBF(3, 188) = &H4DAD0FC4
        m_sBoxBF(0, 189) = &H1E50EF5E
        m_sBoxBF(1, 189) = &HB161E6F8
        m_sBoxBF(2, 189) = &HA28514D9
        m_sBoxBF(3, 189) = &H6C51133C
        m_sBoxBF(0, 190) = &H6FD5C7E7
        m_sBoxBF(1, 190) = &H56E14EC4
        m_sBoxBF(2, 190) = &H362ABFCE
        m_sBoxBF(3, 190) = &HDDC6C837
        m_sBoxBF(0, 191) = &HD79A3234
        m_sBoxBF(1, 191) = &H92638212
        m_sBoxBF(2, 191) = &H670EFA8E
        m_sBoxBF(3, 191) = &H406000E0
        m_sBoxBF(0, 192) = &H3A39CE37
        m_sBoxBF(1, 192) = &HD3FAF5CF
        m_sBoxBF(2, 192) = &HABC27737
        m_sBoxBF(3, 192) = &H5AC52D1B
        m_sBoxBF(0, 193) = &H5CB0679E
        m_sBoxBF(1, 193) = &H4FA33742
        m_sBoxBF(2, 193) = &HD3822740
        m_sBoxBF(3, 193) = &H99BC9BBE
        m_sBoxBF(0, 194) = &HD5118E9D
        m_sBoxBF(1, 194) = &HBF0F7315
        m_sBoxBF(2, 194) = &HD62D1C7E
        m_sBoxBF(3, 194) = &HC700C47B
        m_sBoxBF(0, 195) = &HB78C1B6B
        m_sBoxBF(1, 195) = &H21A19045
        m_sBoxBF(2, 195) = &HB26EB1BE
        m_sBoxBF(3, 195) = &H6A366EB4
        m_sBoxBF(0, 196) = &H5748AB2F
        m_sBoxBF(1, 196) = &HBC946E79
        m_sBoxBF(2, 196) = &HC6A376D2
        m_sBoxBF(3, 196) = &H6549C2C8
        m_sBoxBF(0, 197) = &H530FF8EE
        m_sBoxBF(1, 197) = &H468DDE7D
        m_sBoxBF(2, 197) = &HD5730A1D
        m_sBoxBF(3, 197) = &H4CD04DC6
        m_sBoxBF(0, 198) = &H2939BBDB
        m_sBoxBF(1, 198) = &HA9BA4650
        m_sBoxBF(2, 198) = &HAC9526E8
        m_sBoxBF(3, 198) = &HBE5EE304
        m_sBoxBF(0, 199) = &HA1FAD5F0
        m_sBoxBF(1, 199) = &H6A2D519A
        m_sBoxBF(2, 199) = &H63EF8CE2
        m_sBoxBF(3, 199) = &H9A86EE22
        m_sBoxBF(0, 200) = &HC089C2B8
        m_sBoxBF(1, 200) = &H43242EF6
        m_sBoxBF(2, 200) = &HA51E03AA
        m_sBoxBF(3, 200) = &H9CF2D0A4
        m_sBoxBF(0, 201) = &H83C061BA
        m_sBoxBF(1, 201) = &H9BE96A4D
        m_sBoxBF(2, 201) = &H8FE51550
        m_sBoxBF(3, 201) = &HBA645BD6
        m_sBoxBF(0, 202) = &H2826A2F9
        m_sBoxBF(1, 202) = &HA73A3AE1
        m_sBoxBF(2, 202) = &H4BA99586
        m_sBoxBF(3, 202) = &HEF5562E9
        m_sBoxBF(0, 203) = &HC72FEFD3
        m_sBoxBF(1, 203) = &HF752F7DA
        m_sBoxBF(2, 203) = &H3F046F69
        m_sBoxBF(3, 203) = &H77FA0A59
        m_sBoxBF(0, 204) = &H80E4A915
        m_sBoxBF(1, 204) = &H87B08601
        m_sBoxBF(2, 204) = &H9B09E6AD
        m_sBoxBF(3, 204) = &H3B3EE593
        m_sBoxBF(0, 205) = &HE990FD5A
        m_sBoxBF(1, 205) = &H9E34D797
        m_sBoxBF(2, 205) = &H2CF0B7D9
        m_sBoxBF(3, 205) = &H22B8B51
        m_sBoxBF(0, 206) = &H96D5AC3A
        m_sBoxBF(1, 206) = &H17DA67D
        m_sBoxBF(2, 206) = &HD1CF3ED6
        m_sBoxBF(3, 206) = &H7C7D2D28
        m_sBoxBF(0, 207) = &H1F9F25CF
        m_sBoxBF(1, 207) = &HADF2B89B
        m_sBoxBF(2, 207) = &H5AD6B472
        m_sBoxBF(3, 207) = &H5A88F54C
        m_sBoxBF(0, 208) = &HE029AC71
        m_sBoxBF(1, 208) = &HE019A5E6
        m_sBoxBF(2, 208) = &H47B0ACFD
        m_sBoxBF(3, 208) = &HED93FA9B
        m_sBoxBF(0, 209) = &HE8D3C48D
        m_sBoxBF(1, 209) = &H283B57CC
        m_sBoxBF(2, 209) = &HF8D56629
        m_sBoxBF(3, 209) = &H79132E28
        m_sBoxBF(0, 210) = &H785F0191
        m_sBoxBF(1, 210) = &HED756055
        m_sBoxBF(2, 210) = &HF7960E44
        m_sBoxBF(3, 210) = &HE3D35E8C
        m_sBoxBF(0, 211) = &H15056DD4
        m_sBoxBF(1, 211) = &H88F46DBA
        m_sBoxBF(2, 211) = &H3A16125
        m_sBoxBF(3, 211) = &H564F0BD
        m_sBoxBF(0, 212) = &HC3EB9E15
        m_sBoxBF(1, 212) = &H3C9057A2
        m_sBoxBF(2, 212) = &H97271AEC
        m_sBoxBF(3, 212) = &HA93A072A
        m_sBoxBF(0, 213) = &H1B3F6D9B
        m_sBoxBF(1, 213) = &H1E6321F5
        m_sBoxBF(2, 213) = &HF59C66FB
        m_sBoxBF(3, 213) = &H26DCF319
        m_sBoxBF(0, 214) = &H7533D928
        m_sBoxBF(1, 214) = &HB155FDF5
        m_sBoxBF(2, 214) = &H3563482
        m_sBoxBF(3, 214) = &H8ABA3CBB
        m_sBoxBF(0, 215) = &H28517711
        m_sBoxBF(1, 215) = &HC20AD9F8
        m_sBoxBF(2, 215) = &HABCC5167
        m_sBoxBF(3, 215) = &HCCAD925F
        m_sBoxBF(0, 216) = &H4DE81751
        m_sBoxBF(1, 216) = &H3830DC8E
        m_sBoxBF(2, 216) = &H379D5862
        m_sBoxBF(3, 216) = &H9320F991
        m_sBoxBF(0, 217) = &HEA7A90C2
        m_sBoxBF(1, 217) = &HFB3E7BCE
        m_sBoxBF(2, 217) = &H5121CE64
        m_sBoxBF(3, 217) = &H774FBE32
        m_sBoxBF(0, 218) = &HA8B6E37E
        m_sBoxBF(1, 218) = &HC3293D46
        m_sBoxBF(2, 218) = &H48DE5369
        m_sBoxBF(3, 218) = &H6413E680
        m_sBoxBF(0, 219) = &HA2AE0810
        m_sBoxBF(1, 219) = &HDD6DB224
        m_sBoxBF(2, 219) = &H69852DFD
        m_sBoxBF(3, 219) = &H9072166
        m_sBoxBF(0, 220) = &HB39A460A
        m_sBoxBF(1, 220) = &H6445C0DD
        m_sBoxBF(2, 220) = &H586CDECF
        m_sBoxBF(3, 220) = &H1C20C8AE
        m_sBoxBF(0, 221) = &H5BBEF7DD
        m_sBoxBF(1, 221) = &H1B588D40
        m_sBoxBF(2, 221) = &HCCD2017F
        m_sBoxBF(3, 221) = &H6BB4E3BB
        m_sBoxBF(0, 222) = &HDDA26A7E
        m_sBoxBF(1, 222) = &H3A59FF45
        m_sBoxBF(2, 222) = &H3E350A44
        m_sBoxBF(3, 222) = &HBCB4CDD5
        m_sBoxBF(0, 223) = &H72EACEA8
        m_sBoxBF(1, 223) = &HFA6484BB
        m_sBoxBF(2, 223) = &H8D6612AE
        m_sBoxBF(3, 223) = &HBF3C6F47
        m_sBoxBF(0, 224) = &HD29BE463
        m_sBoxBF(1, 224) = &H542F5D9E
        m_sBoxBF(2, 224) = &HAEC2771B
        m_sBoxBF(3, 224) = &HF64E6370
        m_sBoxBF(0, 225) = &H740E0D8D
        m_sBoxBF(1, 225) = &HE75B1357
        m_sBoxBF(2, 225) = &HF8721671
        m_sBoxBF(3, 225) = &HAF537D5D
        m_sBoxBF(0, 226) = &H4040CB08
        m_sBoxBF(1, 226) = &H4EB4E2CC
        m_sBoxBF(2, 226) = &H34D2466A
        m_sBoxBF(3, 226) = &H115AF84
        m_sBoxBF(0, 227) = &HE1B00428
        m_sBoxBF(1, 227) = &H95983A1D
        m_sBoxBF(2, 227) = &H6B89FB4
        m_sBoxBF(3, 227) = &HCE6EA048
        m_sBoxBF(0, 228) = &H6F3F3B82
        m_sBoxBF(1, 228) = &H3520AB82
        m_sBoxBF(2, 228) = &H11A1D4B
        m_sBoxBF(3, 228) = &H277227F8
        m_sBoxBF(0, 229) = &H611560B1
        m_sBoxBF(1, 229) = &HE7933FDC
        m_sBoxBF(2, 229) = &HBB3A792B
        m_sBoxBF(3, 229) = &H344525BD
        m_sBoxBF(0, 230) = &HA08839E1
        m_sBoxBF(1, 230) = &H51CE794B
        m_sBoxBF(2, 230) = &H2F32C9B7
        m_sBoxBF(3, 230) = &HA01FBAC9
        m_sBoxBF(0, 231) = &HE01CC87E
        m_sBoxBF(1, 231) = &HBCC7D1F6
        m_sBoxBF(2, 231) = &HCF0111C3
        m_sBoxBF(3, 231) = &HA1E8AAC7
        m_sBoxBF(0, 232) = &H1A908749
        m_sBoxBF(1, 232) = &HD44FBD9A
        m_sBoxBF(2, 232) = &HD0DADECB
        m_sBoxBF(3, 232) = &HD50ADA38
        m_sBoxBF(0, 233) = &H339C32A
        m_sBoxBF(1, 233) = &HC6913667
        m_sBoxBF(2, 233) = &H8DF9317C
        m_sBoxBF(3, 233) = &HE0B12B4F
        m_sBoxBF(0, 234) = &HF79E59B7
        m_sBoxBF(1, 234) = &H43F5BB3A
        m_sBoxBF(2, 234) = &HF2D519FF
        m_sBoxBF(3, 234) = &H27D9459C
        m_sBoxBF(0, 235) = &HBF97222C
        m_sBoxBF(1, 235) = &H15E6FC2A
        m_sBoxBF(2, 235) = &HF91FC71
        m_sBoxBF(3, 235) = &H9B941525
        m_sBoxBF(0, 236) = &HFAE59361
        m_sBoxBF(1, 236) = &HCEB69CEB
        m_sBoxBF(2, 236) = &HC2A86459
        m_sBoxBF(3, 236) = &H12BAA8D1
        m_sBoxBF(0, 237) = &HB6C1075E
        m_sBoxBF(1, 237) = &HE3056A0C
        m_sBoxBF(2, 237) = &H10D25065
        m_sBoxBF(3, 237) = &HCB03A442
        m_sBoxBF(0, 238) = &HE0EC6E0E
        m_sBoxBF(1, 238) = &H1698DB3B
        m_sBoxBF(2, 238) = &H4C98A0BE
        m_sBoxBF(3, 238) = &H3278E964
        m_sBoxBF(0, 239) = &H9F1F9532
        m_sBoxBF(1, 239) = &HE0D392DF
        m_sBoxBF(2, 239) = &HD3A0342B
        m_sBoxBF(3, 239) = &H8971F21E
        m_sBoxBF(0, 240) = &H1B0A7441
        m_sBoxBF(1, 240) = &H4BA3348C
        m_sBoxBF(2, 240) = &HC5BE7120
        m_sBoxBF(3, 240) = &HC37632D8
        m_sBoxBF(0, 241) = &HDF359F8D
        m_sBoxBF(1, 241) = &H9B992F2E
        m_sBoxBF(2, 241) = &HE60B6F47
        m_sBoxBF(3, 241) = &HFE3F11D
        m_sBoxBF(0, 242) = &HE54CDA54
        m_sBoxBF(1, 242) = &H1EDAD891
        m_sBoxBF(2, 242) = &HCE6279CF
        m_sBoxBF(3, 242) = &HCD3E7E6F
        m_sBoxBF(0, 243) = &H1618B166
        m_sBoxBF(1, 243) = &HFD2C1D05
        m_sBoxBF(2, 243) = &H848FD2C5
        m_sBoxBF(3, 243) = &HF6FB2299
        m_sBoxBF(0, 244) = &HF523F357
        m_sBoxBF(1, 244) = &HA6327623
        m_sBoxBF(2, 244) = &H93A83531
        m_sBoxBF(3, 244) = &H56CCCD02
        m_sBoxBF(0, 245) = &HACF08162
        m_sBoxBF(1, 245) = &H5A75EBB5
        m_sBoxBF(2, 245) = &H6E163697
        m_sBoxBF(3, 245) = &H88D273CC
        m_sBoxBF(0, 246) = &HDE966292
        m_sBoxBF(1, 246) = &H81B949D0
        m_sBoxBF(2, 246) = &H4C50901B
        m_sBoxBF(3, 246) = &H71C65614
        m_sBoxBF(0, 247) = &HE6C6C7BD
        m_sBoxBF(1, 247) = &H327A140A
        m_sBoxBF(2, 247) = &H45E1D006
        m_sBoxBF(3, 247) = &HC3F27B9A
        m_sBoxBF(0, 248) = &HC9AA53FD
        m_sBoxBF(1, 248) = &H62A80F00
        m_sBoxBF(2, 248) = &HBB25BFE2
        m_sBoxBF(3, 248) = &H35BDD2F6
        m_sBoxBF(0, 249) = &H71126905
        m_sBoxBF(1, 249) = &HB2040222
        m_sBoxBF(2, 249) = &HB6CBCF7C
        m_sBoxBF(3, 249) = &HCD769C2B
        m_sBoxBF(0, 250) = &H53113EC0
        m_sBoxBF(1, 250) = &H1640E3D3
        m_sBoxBF(2, 250) = &H38ABBD60
        m_sBoxBF(3, 250) = &H2547ADF0
        m_sBoxBF(0, 251) = &HBA38209C
        m_sBoxBF(1, 251) = &HF746CE76
        m_sBoxBF(2, 251) = &H77AFA1C5
        m_sBoxBF(3, 251) = &H20756060
        m_sBoxBF(0, 252) = &H85CBFE4E
        m_sBoxBF(1, 252) = &H8AE88DD8
        m_sBoxBF(2, 252) = &H7AAAF9B0
        m_sBoxBF(3, 252) = &H4CF9AA7E
        m_sBoxBF(0, 253) = &H1948C25C
        m_sBoxBF(1, 253) = &H2FB8A8C
        m_sBoxBF(2, 253) = &H1C36AE4
        m_sBoxBF(3, 253) = &HD6EBE1F9
        m_sBoxBF(0, 254) = &H90D4F869
        m_sBoxBF(1, 254) = &HA65CDEA0
        m_sBoxBF(2, 254) = &H3F09252D
        m_sBoxBF(3, 254) = &HC208E69F
        m_sBoxBF(0, 255) = &HB74E6132
        m_sBoxBF(1, 255) = &HCE77E25B
        m_sBoxBF(2, 255) = &H578FDFE3
        m_sBoxBF(3, 255) = &H3AC372E6

End Sub

Public Sub Encryption_Blowfish_SetKey(KeyValue As String)

Dim i As Long
Dim j As Long
Dim K As Long
Dim dataX As Long
Dim datal As Long
Dim datar As Long
Dim Key() As Byte
Dim KeyLength As Long

'Do nothing if the key is buffered

    If (m_KeyValue = KeyValue) Then Exit Sub
    m_KeyValue = KeyValue

    'Convert the new key into a bytearray
    KeyLength = Len(KeyValue)
    Key() = StrConv(KeyValue, vbFromUnicode)

    'Create key-dependant p-boxes
    j = 0
    For i = 0 To (ROUNDS + 1)
        dataX = 0
        For K = 0 To 3
            Call CopyMem(ByVal VarPtr(dataX) + 1, dataX, 3)
            dataX = (dataX Or Key(j))
            j = j + 1
            If (j >= KeyLength) Then j = 0
        Next
        m_pBox(i) = m_pBox(i) Xor dataX
    Next

    datal = 0
    datar = 0
    For i = 0 To (ROUNDS + 1) Step 2
        Call Encryption_Blowfish_EncryptBlock(datal, datar)
        m_pBox(i) = datal
        m_pBox(i + 1) = datar
    Next

    'Create key-dependant s-boxes
    For i = 0 To 3
        For j = 0 To 255 Step 2
            Call Encryption_Blowfish_EncryptBlock(datal, datar)
            m_sBoxBF(i, j) = datal
            m_sBoxBF(i, j + 1) = datar
        Next
    Next

End Sub

Public Sub Encryption_CryptAPI_DecryptByte(ByteArray() As Byte, Optional Password As String)

'Convert the array into a string, decrypt it
'and then convert it back to an array

    ByteArray() = StrConv(Encryption_CryptAPI_DecryptString(StrConv(ByteArray(), vbUnicode), Password), vbFromUnicode)

End Sub

Public Sub Encryption_CryptAPI_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_CryptAPI_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_CryptAPI_DecryptString(Text As String, Optional Password As String) As String

'Set the new key if any was sent to the function

    If (Len(Password) > 0) Then Encryption_CryptAPI_SetKey Password

    'Return the decrypted data
    Encryption_CryptAPI_DecryptString = Encryption_CryptAPI_EncryptDecrypt(Text, False)

End Function

Public Sub Encryption_CryptAPI_EncryptByte(ByteArray() As Byte, Optional Password As String)

'Convert the array into a string, encrypt it
'and then convert it back to an array

    ByteArray() = StrConv(Encryption_CryptAPI_EncryptString(StrConv(ByteArray(), vbUnicode), Password), vbFromUnicode)

End Sub

Private Function Encryption_CryptAPI_EncryptDecrypt(ByVal Text As String, Encrypt As Boolean) As String

Dim hKey As Long
Dim hHash As Long
Dim lLength As Long
Dim hCryptProv As Long

'Get handle to CSP

    If (CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0) Then
        If (CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) = 0) Then
            Call Err.Raise(vbObjectError, , "Error during CryptAcquireContext for a new key container." & vbNewLine & "A container with this name probably already exists.")
        End If
    End If

    'Create a hash object to calculate a session
    'key from the password (instead of encrypting
    'with the actual key)
    If (CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0) Then
        Call Err.Raise(vbObjectError, , "Could not create a Hash Object (CryptCreateHash API)")
    End If

    'Hash the password
    If (CryptHashData(hHash, m_KeyS, Len(m_KeyS), 0) = 0) Then
        Call Err.Raise(vbObjectError, , "Could not calculate a Hash Value (CryptHashData API)")
    End If

    'Derive a session key from the hash object
    If (CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0) Then
        Call Err.Raise(vbObjectError, , "Could not create a session key (CryptDeriveKey API)")
    End If

    'Encrypt or decrypt depending on the Encrypt parameter
    lLength = Len(Text)
    If (Encrypt) Then
        If (CryptEncrypt(hKey, 0, 1, 0, Text, lLength, lLength) = 0) Then
            Call Err.Raise(vbObjectError, , "Error during CryptEncrypt.")
        End If
    Else
        If (CryptDecrypt(hKey, 0, 1, 0, Text, lLength) = 0) Then
            Call Err.Raise(vbObjectError, , "Error during CryptDecrypt.")
        End If
    End If

    'Return the encrypted/decrypted data
    Encryption_CryptAPI_EncryptDecrypt = Left$(Text, lLength)

    'Destroy the session key
    If (hKey <> 0) Then Call CryptDestroyKey(hKey)

    'Destroy the hash object
    If (hHash <> 0) Then Call CryptDestroyHash(hHash)

    'Release provider handle
    If (hCryptProv <> 0) Then Call CryptReleaseContext(hCryptProv, 0)

End Function

Public Sub Encryption_CryptAPI_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_CryptAPI_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_CryptAPI_EncryptString(Text As String, Optional Password As String) As String

'Set the new key if any was sent to the function

    If (Len(Password) > 0) Then Encryption_CryptAPI_SetKey Password

    'Return the encrypted data
    Encryption_CryptAPI_EncryptString = Encryption_CryptAPI_EncryptDecrypt(Text, True)

End Function

Public Sub Encryption_CryptAPI_SetKey(New_Value As String)

'Do nothing if no change was made

    If (m_KeyS = New_Value) Then Exit Sub

    'Set the new key
    m_KeyS = New_Value

End Sub

Private Static Sub Encryption_Gost_DecryptBlock(LeftWord As Long, RightWord As Long)

Dim i As Long

    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(1))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(2))
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(3))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(4))
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(5))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(6))
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(7))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(8))
    For i = 1 To 3
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(8))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(7))
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(6))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(5))
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(4))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(3))
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(2))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(1))
    Next

End Sub

Public Function Encryption_Gost_DecryptByte(ByteArray() As Byte, Optional Key As String) As String

Dim Offset As Long
Dim OrigLen As Long
Dim LeftWord As Long
Dim RightWord As Long
Dim CipherLen As Long
Dim CipherLeft As Long
Dim CipherRight As Long

'Set the key if one was passed to the function

    If (Len(Key) > 0) Then Encryption_Gost_SetKey Key

    'Get the size of the ciphertext
    CipherLen = UBound(ByteArray) + 1

    'Decrypt the data in 64-bit blocks
    For Offset = 0 To (CipherLen - 1) Step 8
        'Get the next block
        Call Encryption_Misc_GetWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_GetWord(RightWord, ByteArray(), Offset + 4)

        'Decrypt the block
        Call Encryption_Gost_DecryptBlock(RightWord, LeftWord)

        'XOR with the previous cipherblock
        LeftWord = LeftWord Xor CipherLeft
        RightWord = RightWord Xor CipherRight

        'Store the current ciphertext to use
        'XOR with the next block plaintext
        Call Encryption_Misc_GetWord(CipherLeft, ByteArray(), Offset)
        Call Encryption_Misc_GetWord(CipherRight, ByteArray(), Offset + 4)

        'Store the encrypted block
        Call Encryption_Misc_PutWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_PutWord(RightWord, ByteArray(), Offset + 4)

    Next

    'Get the size of the original array
    Call CopyMem(OrigLen, ByteArray(8), 4)

    'Make sure OrigLen is a reasonable value,
    'if we used the wrong key the next couple
    'of statements could be dangerous (GPF)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then
        Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Gost decryption")
    End If

    'Resize the bytearray to hold only the plaintext
    'and not the extra information added by the
    'encryption routine
    Call CopyMem(ByteArray(0), ByteArray(12), OrigLen)
    ReDim Preserve ByteArray(OrigLen - 1)

End Function

Public Sub Encryption_Gost_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_Gost_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_Gost_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Gost_DecryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the text into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call Encryption_Gost_DecryptByte(ByteArray(), Key)

    'Convert the byte array back to a string
    Encryption_Gost_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Static Sub Encryption_Gost_EncryptBlock(LeftWord As Long, RightWord As Long)

Dim i As Long

    For i = 1 To 3
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(1))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(2))
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(3))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(4))
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(5))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(6))
        RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(7))
        LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(8))
    Next
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(8))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(7))
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(6))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(5))
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(4))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(3))
    RightWord = RightWord Xor Encryption_Gost_F(LeftWord, K(2))
    LeftWord = LeftWord Xor Encryption_Gost_F(RightWord, K(1))

End Sub

Public Function Encryption_Gost_EncryptByte(ByteArray() As Byte, Optional Key As String) As String

Dim Offset As Long
Dim OrigLen As Long
Dim LeftWord As Long
Dim RightWord As Long
Dim CipherLen As Long
Dim CipherLeft As Long
Dim CipherRight As Long

'Set the key if one was passed to the function

    If (Len(Key) > 0) Then Encryption_Gost_SetKey Key

    'Get the length of the plaintext
    OrigLen = UBound(ByteArray) + 1

    'First we add 12 bytes (4 bytes for the
    'length and 8 bytes for the seed values
    'for the CBC routine), and the ciphertext
    'must be a multiple of 8 bytes
    CipherLen = OrigLen + 12
    If (CipherLen Mod 8 <> 0) Then
        CipherLen = CipherLen + 8 - (CipherLen Mod 8)
    End If
    ReDim Preserve ByteArray(CipherLen - 1)
    Call CopyMem(ByteArray(12), ByteArray(0), OrigLen)

    'Store the length descriptor in bytes [9-12]
    Call CopyMem(ByteArray(8), OrigLen, 4)

    'Store a block of random data in bytes [1-8],
    'these work as seed values for the CBC routine
    'and is used to produce different ciphertext
    'even when encrypting the same data with the
    'same key)
    Call Randomize
    Call CopyMem(ByteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(ByteArray(4), CLng(2147483647 * Rnd), 4)

    'Encrypt the data
    For Offset = 0 To (CipherLen - 1) Step 8
        'Get the next block of plaintext
        Call Encryption_Misc_GetWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_GetWord(RightWord, ByteArray(), Offset + 4)

        'XOR the plaintext with the previous
        'ciphertext (CBC, Cipher-Block Chaining)
        LeftWord = LeftWord Xor CipherLeft
        RightWord = RightWord Xor CipherRight

        'Encrypt the block
        Call Encryption_Gost_EncryptBlock(LeftWord, RightWord)

        'Store the block
        Call Encryption_Misc_PutWord(LeftWord, ByteArray(), Offset)
        Call Encryption_Misc_PutWord(RightWord, ByteArray(), Offset + 4)

        'Store the cipherblocks (for CBC)
        CipherLeft = LeftWord
        CipherRight = RightWord
        
    Next

End Function

Public Sub Encryption_Gost_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_Gost_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_Gost_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Gost_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the text into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call Encryption_Gost_EncryptByte(ByteArray(), Key)

    'Convert the byte array back to a string
    Encryption_Gost_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Static Function Encryption_Gost_F(R As Long, K As Long) As Long

Dim X As Long
Dim xb(0 To 3) As Byte
Dim xx(0 To 3) As Byte
Dim a As Byte, b As Byte, C As Byte, D As Byte

    If (m_RunningCompiled) Then
        X = R + K
    Else
        X = Encryption_Misc_UnsignedAdd(R, K)
    End If

    'Extract byte sequence
    D = X And &HFF
    X = X \ 256
    C = X And &HFF
    X = X \ 256
    b = X And &HFF
    X = X \ 256
    a = X And &HFF

    'Key-dependant substutions
    xb(0) = k21(a)
    xb(1) = k43(b)
    xb(2) = k65(C)
    xb(3) = k87(D)

    'LeftShift 11 bits
    xx(0) = ((xb(3) And 31) * 8) Or ((xb(2) And 224) \ 32)
    xx(1) = ((xb(0) And 31) * 8) Or ((xb(3) And 224) \ 32)
    xx(2) = ((xb(1) And 31) * 8) Or ((xb(0) And 224) \ 32)
    xx(3) = ((xb(2) And 31) * 8) Or ((xb(1) And 224) \ 32)
    Call CopyMem(Encryption_Gost_F, xx(0), 4)

End Function

Private Sub Encryption_Gost_Init()

Dim a As Long
Dim b As Long
Dim C As Long
Dim LeftWord As Long
Dim s(0 To 7) As Variant

'We need to check if we are running in compiled
'(EXE) mode or in the IDE, this will allow us to
'use optimized code with unsigned integers in
'compiled mode without any overflow errors when
'running the code in the IDE

    On Local Error Resume Next
        m_RunningCompiled = ((2147483647 + 1) < 0)

        'Initialize s-boxes
        s(0) = Array(6, 5, 1, 7, 14, 0, 4, 10, 11, 9, 3, 13, 8, 12, 2, 15)
        s(1) = Array(14, 13, 9, 0, 8, 10, 12, 4, 7, 15, 6, 11, 3, 1, 5, 2)
        s(2) = Array(6, 5, 1, 7, 2, 4, 10, 0, 11, 13, 14, 3, 8, 12, 15, 9)
        s(3) = Array(8, 7, 3, 9, 6, 4, 14, 5, 2, 13, 0, 12, 1, 11, 10, 15)
        s(4) = Array(10, 9, 6, 11, 5, 1, 8, 4, 0, 13, 7, 2, 14, 3, 15, 12)
        s(5) = Array(5, 3, 0, 6, 11, 13, 4, 14, 10, 7, 1, 12, 2, 8, 15, 9)
        s(6) = Array(2, 1, 12, 3, 11, 13, 15, 7, 10, 6, 9, 14, 0, 8, 4, 5)
        s(7) = Array(6, 5, 1, 7, 8, 9, 4, 2, 15, 3, 13, 12, 10, 14, 11, 0)

        'Convert the variants to a 2-dimensional array
        For a = 0 To 15
            For b = 0 To 7
                sBox(b, a) = s(b)(a)
            Next
        Next

        'Calculate the substitutions
        For a = 0 To 255
            k87(a) = Encryption_Gost_lBSL(CLng(sBox(7, Encryption_Gost_lBSR(a, 4))), 4) Or sBox(6, a And 15)
            k65(a) = Encryption_Gost_lBSL(CLng(sBox(5, Encryption_Gost_lBSR(a, 4))), 4) Or sBox(4, a And 15)
            k43(a) = Encryption_Gost_lBSL(CLng(sBox(3, Encryption_Gost_lBSR(a, 4))), 4) Or sBox(2, a And 15)
            k21(a) = Encryption_Gost_lBSL(CLng(sBox(1, Encryption_Gost_lBSR(a, 4))), 4) Or sBox(0, a And 15)
        Next

End Sub

Private Static Function Encryption_Gost_lBSL(ByVal lInput As Long, bShiftBits As Byte) As Long

    Encryption_Gost_lBSL = (lInput And (2 ^ (31 - bShiftBits) - 1)) * 2 ^ bShiftBits
    If (lInput And 2 ^ (31 - bShiftBits)) = 2 ^ (31 - bShiftBits) Then Encryption_Gost_lBSL = (Encryption_Gost_lBSL Or &H80000000)

End Function

Private Static Function Encryption_Gost_lBSR(ByVal lInput As Long, bShiftBits As Byte) As Long

    If bShiftBits = 31 Then
        If lInput < 0 Then Encryption_Gost_lBSR = &HFFFFFFFF Else Encryption_Gost_lBSR = 0
    Else
        Encryption_Gost_lBSR = (lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits
    End If

End Function

Public Sub Encryption_Gost_SetKey(New_Value As String)

Dim a As Long
Dim Key() As Byte
Dim KeyLen As Long
Dim ByteArray() As Byte

'Do nothing if no change was made

    If (m_KeyValue = New_Value) Then Exit Sub

    'Convert the key into a bytearray
    KeyLen = Len(New_Value)
    Key() = StrConv(New_Value, vbFromUnicode)

    'Create a 32-byte key
    ReDim ByteArray(0 To 31)
    For a = 0 To 31
        ByteArray(a) = Key(a Mod KeyLen)
    Next

    'Create the key
    Call CopyMem(K(1), ByteArray(0), 32)

    'Show this key is buffered
    m_KeyValue = New_Value

End Sub

Private Function Encryption_Misc_FileExist(FileName As String) As Boolean

    On Error GoTo NotExist

    Call FileLen(FileName)
    Encryption_Misc_FileExist = True

Exit Function

NotExist:

End Function

Private Static Sub Encryption_Misc_GetWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)

Dim bb(0 To 3) As Byte

    bb(3) = CryptBuffer(Offset)
    bb(2) = CryptBuffer(Offset + 1)
    bb(1) = CryptBuffer(Offset + 2)
    bb(0) = CryptBuffer(Offset + 3)
    Call CopyMem(LongValue, bb(0), 4)

End Sub

Function Encryption_Misc_HexToStr(HexText As String, Optional ByVal Separators As Long = 1) As String

Dim a As Long
Dim Pos As Long
Dim PosAdd As Long
Dim ByteSize As Long
Dim HexByte() As Byte
Dim ByteArray() As Byte

'Initialize the hex routine

    If (Not m_Encryption_Misc_InitHex) Then Call Encryption_Misc_InitHex

    'The destination string is half
    'the size of the source string
    'when the separators are removed
    If (Len(HexText) = 2) Then
        ByteSize = 1
    Else
        ByteSize = ((Len(HexText) + 1) \ (2 + Separators))
    End If
    ReDim ByteArray(0 To ByteSize - 1)

    'Convert every HEX code to the
    'equivalent ASCII character
    PosAdd = 2 + Separators
    HexByte() = StrConv(HexText, vbFromUnicode)
    For a = 0 To (ByteSize - 1)
        ByteArray(a) = m_HexToByte(HexByte(Pos), HexByte(Pos + 1))
        Pos = Pos + PosAdd
    Next

    'Now finally convert the byte
    'array to the return string
    Encryption_Misc_HexToStr = StrConv(ByteArray, vbUnicode)

End Function

Private Sub Encryption_Misc_InitHex()

Dim a As Long
Dim b As Long
Dim HexBytes() As Byte
Dim HexString As String

'The routine is initialized

    m_Encryption_Misc_InitHex = True

    'Create a string with all hex values
    HexString = String$(512, "0")
    For a = 1 To 255
        Mid$(HexString, 1 + a * 2 + -(a < 16)) = Hex(a)
    Next
    HexBytes = StrConv(HexString, vbFromUnicode)

    'Create the Str->Hex array
    For a = 0 To 255
        m_ByteToHex(a, 0) = HexBytes(a * 2)
        m_ByteToHex(a, 1) = HexBytes(a * 2 + 1)
    Next

    'Create the Str->Hex array
    For a = 0 To 255
        m_HexToByte(m_ByteToHex(a, 0), m_ByteToHex(a, 1)) = a
    Next

End Sub

Private Static Sub Encryption_Misc_PutWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)

Dim bb(0 To 3) As Byte

    Call CopyMem(bb(0), LongValue, 4)
    CryptBuffer(Offset) = bb(3)
    CryptBuffer(Offset + 1) = bb(2)
    CryptBuffer(Offset + 2) = bb(1)
    CryptBuffer(Offset + 3) = bb(0)

End Sub

Function Encryption_Misc_StrToHex(Text As String, Optional Separator As String = " ") As String

Dim a As Long
Dim Pos As Long
Dim Char As Byte
Dim PosAdd As Long
Dim ByteSize As Long
Dim ByteArray() As Byte
Dim ByteReturn() As Byte
Dim SeparatorLen As Long
Dim SeparatorChar As Byte

'Initialize the hex routine

    If (Not m_Encryption_Misc_InitHex) Then Call Encryption_Misc_InitHex

    'Initialize variables
    SeparatorLen = Len(Separator)

    'Create the destination bytearray, this
    'will be converted to a string later
    ByteSize = (Len(Text) * 2 + (Len(Text) - 1) * SeparatorLen)
    ReDim ByteReturn(ByteSize - 1)
    Call FillMemory(ByteReturn(0), ByteSize, Asc(Separator))

    'We convert the source string into a
    'byte array to speed this up a tad
    ByteArray() = StrConv(Text, vbFromUnicode)

    'Now convert every character to
    'it's equivalent HEX code
    PosAdd = 2 + SeparatorLen
    For a = 0 To (Len(Text) - 1)
        ByteReturn(Pos) = m_ByteToHex(ByteArray(a), 0)
        ByteReturn(Pos + 1) = m_ByteToHex(ByteArray(a), 1)
        Pos = Pos + PosAdd
    Next

    'Convert the bytearray to a string
    Encryption_Misc_StrToHex = StrConv(ByteReturn(), vbUnicode)

End Function

Private Static Function Encryption_Misc_UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long

Dim x1(0 To 3) As Byte
Dim x2(0 To 3) As Byte
Dim xx(0 To 3) As Byte
Dim Rest As Long
Dim Value As Long
Dim a As Long

    Call CopyMem(x1(0), Data1, 4)
    Call CopyMem(x2(0), Data2, 4)

    Rest = 0
    For a = 0 To 3
        Value = CLng(x1(a)) + CLng(x2(a)) + Rest
        xx(a) = Value And 255
        Rest = Value \ 256
    Next

    Call CopyMem(Encryption_Misc_UnsignedAdd, xx(0), 4)

End Function

Private Function Encryption_Misc_UnsignedDel(Data1 As Long, Data2 As Long) As Long

Dim x1(0 To 3) As Byte
Dim x2(0 To 3) As Byte
Dim xx(0 To 3) As Byte
Dim Rest As Long
Dim Value As Long
Dim a As Long

    Call CopyMem(x1(0), Data1, 4)
    Call CopyMem(x2(0), Data2, 4)
    Call CopyMem(xx(0), Encryption_Misc_UnsignedDel, 4)

    For a = 0 To 3
        Value = CLng(x1(a)) - CLng(x2(a)) - Rest
        If (Value < 0) Then
            Value = Value + 256
            Rest = 1
        Else
            Rest = 0
        End If
        xx(a) = Value
    Next

    Call CopyMem(Encryption_Misc_UnsignedDel, xx(0), 4)

End Function

Public Function Encryption_Skipjack_DecryptByte(ByteArray() As Byte, Optional Key As String) As String

Dim i As Long
Dim u As Long
Dim K As Long
Dim Temp As Byte
Dim Round As Long
Dim Offset As Long
Dim OrigLen As Long
Dim CipherLen As Long
Dim G(0 To 5) As Byte
Dim Counter(0 To 32) As Byte
Dim w(0 To 3, 0 To 33) As Integer

    'Set the new key
    If (Len(Key) > 0) Then Encryption_Skipjack_SetKey Key

    'Get the size of the bytearray
    CipherLen = UBound(ByteArray) + 1

    'Switch bytes to convert bytes into integers
    For Offset = 0 To (CipherLen - 1) Step 2
        Temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = Temp
    Next

    'Decrypt the data 8-bytes at a time
    For Offset = 0 To (CipherLen - 1) Step 8
        'Read the next 4 integers from the bytearray
        Call CopyMem(w(0, 32), ByteArray(Offset), 8)

        K = 32
        u = 31
        For i = 0 To 32
            Counter(i) = i + 1
        Next

        For Round = 1 To 2
            'Execute Rule B(inv)
            For i = 1 To 8
                Call CopyMem(G(4), w(1, K), 2)
                G(3) = m_SJF(G(5) Xor m_SJKey(4 * u + 3)) Xor G(4)
                G(2) = m_SJF(G(3) Xor m_SJKey(4 * u + 2)) Xor G(5)
                G(0) = m_SJF(G(2) Xor m_SJKey(4 * u + 1)) Xor G(3)
                G(1) = m_SJF(G(0) Xor m_SJKey(4 * u)) Xor G(2)
                Call CopyMem(w(0, K - 1), G(0), 2)
                w(1, K - 1) = w(0, K - 1) Xor w(2, K) Xor Counter(K - 1)
                w(2, K - 1) = w(3, K)
                w(3, K - 1) = w(0, K)
                u = u - 1
                K = K - 1
            Next

            'Execute Rule A(inv)
            For i = 1 To 8
                Call CopyMem(G(4), w(1, K), 2)
                G(3) = m_SJF(G(5) Xor m_SJKey(4 * u + 3)) Xor G(4)
                G(2) = m_SJF(G(3) Xor m_SJKey(4 * u + 2)) Xor G(5)
                G(0) = m_SJF(G(2) Xor m_SJKey(4 * u + 1)) Xor G(3)
                G(1) = m_SJF(G(0) Xor m_SJKey(4 * u)) Xor G(2)
                Call CopyMem(w(0, K - 1), G(0), 2)
                w(1, K - 1) = w(2, K)
                w(2, K - 1) = w(3, K)
                w(3, K - 1) = w(0, K) Xor w(1, K) Xor Counter(K - 1)
                u = u - 1
                K = K - 1
            Next
        Next

        'XOR with the previous encrypted data
        w(0, 0) = w(0, 0) Xor w(0, 33)
        w(1, 0) = w(1, 0) Xor w(1, 33)
        w(2, 0) = w(2, 0) Xor w(2, 33)
        w(3, 0) = w(3, 0) Xor w(3, 33)

        'Store the updated integer values in the bytearray
        Call CopyMem(ByteArray(Offset), w(0, 0), 8)

        'Save the encrypted data for later use where blocks are XOR'ed (CBC, Cipher-Block Chaining) for increased security
        Call CopyMem(w(0, 33), w(0, 32), 8)

    Next

    'Switch bytes to convert bytes into integers
    For Offset = 0 To (CipherLen - 1) Step 2
        Temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = Temp
    Next

    'Get the size of the original array
    Call CopyMem(OrigLen, ByteArray(8), 4)

    'Make sure OrigLen is a reasonable value,
    'if we used the wrong key the next couple
    'of statements could be dangerous (GPF)
    If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then
        Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Skipjack decryption")
    End If

    'Resize the bytearray to hold only the plaintext
    'and not the extra information added by the
    'encryption routine
    Call CopyMem(ByteArray(0), ByteArray(12), OrigLen)
    ReDim Preserve ByteArray(OrigLen - 1)

End Function

Public Sub Encryption_Skipjack_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_Skipjack_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_Skipjack_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Skipjack_DecryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the string into a bytearray

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the bytearray
    Call Encryption_Skipjack_DecryptByte(ByteArray(), Key)

    'Convert the bytearray back to a string
    Encryption_Skipjack_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub Encryption_Skipjack_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim i As Long
Dim K As Long
Dim Temp As Byte
Dim Round As Long
Dim Offset As Long
Dim OrigLen As Long
Dim Counter As Long
Dim G(0 To 5) As Byte
Dim CipherLen As Long
Dim w(0 To 3, 0 To 32) As Integer

'Be sure the key is initialized

    If (Len(Key) > 0) Then Encryption_Skipjack_SetKey Key

    'Save the size of the bytearray for future
    'reference (for the length descriptor)
    OrigLen = UBound(ByteArray) + 1

    'First we add 12 bytes (4 bytes for the
    'length and 8 bytes for the seed values
    'for the CBC routine), and the ciphertext
    'must be a multiple of 8 bytes
    CipherLen = OrigLen + 12
    If (CipherLen Mod 8 <> 0) Then
        CipherLen = CipherLen + 8 - (CipherLen Mod 8)
    End If
    ReDim Preserve ByteArray(CipherLen - 1)
    Call CopyMem(ByteArray(12), ByteArray(0), OrigLen)

    'Store the length descriptor in bytes [9-12]
    Call CopyMem(ByteArray(8), OrigLen, 4)

    'Store a block of random data in bytes [1-8],
    'these work as seed values for the CBC routine
    'and is used to produce different ciphertext
    'even when encrypting the same data with the
    'same key)
    Call Randomize
    Call CopyMem(ByteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(ByteArray(4), CLng(2147483647 * Rnd), 4)

    'Switch array of bytes into array of integers
    For Offset = 0 To (CipherLen - 1) Step 2
        Temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = Temp
    Next

    'Encrypt the data 8-bytes at a time
    For Offset = 0 To (CipherLen - 1) Step 8
        'Read the next 4 integers from the bytearray
        Call CopyMem(w(0, 0), ByteArray(Offset), 8)

        'XOR the plaintext with the previous
        'ciphertext (CBC, Cipher-Block Chaining)
        w(0, 0) = w(0, 0) Xor w(0, 32)
        w(1, 0) = w(1, 0) Xor w(1, 32)
        w(2, 0) = w(2, 0) Xor w(2, 32)
        w(3, 0) = w(3, 0) Xor w(3, 32)

        K = 0
        Counter = 1

        For Round = 1 To 2
            'Execute RULE A
            For i = 1 To 8
                Call CopyMem(G(0), w(0, K), 2)
                G(2) = m_SJF(G(0) Xor m_SJKey(4 * K)) Xor G(1)
                G(3) = m_SJF(G(2) Xor m_SJKey(4 * K + 1)) Xor G(0)
                G(5) = m_SJF(G(3) Xor m_SJKey(4 * K + 2)) Xor G(2)
                G(4) = m_SJF(G(5) Xor m_SJKey(4 * K + 3)) Xor G(3)
                Call CopyMem(w(1, K + 1), G(4), 2)
                w(0, K + 1) = w(1, K + 1) Xor w(3, K) Xor Counter
                w(2, K + 1) = w(1, K)
                w(3, K + 1) = w(2, K)
                Counter = Counter + 1
                K = K + 1
            Next

            'Execute RULE B
            For i = 1 To 8
                Call CopyMem(G(0), w(0, K), 2)
                G(2) = m_SJF(G(0) Xor m_SJKey(4 * K)) Xor G(1)
                G(3) = m_SJF(G(2) Xor m_SJKey(4 * K + 1)) Xor G(0)
                G(5) = m_SJF(G(3) Xor m_SJKey(4 * K + 2)) Xor G(2)
                G(4) = m_SJF(G(5) Xor m_SJKey(4 * K + 3)) Xor G(3)
                Call CopyMem(w(1, K + 1), G(4), 2)
                w(0, K + 1) = w(3, K)
                w(2, K + 1) = w(0, K) Xor w(1, K) Xor Counter
                w(3, K + 1) = w(2, K)
                Counter = Counter + 1
                K = K + 1
            Next
        Next

        'Store the new integer values into the array
        Call CopyMem(ByteArray(Offset), w(0, 32), 8)
        
    Next

    'Switch array of integers back to array of bytes
    For Offset = 0 To (CipherLen - 1) Step 2
        Temp = ByteArray(Offset)
        ByteArray(Offset) = ByteArray(Offset + 1)
        ByteArray(Offset + 1) = Temp
    Next

End Sub

Public Sub Encryption_Skipjack_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_Skipjack_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_Skipjack_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Skipjack_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the string into a bytearray

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the bytearray
    Call Encryption_Skipjack_EncryptByte(ByteArray(), Key)

    'Convert the bytearray back to a string
    Encryption_Skipjack_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Sub Encryption_Skipjack_Init()

Dim a As Long
Dim Ftable As Variant

'Initialize the F-table

    Ftable = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", "F9", _
             "E7", "2D", "4D", "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", _
             "0A", "DF", "02", "A0", "17", "F1", "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", _
             "96", "84", "6B", "BA", "F2", "63", "9A", "19", "7C", "AE", "E5", "F5", "F7", "16", "6A", "A2", _
             "39", "B6", "7B", "0F", "C1", "93", "81", "1B", "EE", "B4", "1A", "EA", "D0", "91", "2F", "B8", _
             "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", "5F", "66", "0B", "D8", "90", _
             "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", "9B", "76", _
             "97", "FC", "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", _
             "42", "ED", "9E", "6E", "49", "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", _
             "89", "CB", "30", "1F", "8D", "C6", "8F", "AA", "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", _
             "70", "88", "61", "2C", "9F", "0D", "2B", "87", "50", "82", "54", "64", "26", "7D", "03", "40", _
             "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", "7F", "AB", "E6", "3E", "5B", "A5", _
             "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", "8C", "0E", "E2", _
             "0C", "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", _
             "08", "77", "11", "BE", "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", _
             "5E", "6C", "A9", "13", "57", "25", "B5", "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")

    'Convert the F-table into a linear byte
    'array for faster access later
    For a = 0 To 255
        m_SJF(a) = Val("&H" & Ftable(a))
    Next

    'Initialize the CBC (random) seed values to work
    'as a starting ground for the CRC XOR (this is
    'optional but must be the same for the both
    'transmitter and receiver)
    'm_CBCSeed(0) = -923
    'm_CBCSeed(1) = 19843
    'm_CBCSeed(2) = 154
    'm_CBCSeed(3) = 8123

End Sub

Public Sub Encryption_Skipjack_SetKey(New_Value As String)

Dim i As Long
Dim Pass() As Byte
Dim PassLen As Long

'Do nothing if the new key is the same as the last
'one used because that it is already initialized

    If (New_Value = m_SJKeyValue) Then Exit Sub

    'The key must have at least one character
    If (Len(New_Value) = 0) Then
        Err.Raise vbObjectError, , "Invalid key given to SkipJack encryption or decryption (Zero Length)"
    End If

    'Convert the password into a bytearray
    PassLen = Len(New_Value)
    Pass() = StrConv(New_Value, vbFromUnicode)

    'Extract a 128-bit key from the bytearray
    For i = 0 To 127
        m_SJKey(i) = Pass(i Mod PassLen)
    Next

    'Store a copy of the key as string value to
    'show that this key is buffered
    m_SJKeyValue = New_Value

End Sub

Private Sub Encryption_Twofish_DecryptBlock(DWord() As Long)

Dim K As Long
Dim R As Long
Dim t0 As Long
Dim t1 As Long

    DWord(2) = DWord(2) Xor sKeyTF(OUTPUT_WHITEN)
    DWord(3) = DWord(3) Xor sKeyTF(OUTPUT_WHITEN + 1)
    DWord(0) = DWord(4) Xor sKeyTF(OUTPUT_WHITEN + 2)
    DWord(1) = DWord(5) Xor sKeyTF(OUTPUT_WHITEN + 3)

    K = ROUND_SUBKEYS + 2 * ROUNDSTF - 1
    For R = 0 To ROUNDSTF - 1 Step 2
        If (m_RunningCompiled) Then
            t0 = Encryption_Twofish_Fe32(DWord(2), 0)
            t1 = Encryption_Twofish_Fe32(DWord(3), 3)
            t0 = t0 + t1
            DWord(1) = Encryption_Twofish_Rot1(DWord(1) Xor (t0 + t1 + sKeyTF(K)))
            K = K - 1
            DWord(0) = Encryption_Twofish_Rot31(DWord(0)) Xor (t0 + sKeyTF(K))
            K = K - 1
            t0 = Encryption_Twofish_Fe32(DWord(0), 0)
            t1 = Encryption_Twofish_Fe32(DWord(1), 3)
            t0 = t0 + t1
            DWord(3) = Encryption_Twofish_Rot1(DWord(3) Xor (t0 + t1 + sKeyTF(K)))
            K = K - 1
            DWord(2) = Encryption_Twofish_Rot31(DWord(2)) Xor (t0 + sKeyTF(K))
            K = K - 1
        Else
            t0 = Encryption_Twofish_Fe32(DWord(2), 0)
            t1 = Encryption_Twofish_Fe32(DWord(3), 3)
            t0 = Encryption_Misc_UnsignedAdd(t0, t1)
            DWord(1) = Encryption_Twofish_Rot1(DWord(1) Xor (Encryption_Misc_UnsignedAdd(Encryption_Misc_UnsignedAdd(t0, t1), sKeyTF(K))))
            K = K - 1
            DWord(0) = Encryption_Twofish_Rot31(DWord(0)) Xor (Encryption_Misc_UnsignedAdd(t0, sKeyTF(K)))
            K = K - 1
            t0 = Encryption_Twofish_Fe32(DWord(0), 0)
            t1 = Encryption_Twofish_Fe32(DWord(1), 3)
            t0 = Encryption_Misc_UnsignedAdd(t0, t1)
            DWord(3) = Encryption_Twofish_Rot1(DWord(3) Xor (Encryption_Misc_UnsignedAdd(Encryption_Misc_UnsignedAdd(t0, t1), sKeyTF(K))))
            K = K - 1
            DWord(2) = Encryption_Twofish_Rot31(DWord(2)) Xor (Encryption_Misc_UnsignedAdd(t0, sKeyTF(K)))
            K = K - 1
        End If
    Next

    DWord(0) = DWord(0) Xor sKeyTF(INPUT_WHITEN)
    DWord(1) = DWord(1) Xor sKeyTF(INPUT_WHITEN + 1)
    DWord(2) = DWord(2) Xor sKeyTF(INPUT_WHITEN + 2)
    DWord(3) = DWord(3) Xor sKeyTF(INPUT_WHITEN + 3)

End Sub

Public Sub Encryption_Twofish_DecryptByte(ByteArray() As Byte, Optional Key As String)

Dim Offset As Long
Dim OrigLen As Long
Dim CipherLen As Long
Dim DWord(0 To 5) As Long
Dim CipherWord(0 To 3) As Long

'Set the new key if any was provided

    If (Len(Key) > 0) Then Encryption_Twofish_SetKey Key

    'Get the length of the ciphertext
    CipherLen = UBound(ByteArray) + 1

    'Decrypt the data in 128-bits blocks
    For Offset = 0 To (CipherLen - 1) Step 16
        'Get the next block
        Call CopyMem(DWord(2), ByteArray(Offset), 16)

        'Decrypt the block
        Call Encryption_Twofish_DecryptBlock(DWord())

        'XOR with the previous cipherblock
        DWord(0) = DWord(0) Xor CipherWord(0)
        DWord(1) = DWord(1) Xor CipherWord(1)
        DWord(2) = DWord(2) Xor CipherWord(2)
        DWord(3) = DWord(3) Xor CipherWord(3)

        'Store the current ciphertext to use
        'XOR with the next block plaintext
        Call CopyMem(CipherWord(0), ByteArray(Offset), 16)

        'Store the block
        Call CopyMem(ByteArray(Offset), DWord(0), 16)

    Next

    'Get the size of the original array
    Call CopyMem(OrigLen, ByteArray(8), 4)

    'Make sure OrigLen is a reasonable value,
    'if we used the wrong key the next couple
    'of statements could be dangerous (GPF)
    If (CipherLen - OrigLen > 27) Or (CipherLen - OrigLen < 12) Then
        Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Twofish decryption")
    End If

    'Resize the bytearray to hold only the plaintext
    'and not the extra information added by the
    'encryption routine
    Call CopyMem(ByteArray(0), ByteArray(12), OrigLen)
    ReDim Preserve ByteArray(OrigLen - 1)

End Sub

Public Sub Encryption_Twofish_DecryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_Twofish_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Decrypt the bytearray
    Call Encryption_Twofish_DecryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Encryption_Twofish_DecryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the string to a bytearray

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the array
    Call Encryption_Twofish_DecryptByte(ByteArray(), Key)

    'Return the encrypted data as a string
    Encryption_Twofish_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Static Sub Encryption_Twofish_EncryptBlock(DWord() As Long)

Dim t0 As Long
Dim t1 As Long
Dim K As Long
Dim R As Long

    DWord(0) = DWord(0) Xor sKeyTF(INPUT_WHITEN)
    DWord(1) = DWord(1) Xor sKeyTF(INPUT_WHITEN + 1)
    DWord(2) = DWord(2) Xor sKeyTF(INPUT_WHITEN + 2)
    DWord(3) = DWord(3) Xor sKeyTF(INPUT_WHITEN + 3)

    K = ROUND_SUBKEYS
    For R = 0 To (ROUNDSTF - 1) Step 2
        If (m_RunningCompiled) Then
            'This is the algorithm when run in compiled
            'mode, where VB won't raise overflow errors
            t0 = Encryption_Twofish_Fe32(DWord(0), 0)
            t1 = Encryption_Twofish_Fe32(DWord(1), 3)
            t0 = t0 + t1
            DWord(2) = Encryption_Twofish_Rot1(DWord(2) Xor (t0 + sKeyTF(K)))
            K = K + 1
            DWord(3) = Encryption_Twofish_Rot31(DWord(3)) Xor (t0 + t1 + sKeyTF(K))
            K = K + 1
            t0 = Encryption_Twofish_Fe32(DWord(2), 0)
            t1 = Encryption_Twofish_Fe32(DWord(3), 3)
            t0 = t0 + t1
            DWord(0) = Encryption_Twofish_Rot1(DWord(0) Xor (t0 + sKeyTF(K)))
            K = K + 1
            DWord(1) = Encryption_Twofish_Rot31(DWord(1)) Xor (t0 + t1 + sKeyTF(K))
            K = K + 1
        Else
            'This is the algorithm when running in the IDE,
            'although it's slower it makes the code able
            'to run in the IDE without overflow errors
            t0 = Encryption_Twofish_Fe32(DWord(0), 0)
            t1 = Encryption_Twofish_Fe32(DWord(1), 3)
            t0 = Encryption_Misc_UnsignedAdd(t0, t1)
            DWord(2) = Encryption_Twofish_Rot1(DWord(2) Xor (Encryption_Misc_UnsignedAdd(t0, sKeyTF(K))))
            K = K + 1
            DWord(3) = Encryption_Twofish_Rot31(DWord(3)) Xor (Encryption_Misc_UnsignedAdd(Encryption_Misc_UnsignedAdd(t0, t1), sKeyTF(K)))
            K = K + 1
            t0 = Encryption_Twofish_Fe32(DWord(2), 0)
            t1 = Encryption_Twofish_Fe32(DWord(3), 3)
            t0 = Encryption_Misc_UnsignedAdd(t0, t1)
            DWord(0) = Encryption_Twofish_Rot1(DWord(0) Xor (Encryption_Misc_UnsignedAdd(t0, sKeyTF(K))))
            K = K + 1
            DWord(1) = Encryption_Twofish_Rot31(DWord(1)) Xor (Encryption_Misc_UnsignedAdd(Encryption_Misc_UnsignedAdd(t0, t1), sKeyTF(K)))
            K = K + 1
        End If
    Next

    DWord(2) = DWord(2) Xor sKeyTF(OUTPUT_WHITEN)
    DWord(3) = DWord(3) Xor sKeyTF(OUTPUT_WHITEN + 1)
    DWord(4) = DWord(0) Xor sKeyTF(OUTPUT_WHITEN + 2)
    DWord(5) = DWord(1) Xor sKeyTF(OUTPUT_WHITEN + 3)
    Call CopyMem(DWord(0), DWord(2), 16)

End Sub

Public Sub Encryption_Twofish_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim Offset As Long
Dim OrigLen As Long
Dim CipherLen As Long
Dim DWord(0 To 5) As Long
Dim CipherWord(0 To 3) As Long

'Set the new key if any was provided

    If (Len(Key) > 0) Then Encryption_Twofish_SetKey Key

    'Get the length of the plaintext
    OrigLen = UBound(ByteArray) + 1

    'First we add 12 bytes (4 bytes for the
    'length and 8 bytes for the seed values
    'for the CBC routine), and the ciphertext
    'must be a multiple of 16 bytes
    CipherLen = OrigLen + 12
    If (CipherLen Mod 16 <> 0) Then
        CipherLen = CipherLen + 16 - (CipherLen Mod 16)
    End If
    ReDim Preserve ByteArray(CipherLen - 1)
    Call CopyMem(ByteArray(12), ByteArray(0), OrigLen)

    'Store the length descriptor in bytes [9-12]
    Call CopyMem(ByteArray(8), OrigLen, 4)

    'Store a block of random data in bytes [1-8],
    'these work as seed values for the CBC routine
    'and is used to produce different ciphertext
    'even when encrypting the same data with the
    'same key)
    Call Randomize
    Call CopyMem(ByteArray(0), CLng(2147483647 * Rnd), 4)
    Call CopyMem(ByteArray(4), CLng(2147483647 * Rnd), 4)

    'Encrypt the data in 128-bits blocks
    For Offset = 0 To (CipherLen - 1) Step 16
        'Get the next block
        Call CopyMem(DWord(0), ByteArray(Offset), 16)

        'XOR the plaintext with the previous
        'ciphertext (CBC, Cipher-Block Chaining)
        DWord(0) = DWord(0) Xor CipherWord(0)
        DWord(1) = DWord(1) Xor CipherWord(1)
        DWord(2) = DWord(2) Xor CipherWord(2)
        DWord(3) = DWord(3) Xor CipherWord(3)

        'Encrypt the block
        Call Encryption_Twofish_EncryptBlock(DWord())

        'Store the new block
        Call CopyMem(ByteArray(Offset), DWord(0), 16)

        'Store the cipherblock (for CBC)
        Call CopyMem(CipherWord(0), DWord(0), 16)

    Next

End Sub

Public Sub Encryption_Twofish_EncryptFile(SourceFile As String, DestFile As String, Optional Key As String)

Dim Filenr As Integer
Dim ByteArray() As Byte

'Make sure the source file do exist

    If (Not Encryption_Misc_FileExist(SourceFile)) Then
        Call Err.Raise(vbObjectError, , "Error in Skipjack Encryption_Twofish_EncryptFile procedure (Source file does not exist).")
        Exit Sub
    End If

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray
    Call Encryption_Twofish_EncryptByte(ByteArray(), Key)

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Encryption_Misc_FileExist(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Private Function Encryption_Twofish_Encryption_Twofish_lBSRU(lInput As Long, bShiftBits As Byte) As Long

    If (bShiftBits = 31) Then
        Encryption_Twofish_Encryption_Twofish_lBSRU = -(lInput < 0)
    Else
        Encryption_Twofish_Encryption_Twofish_lBSRU = (((lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits) And Not (&H80000000 + (2 ^ bShiftBits - 2) * 2 ^ (31 - bShiftBits)))
    End If

End Function

Public Function Encryption_Twofish_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the string to a bytearray

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the array
    Call Encryption_Twofish_EncryptByte(ByteArray(), Key)

    'Return the encrypted data as a string
    Encryption_Twofish_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Private Static Function Encryption_Twofish_F32(k64Cnt As Long, X As Long, k32() As Long) As Long

Dim xb(0 To 3) As Byte
Dim Key(0 To 3, 0 To 3) As Byte

    Call CopyMem(xb(0), X, 4)
    Call CopyMem(Key(0, 0), k32(0), 16)

    If ((k64Cnt And 3) = 1) Then
        Encryption_Twofish_F32 = MDS(0, P(0, xb(0)) Xor Key(0, 0)) Xor _
                                 MDS(1, P(0, xb(1)) Xor Key(1, 0)) Xor _
                                 MDS(2, P(1, xb(2)) Xor Key(2, 0)) Xor _
                                 MDS(3, P(1, xb(3)) Xor Key(3, 0))
    Else
        If ((k64Cnt And 3) = 0) Then
            xb(0) = P(1, xb(0)) Xor Key(0, 3)
            xb(1) = P(0, xb(1)) Xor Key(1, 3)
            xb(2) = P(0, xb(2)) Xor Key(2, 3)
            xb(3) = P(1, xb(3)) Xor Key(3, 3)
        End If
        If ((k64Cnt And 3) = 3) Or ((k64Cnt And 3) = 0) Then
            xb(0) = P(1, xb(0)) Xor Key(0, 2)
            xb(1) = P(1, xb(1)) Xor Key(1, 2)
            xb(2) = P(0, xb(2)) Xor Key(2, 2)
            xb(3) = P(0, xb(3)) Xor Key(3, 2)
        End If
        Encryption_Twofish_F32 = MDS(0, P(0, P(0, xb(0)) Xor Key(0, 1)) Xor Key(0, 0)) Xor _
                                 MDS(1, P(0, P(1, xb(1)) Xor Key(1, 1)) Xor Key(1, 0)) Xor _
                                 MDS(2, P(1, P(0, xb(2)) Xor Key(2, 1)) Xor Key(2, 0)) Xor _
                                 MDS(3, P(1, P(1, xb(3)) Xor Key(3, 1)) Xor Key(3, 0))
    End If

End Function

Private Static Function Encryption_Twofish_Fe32(X As Long, R As Long) As Long

Dim xb(0 To 3) As Byte

'Extract the byte sequence

    Call CopyMem(xb(0), X, 4)

    'Calculate the FE32 function
    Encryption_Twofish_Fe32 = sBoxTF(2 * xb(R Mod 4)) Xor _
                              sBoxTF(2 * xb((R + 1) Mod 4) + 1) Xor _
                              sBoxTF(&H200 + 2 * xb((R + 2) Mod 4)) Xor _
                              sBoxTF(&H200 + 2 * xb((R + 3) Mod 4) + 1)

End Function

Private Static Sub Encryption_Twofish_KeyCreate(K() As Byte, KeyLength As Long)

Dim i As Long
Dim lA As Long
Dim lB As Long
Dim b(3) As Byte
Dim k64Cnt As Long
Dim k32e(3) As Long
Dim k32o(3) As Long
Dim subkeyCnt As Long
Dim sBoxTFKey(3) As Long
Dim Key(0 To 3, 0 To 3) As Byte

Const SK_STEP = &H2020202
Const SK_BUMP = &H1010101
Const SK_ROTL = 9

    k64Cnt = KeyLength \ 8
    subkeyCnt = ROUND_SUBKEYS + 2 * ROUNDSTF

    For i = 0 To IIf(KeyLength < 32, KeyLength \ 8 - 1, 3)
        Call CopyMem(k32e(i), K(i * 8), 4)
        Call CopyMem(k32o(i), K(i * 8 + 4), 4)
        sBoxTFKey(KeyLength \ 8 - 1 - i) = Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(Encryption_Twofish_RS_Rem(k32o(i))))) Xor k32e(i)))))
    Next

    ReDim sKeyTF(subkeyCnt)
    For i = 0 To ((subkeyCnt / 2) - 1)
        lA = Encryption_Twofish_F32(k64Cnt, i * SK_STEP, k32e)
        lB = Encryption_Twofish_F32(k64Cnt, i * SK_STEP + SK_BUMP, k32o)
        lB = Encryption_Twofish_lBSL(lB, 8) Or Encryption_Twofish_Encryption_Twofish_lBSRU(lB, 24)
        If (m_RunningCompiled) Then
            lA = lA + lB
        Else
            lA = Encryption_Misc_UnsignedAdd(lA, lB)
        End If
        sKeyTF(2 * i) = lA
        If (m_RunningCompiled) Then
            lA = lA + lB
        Else
            lA = Encryption_Misc_UnsignedAdd(lA, lB)
        End If
        sKeyTF(2 * i + 1) = Encryption_Twofish_lBSL(lA, SK_ROTL) Or Encryption_Twofish_Encryption_Twofish_lBSRU(lA, 32 - SK_ROTL)
    Next

    Call CopyMem(Key(0, 0), sBoxTFKey(0), 16)

    For i = 0 To 255
        If ((k64Cnt And 3) = 1) Then
            sBoxTF(2 * i) = MDS(0, P(0, i) Xor Key(0, 0))
            sBoxTF(2 * i + 1) = MDS(1, P(0, i) Xor Key(1, 0))
            sBoxTF(&H200 + 2 * i) = MDS(2, P(1, i) Xor Key(2, 0))
            sBoxTF(&H200 + 2 * i + 1) = MDS(3, P(1, i) Xor Key(3, 0))
        Else
            b(0) = i
            b(1) = i
            b(2) = i
            b(3) = i
            If ((k64Cnt And 3) = 0) Then
                b(0) = P(1, b(0)) Xor Key(0, 3)
                b(1) = P(0, b(1)) Xor Key(1, 3)
                b(2) = P(0, b(2)) Xor Key(2, 3)
                b(3) = P(1, b(3)) Xor Key(3, 3)
            End If
            If ((k64Cnt And 3) = 3) Or ((k64Cnt And 3) = 0) Then '(exception = True) Then
                b(0) = P(1, b(0)) Xor Key(0, 2)
                b(1) = P(1, b(1)) Xor Key(1, 2)
                b(2) = P(0, b(2)) Xor Key(2, 2)
                b(3) = P(0, b(3)) Xor Key(3, 2)
            End If
            sBoxTF(2 * i) = MDS(0, P(0, P(0, b(0)) Xor Key(0, 1)) Xor Key(0, 0))
            sBoxTF(2 * i + 1) = MDS(1, P(0, P(1, b(1)) Xor Key(1, 1)) Xor Key(1, 0))
            sBoxTF(&H200 + 2 * i) = MDS(2, P(1, P(0, b(2)) Xor Key(2, 1)) Xor Key(2, 0))
            sBoxTF(&H200 + 2 * i + 1) = MDS(3, P(1, P(1, b(3)) Xor Key(3, 1)) Xor Key(3, 0))
        End If
    Next

End Sub

Private Function Encryption_Twofish_lBSL(ByRef lInput As Long, ByRef bShiftBits As Byte) As Long

    Encryption_Twofish_lBSL = (lInput And (2 ^ (31 - bShiftBits) - 1)) * 2 ^ bShiftBits
    If (lInput And 2 ^ (31 - bShiftBits)) = 2 ^ (31 - bShiftBits) Then Encryption_Twofish_lBSL = (Encryption_Twofish_lBSL Or &H80000000)

End Function

Private Function Encryption_Twofish_lBSR(ByRef lInput As Long, ByRef bShiftBits As Byte) As Long

    If (bShiftBits = 31) Then
        If (lInput < 0) Then Encryption_Twofish_lBSR = &HFFFFFFFF Else Encryption_Twofish_lBSR = 0
    Else
        Encryption_Twofish_lBSR = (lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits
    End If

End Function

Private Static Function Encryption_Twofish_LFSR1(ByRef X As Long) As Long

    Encryption_Twofish_LFSR1 = Encryption_Twofish_lBSR(X, 1) Xor ((X And 1) * GF256_FDBK_2)

End Function

Private Static Function Encryption_Twofish_LFSR2(ByRef X As Long) As Long

    Encryption_Twofish_LFSR2 = Encryption_Twofish_lBSR(X, 2) Xor ((X And &H2) / &H2 * GF256_FDBK_2) Xor ((X And &H1) * GF256_FDBK_4)

End Function

Private Static Function Encryption_Twofish_Rot1(Value As Long) As Long

Dim Temp As Byte
Dim X(0 To 3) As Byte

    Call CopyMem(X(0), Value, 4)

    Temp = X(0)
    X(0) = (X(0) \ 2) Or ((X(1) And 1) * 128)
    X(1) = (X(1) \ 2) Or ((X(2) And 1) * 128)
    X(2) = (X(2) \ 2) Or ((X(3) And 1) * 128)
    X(3) = (X(3) \ 2) Or ((Temp And 1) * 128)

    Call CopyMem(Encryption_Twofish_Rot1, X(0), 4)

End Function

Private Static Function Encryption_Twofish_Rot31(Value As Long) As Long

Dim Temp As Byte
Dim X(0 To 3) As Byte

    Call CopyMem(X(0), Value, 4)

    Temp = X(3)
    X(3) = ((X(3) And 127) * 2) Or -CBool(X(2) And 128)
    X(2) = ((X(2) And 127) * 2) Or -CBool(X(1) And 128)
    X(1) = ((X(1) And 127) * 2) Or -CBool(X(0) And 128)
    X(0) = ((X(0) And 127) * 2) Or -CBool(Temp And 128)

    Call CopyMem(Encryption_Twofish_Rot31, X(0), 4)

End Function

Private Static Function Encryption_Twofish_RS_Rem(X As Long) As Long

Dim b As Long
Dim g2 As Long
Dim g3 As Long

    b = (Encryption_Twofish_Encryption_Twofish_lBSRU(X, 24) And &HFF)
    g2 = ((Encryption_Twofish_lBSL(b, 1) Xor (b And &H80) / &H80 * &H14D) And &HFF)
    g3 = (Encryption_Twofish_Encryption_Twofish_lBSRU(b, 1) Xor ((b And &H1) * Encryption_Twofish_Encryption_Twofish_lBSRU(&H14D, 1)) Xor g2)
    Encryption_Twofish_RS_Rem = Encryption_Twofish_lBSL(X, 8) Xor Encryption_Twofish_lBSL(g3, 24) Xor Encryption_Twofish_lBSL(g2, 16) Xor Encryption_Twofish_lBSL(g3, 8) Xor b

End Function

Public Sub Encryption_Twofish_SetKey(New_Value As String, Optional ByVal MinKeyLength As TWOFISHKEYLENGTH)

Dim KeyLength As Long
Dim Key() As Byte

'Convert the key into a bytearray

    KeyLength = Len(New_Value) * 8
    Key() = StrConv(New_Value, vbFromUnicode)

    'Resize the key array if it is too small
    If (KeyLength < MinKeyLength) Then
        ReDim Preserve Key(MinKeyLength \ 8 - 1)
        KeyLength = MinKeyLength
    End If

    'The key array can only be of certain sizes,
    'if the size is invalid resize to the closes
    'size (preferably by making it larger)
    If (KeyLength > 192) Then
        ReDim Preserve Key(31)
        KeyLength = 256
    ElseIf (KeyLength > 128) Then
        ReDim Preserve Key(23)
        KeyLength = 192
    ElseIf (KeyLength > 64) Then
        ReDim Preserve Key(15)
        KeyLength = 128
    ElseIf (KeyLength > 32) Then
        ReDim Preserve Key(7)
        KeyLength = 64
    Else
        ReDim Preserve Key(3)
        KeyLength = 32
    End If

    'Create the key-dependant sboxes
    Call Encryption_Twofish_KeyCreate(Key, KeyLength \ 8)

End Sub
