Attribute VB_Name = "AllFilesInFolder"
Option Explicit

Private Sub AddItem2Array1D(ByRef VarArray As Variant, ByVal VarValue As Variant)

Dim i  As Long
Dim iVarType As Integer

    iVarType = VarType(VarArray) - 8192
    i = UBound(VarArray)

    Select Case iVarType

    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte

        If VarArray(0) = 0 Then
            i = 0
        Else
            i = i + 1
        End If

    Case vbDate

        If VarArray(0) = "00:00:00" Then
            i = 0
        Else
            i = i + 1
        End If

    Case vbString

        If VarArray(0) = vbNullString Then
            i = 0
        Else
            i = i + 1
        End If

    Case vbBoolean

        If VarArray(0) = False Then
            i = 0
        Else
            i = i + 1
        End If

    Case Else

    End Select

    ReDim Preserve VarArray(i)
    VarArray(i) = VarValue

End Sub

Public Function AllFilesInFolders(ByVal sFolderPath As String, Optional bWithSubFolders As Boolean = False) As String()

Dim sTemp As String
Dim sDirIn As String
Dim i As Integer, j As Integer

    ReDim sFilelist(0) As String, sSubFolderList(0) As String, sToProcessFolderList(0) As String
    sDirIn = sFolderPath
    If Not (Right$(sDirIn, 1) = "\") Then sDirIn = sDirIn & "\"
    
    On Error Resume Next
    
        sTemp = Dir$(sDirIn & "*.*")
        Do While LenB(sTemp) <> 0
            AddItem2Array1D sFilelist(), sDirIn & sTemp
            sTemp = Dir
        Loop
        If bWithSubFolders Then

            sTemp = Dir$(sDirIn & "*.*", vbDirectory)
            Do While LenB(sTemp) <> 0

                If sTemp <> "." Then
                    If sTemp <> ".." Then
                        If (GetAttr(sDirIn & sTemp) And vbDirectory) = vbDirectory Then AddItem2Array1D sToProcessFolderList, sDirIn & sTemp
                    End If
                End If
                sTemp = Dir
                
            Loop

            If UBound(sToProcessFolderList) > 0 Or UBound(sToProcessFolderList) = 0 And LenB(sToProcessFolderList(0)) <> 0 Then
                For i = 0 To UBound(sToProcessFolderList)
                    sSubFolderList = AllFilesInFolders(sToProcessFolderList(i), bWithSubFolders)
                    If UBound(sSubFolderList) > 0 Or UBound(sSubFolderList) = 0 And LenB(sSubFolderList(0)) <> 0 Then
                        For j = 0 To UBound(sSubFolderList)
                            AddItem2Array1D sFilelist(), sSubFolderList(j)
                        Next
                    End If
                Next
            End If

        End If

        AllFilesInFolders = sFilelist
        
    On Error GoTo 0

End Function
