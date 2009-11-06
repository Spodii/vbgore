Attribute VB_Name = "FileIO"
Function Engine_FileExist(File As String, FileType As VbFileAttribute) As Boolean
'*****************************************************************
'Checks to see if a file exists
'*****************************************************************
    Engine_FileExist = (Dir$(File, FileType) <> "")
End Function

Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Var_Get = Space$(1000)
    getprivateprofilestring Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
End Function
