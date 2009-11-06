Attribute VB_Name = "FileIO"
Option Explicit

Public Function Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional ByVal DontLog As Byte = 0) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************


    Var_Get = Space$(1000)
    getprivateprofilestring Main, Var, vbNullString, Var_Get, 1000, File
    Var_Get = RTrim$(Var_Get)
    If LenB(Var_Get) <> 0 Then Var_Get = Left$(Var_Get, Len(Var_Get) - 1)
    

End Function

Public Sub Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, File

End Sub
