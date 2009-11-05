Attribute VB_Name = "OldMailData"
'*******************************************************************************
'Place in this module the old variables (so that way it can load correctly)
'*******************************************************************************
Option Explicit

Public Const MaxMailObjs As Byte = 10       'How many objects can be attached to a message maximum
Public Type Obj 'Holds info about a object
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type
Type MailData   'Mailing system information
    Subject As String
    WriterName As String
    RecieveDate As Date
    Message As String
    New As Byte
    Obj(1 To MaxMailObjs) As OldMailData.Obj
End Type

