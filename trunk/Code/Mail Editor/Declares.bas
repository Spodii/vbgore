Attribute VB_Name = "Declares"
Option Explicit

Public FilePath As String

Public Const MaxMailObjs As Byte = 10       'How many objects can be attached to a message maximum

'Holds info about a object
Public Type Obj
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type

'Mail information
Type MailData
    Subject As String
    WriterName As String
    RecieveDate As Date
    Message As String
    New As Byte
    Obj(1 To MaxMailObjs) As Obj
End Type

Public MailData As MailData

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:56)  Decl: 26  Code: 0  Total: 26 Lines
':) CommentOnly: 2 (7.7%)  Commented: 3 (11.5%)  Empty: 5 (19.2%)  Max Logic Depth: 1
