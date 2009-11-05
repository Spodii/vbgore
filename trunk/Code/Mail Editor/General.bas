Attribute VB_Name = "General"
Option Explicit

Sub Main()

    InitFilePaths

    'Load the first message
    If Command$ = "" Then
        If Engine_FileExist(App.Path & "\Mail\1.mail", vbNormal) Then LoadMail App.Path & "\Mail\1.mail", MailData
    Else
        FilePath = Mid$(Command$, 2, Len(Command$) - 2) 'Retrieve the filepath from Command$ and crop off the "'s
        LoadMail FilePath, MailData
    End If

    'Fill in information
    FillInInformation
    
    frmMain.Show

End Sub

Function Engine_FileExist(file As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(file, FileType) <> "")

End Function

Sub FillInInformation()

Dim i As Byte

'Fill in the form's information

    With frmMain
        .SubjectTxt.Text = MailData.Subject
        .WriterTxt.Text = MailData.WriterName
        .MessageTxt.Text = MailData.Message
        .NewTxt.Text = MailData.New
        .DateTxt.Text = MailData.RecieveDate
        For i = 1 To .ItemTxt.UBound
            If i > MaxMailObjs Then
                .ItemTxt(i).Text = "N/A"
                .ItemTxt(i).Enabled = False
                .AmountTxt(i).Text = "X"
                .AmountTxt(i).Enabled = False
            Else
                .ItemTxt(i).Text = MailData.Obj(i).ObjIndex
                .ItemTxt(i).Enabled = True
                .AmountTxt(i).Text = MailData.Obj(i).Amount
                .AmountTxt(i).Enabled = True
            End If
        Next i
    End With

End Sub

Sub LoadMail(ByVal MailPath As String, ByRef MailHandler As MailData)

Dim FileNum As Byte
Dim LengthI As Integer
Dim LengthB As Byte

    'Make sure the file exists
    If Engine_FileExist(MailPath, vbNormal) = False Then
        MsgBox "Error! Mail file (" & MailPath & ") not found!", vbOKOnly
        Exit Sub
    End If
    
    FilePath = MailPath
    
    'Change the caption
    frmMain.Caption = "Mail Editor - Mail: " & Val(Right$(MailPath, Len(MailPath) - Len(App.Path & "\Mail\")))

    'Open the file and retrieve the data
    FileNum = FreeFile
    Open MailPath For Binary As FileNum

    Get FileNum, , LengthI
    MailHandler.Message = Space$(LengthI)
    Get FileNum, , MailHandler.Message

    Get FileNum, , LengthB
    MailHandler.Subject = Space$(LengthB)
    Get FileNum, , MailHandler.Subject

    Get FileNum, , LengthB
    MailHandler.WriterName = Space$(LengthB)
    Get FileNum, , MailHandler.WriterName

    Get FileNum, , MailHandler.New
    Get FileNum, , MailHandler.Obj
    Get FileNum, , MailHandler.RecieveDate
    Close FileNum

End Sub

Sub SaveMail(ByVal MailPath As String, ByRef MailData As MailData)

Dim FileNum As Byte
Dim LengthI As Integer  'Length of a string as an integer
Dim LengthB As Byte     'Length of a string as a byte

'Open the file and save the data

    FileNum = FreeFile
    Open MailPath For Binary As FileNum
    LengthI = Len(MailData.Message)
    Put FileNum, , LengthI
    Put FileNum, , MailData.Message

    LengthB = Len(MailData.Subject)
    Put FileNum, , LengthB
    Put FileNum, , MailData.Subject

    LengthB = Len(MailData.WriterName)
    Put FileNum, , LengthB
    Put FileNum, , MailData.WriterName

    Put FileNum, , MailData.New
    Put FileNum, , MailData.Obj
    Put FileNum, , MailData.RecieveDate

    Close FileNum

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:56)  Decl: 1  Code: 92  Total: 93 Lines
':) CommentOnly: 3 (3.2%)  Commented: 2 (2.2%)  Empty: 22 (23.7%)  Max Logic Depth: 4
