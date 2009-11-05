Attribute VB_Name = "General"
'**************************************************************************************
'**************************************************************************************
'***                                   HOW TO USE                                   ***
'**************************************************************************************
'**************************************************************************************
'*** To use this conversion tool, you must place the old variable formats in the    ***
'*** OldData module so it can load correctly. Place the variable format you wish to ***
'*** change to in the NewData module. Note that changes may need to take place in-  ***
'*** between the loading and saving process, so use this if needed.                 ***
'***                                                                                ***
'*** Once you finish a conversion, make sure you copy the variables from NewData    ***
'*** to OldData, since they are now your most up-to-date ones. :)                   ***
'**************************************************************************************
'**************************************************************************************
Option Explicit
Private NumMail As Long
Private OldMail() As OldMailData.MailData
Private NewMail() As NewMailData.MailData
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Sub OldToNew()
Dim i As Long
Dim j As Long

    'Convert the old to the new - because the types arn't the same (if they were, we wouldn't
    ' even be needing to use this code!), we have to copy every piece of information over
    ' variable by variable. Use this section to also do algorithms to set the variables, too,
    ' such as if you wanted to add in a Value for mail, and wanted to make the value
    ' a combination of a bunch of the Mail variables.
    For i = 1 To NumMail
        NewMail(i).Message = OldMail(i).Message
        NewMail(i).New = OldMail(i).New
        For j = 1 To UBound(OldMail(i).Obj())
            NewMail(i).Obj(j).Amount = OldMail(i).Obj(j).Amount
            NewMail(i).Obj(j).ObjIndex = OldMail(i).Obj(j).ObjIndex
        Next j
        NewMail(i).RecieveDate = OldMail(i).RecieveDate
        NewMail(i).Subject = OldMail(i).Subject
        NewMail(i).WriterName = OldMail(i).WriterName
    Next i

End Sub

Sub Main()
Dim FileNum As Byte

    'Load the file paths
    InitFilePaths
    
    'Get the number of mail
    FileNum = FreeFile
    Open MailPath & "Count.mail" For Binary As FileNum
        Get FileNum, , NumMail
    Close FileNum
    
    'Resize our arrays
    ReDim OldMail(0 To NumMail)
    ReDim NewMail(0 To NumMail)

    'Load the mail
    Load_Mail
    
    'Save the backups
    Save_Mail_Backup
    
    'Convert the old variables to the new
    OldToNew
    
    'Save the mail
    Save_Mail
    
    'Done
    MsgBox "Mail conversion successful!" & vbCrLf & _
           "Old type size: " & Len(OldMail(0)) & vbCrLf & _
           "New type size: " & Len(NewMail(0)) & vbCrLf & vbCrLf & _
           "Backups were made and placed in the following folder: " & vbCrLf & _
           MailPath & "Backups\ folder!" & vbCrLf & vbCrLf & _
           "Be sure to copy your NewMailData to OldMailData so the next time it will load properly!", vbOKOnly

End Sub

Sub Load_Mail()
Dim Mail As Long
Dim FileNum As Byte
Dim LengthI As Integer
Dim LengthB As Byte

    'Fill Mail List
    FileNum = FreeFile
    For Mail = 1 To NumMail
        Open MailPath & Mail & ".mail" For Binary As FileNum
            Get FileNum, , LengthI
            OldMail(Mail).Message = Space$(LengthI)
            Get FileNum, , OldMail(Mail).Message
        
            Get FileNum, , LengthB
            OldMail(Mail).Subject = Space$(LengthB)
            Get FileNum, , OldMail(Mail).Subject
        
            Get FileNum, , LengthB
            OldMail(Mail).WriterName = Space$(LengthB)
            Get FileNum, , OldMail(Mail).WriterName
        
            Get FileNum, , OldMail(Mail).New
            Get FileNum, , OldMail(Mail).Obj
            Get FileNum, , OldMail(Mail).RecieveDate
        Close FileNum
    Next Mail
    
End Sub

Sub Save_Mail()
Dim Mail As Long
Dim FileNum As Byte
Dim LengthI As Integer
Dim LengthB As Byte
    
    'Get the number of mail
    FileNum = FreeFile

    'Fill Mail List
    For Mail = 1 To NumMail
        Open MailPath & Mail & ".mail" For Binary As FileNum
            LengthI = Len(OldMail(Mail).Message)
            Put FileNum, , LengthI
            Put FileNum, , OldMail(Mail).Message
        
            LengthB = Len(OldMail(Mail).Subject)
            Put FileNum, , LengthB
            Put FileNum, , OldMail(Mail).Subject
        
            LengthB = Len(OldMail(Mail).WriterName)
            Put FileNum, , LengthB
            Put FileNum, , OldMail(Mail).WriterName
        
            Put FileNum, , OldMail(Mail).New
            Put FileNum, , OldMail(Mail).Obj
            Put FileNum, , OldMail(Mail).RecieveDate
        Close FileNum
    Next Mail

End Sub

Sub Save_Mail_Backup()
Dim Mail As Long
Dim FileNum As Byte
Dim LengthI As Integer
Dim LengthB As Byte

    'Get the number of mail
    FileNum = FreeFile
    
    'Ensure the path exists
    MakeSureDirectoryPathExists MailPath & "Backup\"

    'Fill Mail List
    For Mail = 1 To NumMail
        Open MailPath & "Backup\" & Mail & ".mail" For Binary As FileNum
            LengthI = Len(OldMail(Mail).Message)
            Put FileNum, , LengthI
            Put FileNum, , OldMail(Mail).Message
        
            LengthB = Len(OldMail(Mail).Subject)
            Put FileNum, , LengthB
            Put FileNum, , OldMail(Mail).Subject
        
            LengthB = Len(OldMail(Mail).WriterName)
            Put FileNum, , LengthB
            Put FileNum, , OldMail(Mail).WriterName
        
            Put FileNum, , OldMail(Mail).New
            Put FileNum, , OldMail(Mail).Obj
            Put FileNum, , OldMail(Mail).RecieveDate
        Close FileNum
    Next Mail

End Sub

