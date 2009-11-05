Attribute VB_Name = "General"
Option Explicit

Public Sub Main()

    InitFilePaths

    'Load the user
    If Command$ <> "" Then
        FilePath = Mid$(Command$, 2, Len(Command$) - 2) 'Retrieve the filepath from Command$ and crop off the "'s
        LoadUser FilePath
    End If

    frmMain.Show
    
End Sub

Public Sub FillInInformation()

Dim i As Integer

'Fill in all the information on the form

    With frmMain
        .NameTxt.Text = UserChar.Name
        .PassTxt.Text = UserChar.Password
        .GoldTxt.Text = UserChar.Gold
        .MapTxt.Text = UserChar.Pos.Map
        .XTxt.Text = UserChar.Pos.X
        .YTxt.Text = UserChar.Pos.Y
        .HairTxt.Text = UserChar.Char.Hair
        .HeadTxt.Text = UserChar.Char.Head
        .BodyTxt.Text = UserChar.Char.Body
        .WeaponTxt.Text = UserChar.Char.Weapon
        .HeadingTxt.Text = UserChar.Char.Heading
        .HeadHeadingTxt.Text = UserChar.Char.HeadHeading
        .DescTxt.Text = UserChar.Desc
        .CompletedQuestsTxt.Text = UserChar.CompletedQuests
        .QuestTxt.Text = UserChar.Quest
        .ArmorSlotTxt.Text = UserChar.ArmorEqpSlot
        .WeaponSlotTxt.Text = UserChar.WeaponEqpSlot
        For i = 1 To .StatTxt.UBound
            If i <= NumStats Then
                .StatTxt(i).Text = UserChar.Stats.BaseStat(i)
                .StatTxt(i).Enabled = True
            Else
                .StatTxt(i).Text = "N/A"
                .StatTxt(i).Enabled = False
            End If
        Next i
        For i = 1 To .InventoryTxt.UBound
            If i <= MAX_INVENTORY_SLOTS Then
                .InventoryTxt(i).Text = UserChar.Object(i).ObjIndex
                .InventoryTxt(i).Enabled = True
                .AmountTxt(i).Text = UserChar.Object(i).Amount
                .AmountTxt(i).Enabled = True
                .EquiptedTxt(i).Text = UserChar.Object(i).Equipped
                .EquiptedTxt(i).Enabled = True
            Else
                .InventoryTxt(i).Text = "N/A"
                .InventoryTxt(i).Enabled = False
                .AmountTxt(i).Text = "0"
                .AmountTxt(i).Enabled = False
                .EquiptedTxt(i).Text = "X"
                .EquiptedTxt(i).Enabled = False
            End If
        Next i
        For i = 1 To .KnownSkillTxt.UBound
            If i <= NumSkills Then
                .KnownSkillTxt(i).Text = UserChar.KnownSkills(i)
                .KnownSkillTxt(i).Enabled = True
            Else
                .KnownSkillTxt(i).Text = "X"
                .KnownSkillTxt(i).Enabled = False
            End If
        Next i

    End With

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String

'*****************************************************************
'Gets a variable from a text file
'*****************************************************************

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found

    szReturn = ""

    sSpaces = Space$(1000) ' This tells the computer how long the longest string can be. If you want, you can change the number 75 to any number you wish

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), file

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Function Engine_FileExist(file As String, FileType As VbFileAttribute) As Boolean

'*****************************************************************
'Checks to see if a file exists
'*****************************************************************

    Engine_FileExist = (Dir$(file, FileType) <> "")

End Function

Public Sub LoadUser(ByVal FilePathx As String)

Dim FileNum As Byte
Dim a As String
Dim s() As String

    'Make sure the file exists
    If Engine_FileExist(FilePath, vbNormal) = False Then
        MsgBox "Error! Character file (" & FilePath & ") does not exist!", vbOKOnly
        Exit Sub
    End If
    
    FilePath = FilePathx

    FileNum = FreeFile

    'Create the class
    Set UserChar.Stats = New UserStats

    'Get the name
    s = Split(FilePath, "\")
    UserChar.Name = Left$(s(UBound(s)), Len(s(UBound(s))) - 4)

    'Load the user character
    UserChar.Password = GetVar(FilePath & ".pass", "A", "A")
    Open FilePath For Binary As FileNum
    Get FileNum, , UserChar.ArmorEqpSlot
    Get FileNum, , UserChar.Char
    Get FileNum, , UserChar.CompletedQuests
    Get FileNum, , UserChar.Desc
    Get FileNum, , UserChar.Object
    Get FileNum, , UserChar.Pos
    Get FileNum, , UserChar.Quest
    Get FileNum, , UserChar.QuestRequirements
    Get FileNum, , UserChar.Skills
    Get FileNum, , UserChar.WeaponEqpSlot
    Get FileNum, , UserChar.WeaponType
    Get FileNum, , UserChar.MailID
    Get FileNum, , UserChar.KnownSkills
    UserChar.Stats.LoadClass FileNum

    'Equipt items
    If UserChar.WeaponEqpSlot > 0 Then UserChar.WeaponEqpObjIndex = UserChar.Object(UserChar.WeaponEqpSlot).ObjIndex
    If UserChar.ArmorEqpSlot > 0 Then UserChar.ArmorEqpObjIndex = UserChar.Object(UserChar.ArmorEqpSlot).ObjIndex

    'Fill in all the information
    FillInInformation

End Sub

Sub SaveUser(ByVal FilePath As String)

'*****************************************************************
'Saves a user's data to a .chr file
'*****************************************************************

Dim FileNum As Byte

'Save the user character

    FileNum = FreeFile
    WriteVar FilePath & ".pass", "A", "A", UserChar.Password
    Open FilePath & ".ip" For Append Shared As FileNum
    Print #FileNum, UserChar.IP
    Close FileNum
    Open FilePath For Binary As FileNum
    Put FileNum, , UserChar.ArmorEqpSlot
    Put FileNum, , UserChar.Char
    Put FileNum, , UserChar.CompletedQuests
    Put FileNum, , UserChar.Desc
    Put FileNum, , UserChar.Object
    Put FileNum, , UserChar.Pos
    Put FileNum, , UserChar.Quest
    Put FileNum, , UserChar.QuestRequirements
    Put FileNum, , UserChar.Skills
    Put FileNum, , UserChar.WeaponEqpSlot
    Put FileNum, , UserChar.WeaponType
    Put FileNum, , UserChar.MailID
    Put FileNum, , UserChar.KnownSkills
    UserChar.Stats.SaveClass FileNum

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)

'*****************************************************************
'Writes a var to a text file
'*****************************************************************

    writeprivateprofilestring Main, Var, Value, file

End Sub

':) Ulli's VB Code Formatter V2.19.5 (2006-Jul-31 17:43)  Decl: 1  Code: 169  Total: 170 Lines
':) CommentOnly: 15 (8.8%)  Commented: 3 (1.8%)  Empty: 31 (18.2%)  Max Logic Depth: 4
