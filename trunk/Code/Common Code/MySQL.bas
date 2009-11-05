Attribute VB_Name = "MySQL"
Option Explicit

'Database connection information
Public Const DB_User As String = "root"         'The database username - (default "root")
Public Const DB_Pass As String = "test"         'Password to your database for the corresponding username
Public Const DB_Name As String = "vbgore"       'Name of the table in the database (default "vbgore")
Public Const DB_Host As String = "127.0.0.1"    'IP of the database server - use localhost if hosted locally! Only host remotely for multiple servers!
Public Const DB_Port As String = "3306"         'Port of the database (default "3306")
Public Const DB_PasswordKey As String = "asd23409AFkj123098DSfj"    'Key for encrypting passwords in database (enter no values for no encryption)

'Connection objects
Public DB_Conn As ADODB.Connection
Public DB_RS As ADODB.Recordset

Public Sub MySQL_Init()

    On Error GoTo ErrOut

    'Create the database connection object
    Set DB_Conn = New ADODB.Connection
    Set DB_RS = New ADODB.Recordset
    
    'Create the connection
    DB_Conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & DB_Host & _
        ";DATABASE=" & DB_Name & ";PORT=" & DB_Port & ";UID=" & DB_User & ";PWD=" & DB_Pass & ";OPTION=3"
    DB_Conn.CursorLocation = adUseClient
    DB_Conn.Open
    
    Exit Sub
    
ErrOut:

    'Could not connect to the database
    MsgBox "Error connecting to the MySQL database. Please make sure you have MySQL installed (www.mysql.com) and running!" & vbCrLf & _
        "Also make sure your connection variables are correct (found in the MySQL module's declares section)." & vbCrLf & _
        "If you have your database installed and running, make sure you have executed the database dump on the 'vbgore' table." & vbCrLf & _
        "The database dump can be found in the '_Database dump' file. Select 'Execute batch file' (or something similar) on your 'vbGORE' table.", vbOKOnly
    Unload frmMain

End Sub

Public Sub MySQL_RemoveOnline()
On Error Resume Next

    'Remove the online flag from all the users (recommended for server loading in case of a crash)
    DB_RS.Open "SELECT * FROM users WHERE `online`='1'", DB_Conn, adOpenStatic, adLockOptimistic
    If DB_RS.EOF = False Then
        Do While DB_RS.EOF = False
            DB_RS!online = 0
            DB_RS.MoveNext
        Loop
        DB_RS.Update
    End If
    DB_RS.Close
    
End Sub

Public Sub MySQL_Optimize()

    'Optimize the database tables
    DB_Conn.Execute "OPTIMIZE TABLE mail, mail_lastid, npcs, objects, quests, users"

End Sub