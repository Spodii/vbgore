Attribute VB_Name = "MySQL"
Option Explicit

'Database connection information (values specified in /ServerData/Server.ini)
Public DB_User As String    'The database username - (default "root")
Public DB_Pass As String    'Password to your database for the corresponding username
Public DB_Name As String    'Name of the table in the database (default "vbgore")
Public DB_Host As String    'IP of the database server - use localhost if hosted locally! Only host remotely for multiple servers!
Public DB_Port As Integer   'Port of the database (default "3306")

'Change these values to update the database when the value changes during gameplay
'Most of these values will automatically be set during loading/saving a character (except _Online)
'0 is for false, 1 is for true
Public Const MySQLUpdate_UserMap As Boolean = True

'Connection objects
Public DB_Conn As ADODB.Connection
Public DB_RS As ADODB.Recordset

Public Sub MySQL_Init()

    On Error GoTo ErrOut

    'Create the database connection object
    Set DB_Conn = New ADODB.Connection
    Set DB_RS = New ADODB.Recordset
    
    'Get the variables
    DB_User = Trim$(Var_Get(ServerDataPath & "Server.ini", "MYSQL", "User"))
    DB_Pass = Trim$(Var_Get(ServerDataPath & "Server.ini", "MYSQL", "Password"))
    DB_Name = Trim$(Var_Get(ServerDataPath & "Server.ini", "MYSQL", "Database"))
    DB_Host = Trim$(Var_Get(ServerDataPath & "Server.ini", "MYSQL", "Host"))
    DB_Port = Val(Var_Get(ServerDataPath & "Server.ini", "MYSQL", "Port"))
    
    'Create the connection
    DB_Conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & DB_Host & _
        ";DATABASE=" & DB_Name & ";PORT=" & DB_Port & ";UID=" & DB_User & ";PWD=" & DB_Pass & ";OPTION=3"
    DB_Conn.CursorLocation = adUseClient
    DB_Conn.Open
    
    On Error GoTo 0
    
    Exit Sub
    
ErrOut:

    'Could not connect to the database
    MsgBox "Error connecting to the MySQL database. Please make sure you have MySQL 5.0 running, and that you have ODBC v3.51!" & vbNewLine & _
       "Also make sure your connection variables are correct (found in \ServerData\Server.ini)." & vbNewLine & _
       "If you have your database installed and running, make sure you have executed the database dump on the 'vbgore' database." & vbNewLine & _
       "The database dump can be found in the '_Database dump' folder. Select 'Execute batch file' (or something similar) on your 'vbgore' database." & vbNewLine & vbNewLine & _
       "If you feel you have done all of these steps, confirm the following has been done:" & vbNewLine & _
       " - You have followed ALL of the steps on the MySQL Setup page on the site" & vbNewLine & _
       " - MySQL is running and you can connect to it through a GUI such as SQLyog" & vbNewLine & _
       " - You have imported the vbgore.sql file into the database and can see the information through the GUI" & vbNewLine & _
       " - You have version 5.0 of MySQL and 3.51 of ODBC being used" & vbNewLine & vbNewLine & _
       "If you are positive you have done all of the above, freel free to ask for help on the vbGORE forums.", vbOKOnly
    Server_Unload

End Sub

Public Sub MySQL_RemoveOnline()

    'Remove the online flag from all the users (recommended for server loading in case of a crash)
    If ServerID > 0 Then    'This prevents the map editor making this call
        DB_RS.Open "SELECT * FROM users WHERE `server`='" & ServerID & "'", DB_Conn, adOpenStatic, adLockOptimistic
        If Not DB_RS.EOF Then
            Do While DB_RS.EOF = False
                DB_RS!server = 0
                DB_RS.MoveNext
            Loop
        End If
        DB_RS.Close
    End If
    
End Sub

Public Sub MySQL_Optimize()

    'Optimize the database tables
    DB_Conn.Execute "OPTIMIZE TABLE mail, mail_lastid, npcs, objects, quests, users"

End Sub
