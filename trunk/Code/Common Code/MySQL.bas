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

'API to open the browser (used for MySQL connection errors)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub MySQL_Init()
Dim ErrorString As String
Dim i As Long

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
    
    'Run test queries to make sure the tables are there
    DB_RS.Open "SELECT * FROM banned_ips WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM mail WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM mail_lastid WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM npcs WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM objects WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM quests WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close
    DB_RS.Open "SELECT * FROM users WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close

    On Error GoTo 0
    
    Exit Sub
    
ErrOut:
    
    'Refresh the errors
    DB_Conn.Errors.Refresh
    
    'Get the error string if there is one
    If DB_Conn.Errors.Count > 0 Then ErrorString = DB_Conn.Errors.Item(0).Description

    'Check for known errors
    If InStr(1, ErrorString, "Access denied for user ") Then
        
        'Invalid username or password
        ShellExecute frmMain.hWnd, vbNullString, "http://www.vbgore.com/MySQL_Setup#Access_denied", vbNullString, "c:\", 10
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "An incorrect username and/or password was entered into the configuration file." & vbNewLine & _
            "This information can be changed in the \ServerData\Server.ini file on the 'User=' and 'Password=' lines." & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            "Username: " & DB_User & "   Password: " & DB_Pass & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & _
            "MySQL returned the following error message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"

    ElseIf InStr(1, ErrorString, "Can't connect to MySQL server on ") Then
    
        'Unable to connect to MySQL
        ShellExecute frmMain.hWnd, vbNullString, "http://www.vbgore.com/MySQL_Setup#Can.27t_connect_to_MySQL_server", vbNullString, "c:\", 10
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "Either an invalid MySQL server IP and/or port was entered, or the server isn't running!" & vbNewLine & _
            "Please confirm you installed MySQL 5.0 and ran the Instance Configuration." & vbNewLine & _
            "To manually start the instance, do the following:" & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            "Right-click 'My Computer' -> 'Manage' -> 'Services and Applications' -> 'Services'" & vbNewLine & _
            "Find your MySQL service in this list (name usually starts with 'MySQL'), right-click it and click 'Start'" & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & vbNewLine & _
            "MySQL returned the following error message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
            
    ElseIf InStr(1, ErrorString, "Unknown database ") Then
        
        'Invalid database name / database does not exist
        ShellExecute frmMain.hWnd, vbNullString, "http://www.vbgore.com/MySQL_Setup#Unknown_database", vbNullString, "c:\", 10
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "An invalid or unknown database name, '" & DB_Name & "', was entered." & vbNewLine & _
            "This information can be changed in the \ServerData\Server.ini file on the 'Database=' line." & vbNewLine & vbNewLine & _
            "MySQL returned the following error message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
            
    ElseIf InStr(1, ErrorString, "Data source name not found and no default driver specified") Then
        
        'Invalid database name / database does not exist
        ShellExecute frmMain.hWnd, vbNullString, "http://www.vbgore.com/MySQL_Setup#Driver_not_found", vbNullString, "c:\", 10
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "No valid driver could be found on this computer to connect to MySQL." & vbNewLine & _
            "Please make sure you install ODBC v3.51 (must be v3.51) on this computer!" & vbNewLine & _
            "ODBC can be downloaded from:" & vbNewLine & _
            "http://dev.mysql.com/downloads/connector/odbc/3.51.html" & vbNewLine & vbNewLine & _
            "MySQL returned the following error message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
            
    ElseIf InStr(1, ErrorString, "Table '") & InStr(1, ErrorString, "' doesn't exist") Then
        
        'At least one of the tables are missing
        ShellExecute frmMain.hWnd, vbNullString, "http://www.vbgore.com/MySQL_Setup#Table_doesn.27t_exist", vbNullString, "c:\", 10
        MsgBox "Error connecting to the MySQL database!" & vbNewLine & _
            "One or more of the tables required were not found." & vbNewLine & _
            "Please make sure you import the 'vbgore.sql' file found in the folder '/_Database Dump/' into the database." & vbNewLine & vbNewLine & _
            "MySQL returned the following error message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------"
    
    Else
    
        'Unknown error
        MsgBox "Unknown error connecting to the MySQL database!" & vbNewLine & _
            "Please confirm that you have completed the following tasks:" & vbNewLine & vbNewLine & _
            " - You have followed ALL of the steps on the MySQL Setup page on the vbGORE site" & vbNewLine & _
            " - MySQL is running and you can connect to it through a GUI such as SQLyog" & vbNewLine & _
            " - You have imported the vbgore.sql file into the database and can see the information through the MySQL GUI" & vbNewLine & _
            " - You have version 5.0 of MySQL and 3.51 of ODBC being used" & vbNewLine & _
            " - You changed the \ServerData\Server.ini file to use your MySQL information" & vbNewLine & vbNewLine & _
            "If you are positive you have done all of the above, ask for help on the vbGORE forums." & vbNewLine & vbNewLine & _
            "MySQL returned the following error message: " & vbNewLine & _
            "---------------------------------------------------------------------------------------------------" & vbNewLine & _
            ErrorString & vbNewLine & _
            "---------------------------------------------------------------------------------------------------", vbOKOnly
    
    End If
    
    Server_Unload

End Sub

Public Sub MySQL_RemoveOnline()

    'Remove the online flag from all the users (recommended for server loading in case of a crash)
    If ServerID > 0 Then    'This prevents the map editor making this call
        DB_RS.Open "SELECT * FROM users WHERE `server`='" & ServerID & "'", DB_Conn, adOpenStatic, adLockOptimistic
        If Not DB_RS.EOF Then
            Do While Not DB_RS.EOF
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
