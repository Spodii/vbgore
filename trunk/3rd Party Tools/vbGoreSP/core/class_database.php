<?
/*
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
 vbGORE Server Panel [vbGORE SP]
 @Author: DarkGrave
 @Date: 9:39 PM Tuesday, June 26, 2007
 @Version: 2.0.1 [Excentric]
 @Copyright: Creative Commons Attribution-Noncommercial-Share Alike 3.0 License
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
 [You must leave this comment intact at all times]
   * This work is licensed under the Creative Commons Attribution-Noncommercial-Share Alike 3.0 License.
   * To view a copy of this license, visit http://creativecommons.org/licenses/by-nc-sa/3.0/ or send a letter to:
   * Creative Commons, 171 Second Street, Suite 300, San Francisco, California, 94105, USA.
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
*/
if(!defined('IN_VBGORESP')){
	exit;
}

//Database Class
class db
{
	/*
	-------------------------------------------------------------
	Connect to MySQL Server
	-------------------------------------------------------------
	*/
	function mysql_connect($sqlserver, $sqluser, $sqlpassword, $database, $port = false) {
	// Defining Variables
		$this->user = $sqluser;
		$this->server = $sqlserver . (($port) ? ':' . $port : '');
		$this->dbname = $database;
		
		// Start the actual connection
	    $this->db_link = @mysql_connect($this->server, $this->user, $sqlpassword) or die($this->sql_error(mysql_errno(),mysql_error()));
		
		// Check the database link for errors, if there is one. Spit out the custom mysql_error!
	    if (!$this->db_link) {
			die($this->sql_error(mysql_errno(),mysql_error()));
		}
		
		// Try to connect to the database, if not Spit out the custom mysql_error!
		if(!@mysql_select_db($this->dbname, $this->db_link)){ 
			die($this->sql_error(mysql_errno(),mysql_error())); 
		}
		
	return $this->db_link;
	}
	
	/*
	-------------------------------------------------------------
	Base Query Parsing
	-------------------------------------------------------------
	@param	string	$query		Contains the SQL query which shall be executed.
	@param	string	$table		Contains the Table that contains the data we are after.
	@return	mixed				When parsed, it checks whether the result was successful, If not sends error. If it was it sends the correct information requested.
	-------------------------------------------------------------
	*/
	function sql_query($query = '', $table = '') {
	
	    $this->query_result = mysql_query(str_replace("{{table}}", $table, $query), $this->db_link);
		
		if(!$this->query_result){
			die($this->sql_error(mysql_errno(),mysql_error()));
		}
		
	return $this->query_result;
	}
	
	/*
	-------------------------------------------------------------
	Base Array Sending
	-------------------------------------------------------------
	@param	string	$query		Contains the SQL query which shall be executed.
	@return	mixed				When parsed, it checks whether the result was successful, If not sends error. If it was it sends the correct information requested.
	-------------------------------------------------------------
	*/
	function sql_array($query = '') {
	
	$this->array_result = mysql_fetch_array($query);
		
		if(!$this->array_result){
			die($this->sql_error(mysql_errno(),mysql_error()));
		}
		
	return $this->array_result;
	}
	
	/*
	-------------------------------------------------------------
	Base Row Check
	-------------------------------------------------------------
	@param	string	$query		Contains the SQL query which shall be executed.
	@return	mixed				When parsed, it checks whether the result was successful, If not sends error. If it was it sends the correct information requested.
	-------------------------------------------------------------
	*/
	function sql_rows($query = '') {
	
	$this->row_result = mysql_num_rows($query);
		
	return $this->row_result;
	}
	
	/*
	-------------------------------------------------------------
	Base Query Parsing
	-------------------------------------------------------------
	@param	string	$msg			Contains the string which shall be parsed
	-------------------------------------------------------------
	*/
	function sql_escape($msg)
	{
		if (!$this->db_connect_id)
		{
			return @mysql_real_escape_string($msg);
		}

		return @mysql_real_escape_string($msg);
	}
	
	/*
	-------------------------------------------------------------
	MySQL Error Template
	-------------------------------------------------------------
	@param	string	$errno		Contains the MySQL Error Number
	@param	string	$reason		Contains the MySQL Error
	-------------------------------------------------------------
	*/
	function sql_error($errno,$reason){
	return '<html>
				<head>
					<title>vbGORESp - Mysql Error #'.$errno.'</title>
					<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
					<link rel="stylesheet" type="text/css" href="template/vbgoresp/css/mysql_error.css" />
				</head>
				<body>
					<div id="container">
						<h1 style="border-bottom: 1px solid #ddd; margin-bottom: 4px;">vbGORE<span style="color:limegreen">SP</span> &rarr; Mysql Error #'.$errno.'</h1><br />
						<div class="error">'.$reason.'</div><br />
						<div class="note"><font style="color:#990000;">*Note:</font> If you cannot figure out how to get rid of this message or do not understand the error, please message "DarkGrave" on http://vbgore.com/forums/. Thank you very much.</div>
				</body>
			</html>';
	}

}

?>