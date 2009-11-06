<?php
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

class cookie extends db
{
	/*
		Cookies Are Awesome!
	*/
	function parse_cookie($db, $template) 
	{
	$row = false;
	    
	    if (isset($_COOKIE["vbgore"])) {
		$cookieinfo = explode(" ",$_COOKIE["vbgore"]);
	    $query = $db->sql_query("SELECT * FROM {{table}} WHERE name='$cookieinfo[0]'", "users");
		
	        if ($db->sql_rows($query) != 1) 
			{ 
				die($template->error("Cookie Error #1","Invalid cookie data. Please clear cookies and log in again.")); 
			}
			else 
			{
			$row = $db->sql_array($query);
			}
			if ($row["name"] != $cookieinfo[0]) 
			{ 
				die($template->error("Cookie Error #2","Invalid cookie data. Please clear cookies and log in again."));  
			}
			if (md5($row["password"]) !== $cookieinfo[1]) 
			{ 
				die($template->error("Cookie Error #3","Invalid cookie data. Please clear cookies and log in again."));  
			}
	        // If we've gotten this far, cookie should be valid, so write a new one.
	        $newcookie = implode(" ",$cookieinfo);
			
			if ($cookieinfo[2] == 1) 
			{ 
				$exptime = time()+31536000; 
			} 
			else 
			{ 
				$exptime = 0; 
			}
				
			setcookie("vbgore", $newcookie, $exptime, "/", "", 0);
	        
	    }
	        
	return $row;
	}
	
	/*
	-------------------------------------------------------------
	Cookie_Error Template
	-------------------------------------------------------------
	@param	int		$errno		Contains the Error Number
	@param	string	$reason		Contains the Base Error
	-------------------------------------------------------------
	*/
	function cookie_error($errno,$reason)
	{
	return '<html>
				<head>
					<title>vbGORESp - Cookie Error #'.$errno.'</title>
					<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
					<link rel="stylesheet" type="text/css" href="template/vbgoresp/css/template_error.css" />
				</head>
				<body>
					<div id="container">
						<h1 style="border-bottom: 1px solid #ddd; margin-bottom: 4px;">vbGORE<span style="color:limegreen">SP</span> &rarr; Cookie Error #'.$errno.'</h1><br />
						<div class="error">'.$reason.'</div><br />
						<div class="note"><font style="color:#990000;">*Note:</font> If you cannot figure out how to get rid of this message or do not understand the error, please message "DarkGrave" on http://vbgore.com/forums/. Thank you very much.</div>
				</body>
			</html>';
	}
}

?>