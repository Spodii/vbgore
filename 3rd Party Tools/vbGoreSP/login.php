<? 
/*
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
 vbGORE Server Panel [vbGORE SP]
 @Author: DarkGrave
 @Date: June 23, 2007
 @Version: 2.0.0 (Re-Birth)
 @Copyright: Creative Commons Attribution-Noncommercial-Share Alike 3.0 License
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
 [You must leave this comment intact at all times]
   * This work is licensed under the Creative Commons Attribution-Noncommercial-Share Alike 3.0 License.
   * To view a copy of this license, visit http://creativecommons.org/licenses/by-nc-sa/3.0/ or send a letter to:
   * Creative Commons, 171 Second Street, Suite 300, San Francisco, California, 94105, USA.
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
*/
/* DO NOT REMOVE */
define('IN_VBGORESP', true);
/* DO NOT REMOVE */
include("common.php");
 
// Make sure they didn't want to logout!
if(isset($_GET['logout'])){
	if($userinfo == false)
	{
		header("Location: ".$link['login']);
	}
	else
	{
		setcookie("vbgore", "", time()-100000, "/", "", 0);
		header("Location: ".$link['login']);
	}
    die();
}
else
{
	// Check for a submission
    if (isset($_POST["submit"])) {
        
		// Parse the MySQL Query
		$query = $db->sql_query("SELECT * FROM {{table}} WHERE name='".$db->sql_escape($_POST["username"])."' AND password='".md5($_POST["password"])."' LIMIT 1", "users");

			// Check Database Rows
			if ($db->sql_rows($query) != 1) 
			{ 
				die($template->error("Login Error", "Invalid username or password. Please go back and try again.")); 
			}
			
			// Get an Array from the MySQL Query
			$row = $db->sql_array($query);
				if (isset($_POST["rememberme"])) 
				{ 
					$expiretime = time()+31536000; 
					$rememberme = 1; 
				} 
				else 
				{ 
					$expiretime = 0; 
					$rememberme = 0; 
				}
				$cookie = $row["name"] . " " . md5($row["password"]) . " " . $rememberme;
				setcookie("vbgore", $cookie, $expiretime, "/", "", 0);
			header("Location: ".$link['main']);
		die();
    }
	else 
	{
		$page  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
		$page .= $template->parse_template("template/".$config['template_name']."/login.form.php", Array('r_link' => $link['register']));
		$page .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
    echo $page;
	}
}

?>