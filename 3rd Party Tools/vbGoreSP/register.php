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
/* DO NOT REMOVE */
define('IN_VBGORESP', true);
/* DO NOT REMOVE */
include("common.php");

 /*
   ----------------------------------------------------------------------------------------------------
   REGISTRATION PAGE.
   ----------------------------------------------------------------------------------------------------
	- Lots of things that are explained further down.
   ----------------------------------------------------------------------------------------------------
   */
	if ($userrow == false) 
	{
 
		if (isset($_POST["register"])) 
		{
			$name = $_POST['name'];
			$password = $_POST['password'];
			$passwordag = $_POST['passwordag'];
			$desc = $_POST['desc'];
			$class = html_deep($_POST['class']);
			$desc2 = html_deep($desc);
			$errors = 0;
			$errorlist = "";
				if ($name == "") { 
					$errors++; 
					$errorlist .= "<span style=color:red>* Name is required!</span><br />"; 
				}
				if ($desc2 == "") 
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Description is required!<br />"; 
				}
				if ($password == "") 
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Password Is Required!</span><br />"; 
				}

			 /*
			   ----------------------------------------------------------------------------------------------------
			   REGISTRATION FORM CHECK
			   ----------------------------------------------------------------------------------------------------
			  	+ Check for errors (In order)
					+ Check for username length
						- If below 5 chars give error.
			  		+ Strip Junk characters
			  			- See if it matches original if not, error.
					+ Check for cussing, other.
						- If username/description includes them, error.
			  		+ Check to see if username is taken.
						- If it is, give error.
					+ Check password against re-entered password
						+ If error, they can't remember passwords easily..
							- Give huge error.
			  ----------------------------------------------------------------------------------------------------
			  */
			$subname = stripslashes($name);
			$subpass = stripslashes($password);
				if(strlen($subname) < 4 && empty($errorlist)) { $errors++; $errorlist .= "<span style=color:red>* Username below 5 characters!</span><br />"; } else if(strlen($subname) > 30) { $errors++; $errorlist .= "<span style=color:red>* Username above 30 characters!</span><br />"; }
				if(strlen($subpass) < 4 && empty($errorlist)) { $errors++; $errorlist .= "<span style=color:red>* Password below 5 characters!</span><br />"; } else if(strlen($subpass) > 30) { $errors++; $errorlist .= "<span style=color:red>* Password above 30 characters!</span><br />"; }
				
			$junk = array(',' , '/' , '\\' , '[' ,  ']' , '&', '^', '%', '$', '#', '!', '+', '(', ')', '|', '<', '>', ':', '"', '=', '@', '~','\'');
			$subname = trim($subname);
			$subpass = trim($subpass);
			$subname1 = str_replace($junk, '', $subname);
			$subpass1 = str_replace($junk, '', $subpass);

				if($subname1 != $subname && empty($errorlist))
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Name contains Invalid characters</span><br />"; 
				}
				if($subpass1 != $subpass && empty($errorlist))
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Password contains Invalid characters</span><br />"; 
				}

			$check = eregi("fuck|pussy|penis|rape|slut|whore|bitch|vagina|legendary|greatest|pwnt|n00b|queer|anal|dick|suck|shit|faggot|prick|Admin", $subname);
			$check2 = eregi("fuck|pussy|penis|rape|slut|whore|bitch|vagina|legendary|greatest|pwnt|n00b|queer|anal|dick|suck|shit|faggot|prick|Admin", $desc2);

				if($check)
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Your Character Name Contains Foul Language or Admin/GM.</span><br />"; 
				}
				if($check2)
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Your Description Contains Foul Language or Names.</span><br />"; 
				}

		    $usernamequery = $db->sql_query("SELECT name FROM {{table}} WHERE name='$subname' LIMIT 1","users");

				if (mysql_num_rows($usernamequery) > 0) 
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Username already taken</span><br />"; 
				}
		        if ($passwordag != $password && empty($errorlist))
				{ 
					$errors++; 
					$errorlist .= "<span style=color:red>* Your Passwords do not match!</span><br />"; 
				}

			$password3 = md5($password);

		 /*
		   ----------------------------------------------------------------------------------------------------
		   NO ERRORS?!
		   ----------------------------------------------------------------------------------------------------
			+ Check to see if errors
				+ If errors
					-- Go to show error list
				+ Else
					- Insert Userdata into mysql table!
		   ----------------------------------------------------------------------------------------------------
		   */
    		if($errors == 0){
        	// Insert Data
			$query = $db->sql_query("INSERT INTO {{table}} SET name='$name',ip='',password='$password3',descr='$desc',class='$class',inventory='".$register[0]."',mail='".$register[1]."',knownskills='".$register[2]."',completedquests='".$register[3]."',currentquest='".$register[4]."',date_create='".$register[5]."',date_lastlogin='".$register[5]."',pos_x='".$register[6]."', pos_y='".$register[7]."',pos_map='".$register[8]."',char_hair='".$register[9]."',char_head='".$register[10]."',char_body='".$register[11]."',char_weapon='".$register[12]."',char_wings='".$register[13]."', char_heading='".$register[14]."',char_headheading='".$register[15]."',eq_weapon='".$register[16]."',eq_armor='".$register[17]."',eq_wings='".$register[18]."',stat_str='".$register[19]."', stat_agi='".$register[20]."',stat_mag='".$register[21]."',stat_def='".$register[22]."',stat_speed='".$register[23]."',stat_gold='".$register[24]."',stat_exp='".$register[25]."',stat_elv='".$register[26]."',stat_elu='".$register[27]."',stat_points='".$register[28]."',stat_hit_min='".$register[29]."',stat_hit_max='".$register[30]."',stat_hp_min='".$register[31]."', stat_hp_max='".$register[32]."', stat_mp_min='".$register[33]."',stat_mp_max='".$register[34]."',stat_sp_min='".$register[35]."',stat_sp_max='".$register[36]."',server='0'", "users");
			
			$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
			$show .= $template->parse_template("template/".$config['template_name']."/register.page.success.php", $parse = Array('link_login' => $link['login']));
			$show .=  $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
			
          echo $show;

    } 
	else 
	{

	 /*
	   ----------------------------------------------------------------------------------------------------
	   ERROR LIST
	   ----------------------------------------------------------------------------------------------------
		+ Someone Tried to do some bad things here..
			- Prevent Hacking
			- Prevent Invalid Names
			- Prevent Everything else that isn't right from happening.
	   ----------------------------------------------------------------------------------------------------
	   */
	$errorlist .= "<br /><a href=index.php>Go back!</a>";
	
		$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
		$show .= $template->parse_template("template/".$config['template_name']."/register.page.error.php", Array('error_list' => $errorlist));
		$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
    echo $show;
   }

 /*
   ----------------------------------------------------------------------------------------------------
   REGISTRATION FORM
   ----------------------------------------------------------------------------------------------------
	+ Form of registration
		- Sets information into database and allows users to login to the site
		   and vbgore.
   ----------------------------------------------------------------------------------------------------
   */
  }
  else 
  {
	$page  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
	$page .= $template->parse_template("template/".$config['template_name']."/register.page.header.php", Array());
		// Show Classes
		foreach($classes as $class)
		{
			$class_disperse = explode("[%%]", $class);
			$page .= '<option value="'.$class_disperse[0].'" />'.$class_disperse[1].'</option>';
		}
	$page .= $template->parse_template("template/".$config['template_name']."/register.page.footer.php", Array());
	$page .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));

	echo $page;
  }

 /*
   ----------------------------------------------------------------------------------------------------
   IF LOGGED IN...
   ----------------------------------------------------------------------------------------------------
	+ Prevents the following
		- Hacking Attempts
		- Stuck Users
		- Lost Users
   ----------------------------------------------------------------------------------------------------
   */
 } else {
 
		$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
		$show .= $template->parse_template("template/".$config['template_name']."/register.page.error.php", Array('error_list' => "<span style=color:red>You are already registered!</span><br />"));
		$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));

    echo $show;
}
?>