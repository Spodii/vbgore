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
---------------------------------------------------------------------------------------------------
QUICK SEARCH
----------------------------------------------------------------------------------------------------
	+ Give users the ability to search for someone
		- Returns list of users with same char, name, etc as entered
		- Ability to click the name and view their information
		- Happy Days.
----------------------------------------------------------------------------------------------------
*/
	if ($userrow == false) 
	{
		if(isset($_GET['q']))
		{
		
			$trimmed = mysql_real_escape_string(trim($_GET['q'])); //trim whitespace from the stored variable
			$before = $trimmed;
			$cleanse  = array('/','\\','\'','"',',','.','<','>','?',';',':','[',']','{','}','|','=','+','-','_',')','(','*','&','^','%','$','#','@','!','~','`');//this will remove punctuation
			$clean_search = str_replace($cleanse, '', $before);
			
			$query = $db->sql_query("SELECT * FROM {{table}} WHERE `name` LIKE \"%$clean_search%\" ORDER BY gm DESC,name ASC", "users");
			$numrows=mysql_num_rows($query);
			if ($numrows == 0)
			{
				$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
				$show .= $template->parse_template("template/".$config['template_name']."/search.error.php", Array('error' => "No Search Results =("));
				$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
				echo $show;
			} 
			else 
			{
				$page  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
				$page .= $template->parse_template("template/".$config['template_name']."/search.results.header.php", Array());
				while ($row=mysql_fetch_array($query)) 
				{
					$name = $row['name'];
					$row["descr"] = check_description($row["descr"]);
					$row["gm"] = check_gm($row["gm"]);
					
						// Grab HTML and Parse It
						$page .= $template->parse_template("template/".$config['template_name']."/search.results.php", Array('name'=>$row['name'], 'description' => $row['descr'], 'user_or_gm' => $row['gm'] , 'link_vusers' => $link['vusers']));
				}
				$page .= $template->parse_template("template/".$config['template_name']."/search.results.footer.php", Array());
				$page .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
			echo $page;
			}
		} 
		else 
		{
				$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
				$show .= $template->parse_template("template/".$config['template_name']."/search.error.php", Array('error' => "Please enter a search query!"));
		$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
		echo $show;
		}
	}
	else 
	{
		$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
		$show .= $template->parse_template("template/".$config['template_name']."/search.error.php", Array('error' => "No Search Entered!"));
		$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
    echo $show;
	}
?>