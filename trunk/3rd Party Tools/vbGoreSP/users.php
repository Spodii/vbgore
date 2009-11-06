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
USER LIST, ONLINE LIST
----------------------------------------------------------------------------------------------------
	+ Handles both the User List and Online List
		- User List Shows ALL users.
		- Online List Shows ONLY online users.
----------------------------------------------------------------------------------------------------
*/
	$olist = $_GET["list"];
    $query = $db->sql_query("SELECT * FROM {{table}} ORDER BY gm DESC,name ASC", "users");

	if ($db->sql_rows($query) == 0) 
	{ 
		$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
		$show .= $template->parse_template("template/".$config['template_name']."/user.list.error.php", Array('error' => "No Users in database!"));
		$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
	}
	$page  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
	$page .= $template->parse_template("template/".$config['template_name']."/user.list.results.header.php", Array());
	$count = 0;
	    while ($row = mysql_fetch_array($query))
		{
			$row["descr"] = check_description($row["descr"]);
			$row["gm"] = check_gm($row["gm"]);
			if($olist == "online")
			{
				if($row['server'] == 1)
				{ 		
					$page .= $template->parse_template("template/".$config['template_name']."/user.list.results.php", Array('name'=>$row['name'], 'description' => $row['descr'], 'user_or_gm' => $row['gm'] , 'link_vusers' => $link['vusers']));
					++$count;
				}
			}
			else 
			{ 
				$page .= $template->parse_template("template/".$config['template_name']."/user.list.results.php", Array('name'=>$row['name'], 'description' => $row['descr'], 'user_or_gm' => $row['gm'] , 'link_vusers' => $link['vusers']));
			}
		}
	if($olist == "online")
	{
		if($count == 0)
		{
			$page .= $template->parse_template("template/".$config['template_name']."/user.list.online.error.php", Array('error' => "No Users are online =(!"));
		}
	}
    $page .= $template->parse_template("template/".$config['template_name']."/user.list.results.footer.php", Array());
	$page .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
    echo $page;
?>