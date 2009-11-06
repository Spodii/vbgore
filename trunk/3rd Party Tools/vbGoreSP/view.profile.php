<? 
/*
 ----------------------------------------------------------------------------------------------------------------------------------------------------------------
 vbGore Server Panel [vbGore SP]
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
   VIEW USER PROFILE
   ----------------------------------------------------------------------------------------------------
	+ Check For ID
		+ Id Check Verified
			+ Check ID Against Database
				+ Check Verified
					- Show User Profile
				+ Check Failed
					- Send User Error
		+ Id Check Failed
			- Send Error
   ----------------------------------------------------------------------------------------------------
   */
if(isset($_GET['user']))
{
	$id = $_GET['user'];
	$userquery = $db->sql_query("SELECT * FROM {{table}} WHERE name='".$db->sql_escape($id)."' LIMIT 1", "users");
		if ($db->sql_rows($userquery) == 1) 
		{ 
			$userrow = $db->sql_array($userquery); 
		} 
		else 
		{ 
			$show .= $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
			$show .= $template->parse_template("template/".$config['template_name']."/view.profile.error.php", $parse = Array('base_url' => $config['link']));
			$show .= $template->parse_template("template/".$config['template_name']."/footer.php", $parse = Array('online' => $users_online));
		echo $show;
	die();
    }
	
	$statbar = $template->create_statbar($userrow["stat_hp_max"], $userrow["stat_hp_min"], $userrow["stat_mp_max"], $userrow["stat_mp_min"], $userrow["stat_sp_max"], $userrow["stat_sp_min"], $config);
    $userrow["stat_exp"] = number_format($userrow["stat_exp"]);
    $userrow["stat_gold"] = number_format($userrow["stat_gold"]);
	$userrow["server"] = check_status($userrow["server"]);
	$userrow["descr"] = str_cut(check_description($userrow["descr"]), 25);

 	$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
	$show .= $template->parse_template("template/".$config['template_name']."/view.profile.php", $parse = Array('statbar' => $statbar,'name' => $userrow['name'],'descr' => $userrow['descr'],'stat_elv' => $userrow['stat_elv'],'gm' => $userrow['gm'],'stat_gold' => $userrow['stat_gold'],'stat_exp' => $userrow['stat_exp'],'stat_str' => $userrow['stat_str'],'stat_agi' => $userrow['stat_agi'],'stat_def' => $userrow['stat_def'],'stat_mag' => $userrow['stat_mag'],'link_online' => $link['online'],'server_status' => $userrow['server']));
	$show .= $template->parse_template("template/".$config['template_name']."/footer.php", $parse = Array('online' => $users_online));

	echo $show;
	
}
?>