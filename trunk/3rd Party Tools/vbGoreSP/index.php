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
   USER PROFILE
   ----------------------------------------------------------------------------------------------------
	+ Check logged in
	  - Logged in -> Profile
	  - Other -> Login Form
   ----------------------------------------------------------------------------------------------------
   */
	if($userinfo == false)
	{
		header("Location: ".$link['login']);
	} 
	else 
	{
		if(isset($_POST['cd'])){
			$desc = html_deep($_POST['desc']);
			$db->sql_query("UPDATE `{{table}}` SET `descr`='".$db->sql_escape($desc)."' WHERE `name`='".$db->sql_escape($userinfo['name'])."'","users");
			$userinfo = $cookie->parse_cookie($db, $template);
		}
		
		$statbar = $template->create_statbar($userinfo["stat_hp_max"], $userinfo["stat_hp_min"], $userinfo["stat_mp_max"], $userinfo["stat_mp_min"], $userinfo["stat_sp_max"], $userinfo["stat_sp_min"], $config);
		$userinfo["descr"] = str_cut(check_description($userinfo["descr"]), 25);

		// Create Page HTML
		$show  = $template->parse_template("template/".$config['template_name']."/header.php", Array('navigation' => $navigation, 'base_url' => $config['link'], 'template' => $config['template_name']));
		$show .= $template->parse_template("template/".$config['template_name']."/profile.php", Array('statbar' => $statbar,'name' => $userinfo['name'],'descr' => $userinfo['descr'],'stat_elv' => $userinfo['stat_elv'],'gm' => $userinfo['gm'],'stat_gold' => $userinfo['stat_gold'],'stat_exp' => $userinfo['stat_exp'],'stat_str' => $userinfo['stat_str'],'stat_agi' => $userinfo['stat_agi'],'stat_def' => $userinfo['stat_def'],'stat_mag' => $userinfo['stat_mag']));
		$show .= $template->parse_template("template/".$config['template_name']."/footer.php", Array('online' => $users_online));
		
	// Show Page
	echo $show;
  }
	
?>