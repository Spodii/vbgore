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

// Include files
require('core/class_database.php');
require('core/class_template.php');
require('core/class_cookies.php');
require('core/functions.php');
require('config.php');
require('connect.php');

if (get_magic_quotes_gpc()) {

	$_POST = array_map('stripslashes_deep', $_POST);
	$_GET = array_map('stripslashes_deep', $_GET);
	$_COOKIE = array_map('stripslashes_deep', $_COOKIE);

}
	$_POST = array_map('addslashes_deep', $_POST);
	$_GET = array_map('addslashes_deep', $_GET);
	$_GET = array_map('html_deep', $_GET);
	$_COOKIE = array_map('addslashes_deep', $_COOKIE);
	$_COOKIE = array_map('html_deep', $_COOKIE);

// Instantiate some basic classes
$db			= new db();
$cookie		= new cookie();
$template	= new template();

if ($dbsettings['user'] == "MySql_root" or $dbsettings['pass'] == "MySql_password" or $dbsettings['name'] == "MySql_DBName") 
{
	die($template->error('MySQL Connection Settings..', 'You have not edited <strong style="color:red"><i>connect.php</i></strong> fully. Default dummy settings exist.'));
}

// Connect to Database
$db->mysql_connect($dbsettings['host'], $dbsettings['user'], $dbsettings['pass'], $dbsettings['name'], $dbsettings['port']);

// Make sure they edited the config
if($config['edited'] == 0)
{ 
		die($template->error('Configuration Error', 'Make sure you completely edit "Config.php".'));
}

// We do not need this any longer, unset for safety purposes
unset($dbsettings['pass']);

// Check Cookies
$userinfo = $cookie->parse_cookie($db, $template);

// Check Users Online
$online_query = $db->sql_query("SELECT server FROM {{table}} WHERE server='1' LIMIT 10000", "users");
$users_online = $db->sql_rows($online_query);

// Create Navigation Menu
$navigation = $template->create_nav($userinfo,$link,$users_online);

?>