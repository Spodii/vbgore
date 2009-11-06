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
if(!defined('IN_VBGORESP')){
	exit;
}

$config = Array();
$config['link'] = "localhost/vbGoreSP"; // No ending slash or opening http://
$config['htaccess'] = 1; // 1 for yes, 0 for no || Make sure that you have .htaccess before opening. If links don't work with it on, turn it off! //
$config['template_name'] = "vbgoresp"; // DO NOT CHANGE UNLESS YOU KNOW HOW TO MAKE TEMPLATES!!!!!!!!!!!!!!!! || Name of folder that contains template files

 /*
   ----------------------------------------------------------------------------------------------------
   REGISTRATION SETTINGS
   ----------------------------------------------------------------------------------------------------
	+ Read the comments next to the variable to know which is which.
		- Note on *Inventory, Mail, Known Skills*
		- You must seperate the numbers with ["\\r\\n"] YOU MUST!
   ----------------------------------------------------------------------------------------------------
   */
 $register = Array();
	$register[] = "1 1 5 0\\r\\n2 2 1 0\\r\\n3 3 1 0\\r\\n4 5 1 1\\r\\n5 6 1 1\\r\\n6 7 1 1"; // Inventory
	$register[] = ""; // Mail
	$register[] = "1\\r\\n2\\r\\n3\\r\\n4\\r\\n5\\r\\n6\\r\\n7\\r\\n8"; // Known Skills
	$register[] = ""; // Completed Quests
	$register[] = ""; // Current Quests
	$register[] = date(y.m.d); // Date Create && Date Last Login -> They both set the first date.
	$register[] = "20"; // Position X
	$register[] = "17"; // Position Y
	$register[] = "1"; // Position Starting Map
	$register[] = "1"; // Character Default Hair
	$register[] = "1"; // Character Default Head
	$register[] = "2"; // Character Default Body
	$register[] = "1"; // Character Default Weapon
	$register[] = "1"; // Character Default Wings
	$register[] = "3"; // Character Default Heading
	$register[] = "3"; // Character Default HeadHeading
	$register[] = "5"; // Slot of equipted weapon
	$register[] = "4"; // Slot of equipted armor
	$register[] = "6"; // Slot of equipted wings
	$register[] = "1"; // Stat: Strength
	$register[] = "1"; // Stat: Agility
	$register[] = "1"; // Stat: Magic
	$register[] = "1"; // Stat: Defence
	$register[] = "5"; // Stat: Speed
	$register[] = "100"; // Default Gold
	$register[] = "0"; // Default Exp
	$register[] = "1"; // Default Level
	$register[] = "10"; // Experience required for next level
	$register[] = "0"; // Stat Points
	$register[] = "1"; // Stat Hit Min
	$register[] = "1"; // Stat Hit Max
	$register[] = "50"; // Stat hp min
	$register[] = "50"; // Stat hp max
	$register[] = "50"; // Stat mp min
	$register[] = "50"; // Stat mp max
	$register[] = "50"; // Stat sp min
	$register[] = "50"; // Stat sp max
	
 /*
   ----------------------------------------------------------------------------------------------------
   CLASSES
   ----------------------------------------------------------------------------------------------------
	+ Seperate the number of the class from the name with [%%] like so:
		- 1[%%]GellyMONSTARRAWRGH!!!!
		- 2[%%]Theif
   ----------------------------------------------------------------------------------------------------
   */
 $classes = Array();
	$classes[] = "1[%%]Warrior";
	$classes[] = "2[%%]Mage";
	$classes[] = "4[%%]Rogue";
	
 /*
   ----------------------------------------------------------------------------------------------------
   LINK SEO AND NON SEO URLS
   ----------------------------------------------------------------------------------------------------
	+ Switches SEO Url Contents
		- If htaccess is on, Displays SEO urls
		- If htaccess is off, Displays normal urls
   ----------------------------------------------------------------------------------------------------
   */ 
	if($config['htaccess'] == 1){
		$link=Array();
			$link['login'] = 	"http://".$config['link']."/Login";
			$link['logout'] =	"http://".$config['link']."/Logout";
			$link['main'] = 	"http://".$config['link']."/Main";
			$link['register'] = "http://".$config['link']."/Register";
			$link['users'] =	"http://".$config['link']."/Users";
			$link['vusers'] = 	"http://".$config['link']."/View/";
			$link['online'] =	"http://".$config['link']."/ViewOnline";
		} else  {
		$link=Array();
			$link['login'] = 	"http://".$config['link']."/login.php";
			$link['logout'] =	"http://".$config['link']."/login.php?logout";
			$link['main'] = 	"http://".$config['link']."/index.php";
			$link['register'] = "http://".$config['link']."/register.php?do=register";
			$link['users'] =	"http://".$config['link']."/users.php";
			$link['vusers'] = 	"http://".$config['link']."/users.profile.php?user=";
			$link['online'] =	"http://".$config['link']."/users.php?list=online";
	}
	
 /*
   ----------------------------------------------------------------------------------------------------
   TEMPLATE FUNCTIONS
   ----------------------------------------------------------------------------------------------------
	+ Was getting tired of having to re "announce" this over and over and over
		- Reason for functions: 
			- Allows use of insertion of data instead of getting it when
			    it is not possible to do so.
			- Easier to access.
	+ Functions
		- echo Header
			- $flink is the links at the top of the page or wherever.
			- $config is.. well the configuration so that you can show the link
			   instead of making it global.
		- echo Footer
			- $online -> Displays users online.
   ----------------------------------------------------------------------------------------------------
   */ 
	function echoHeader($flink, $config){
	global $slink;
		return '<html>
					<head>
						<title>Vbgore SP - vbGore PHP Panel</title>
						<link rel="stylesheet" type="text/css" href="http://'.$config['link'].'/style.css" media="screen" />
					</head>
				<body>
					<div align="right">
						<h5><form action="index.php?do=qsearch" method="GET">Quick Search: <input name="q" type="text" size="20" alt="Enter Term and press Enter" title="Enter Term and press Enter"> <input type="submit" name="submit" value="" style="visibility:hidden;"></form>
							<br/>
							'.$flink.'
						</h5>
				</div>';
	}
	function echoFooter($online){
		return '<div class="Foot Apply">
					<p>
						<br /><br />
						Coding and Layout &copy; To DarkGrave Users online: '.$online.'
					</p>
				</div>
				</body>
				</html>';
	}
	
//  Making sure you read the whole thing. Change this to 1
	$config['edited'] = 1;
 
?>