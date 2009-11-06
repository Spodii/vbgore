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

//Template Class
class template
{
	/*
	-------------------------------------------------------------
	Initiate Parse Template Feature
	-------------------------------------------------------------
	*/
	function parse_template($template, $array)
	{
	$this->template_name = $template;
	$this->template_array = $array;
	
		if (file_exists($template))
		{
		include($template);
		$this->template_contents = $base;
		
		    foreach($this->template_array as $a => $b) {
		        $this->template_contents = str_replace('{'.$a.'}', $b, $this->template_contents);
		    }
			
		}
		else
		{
			die($this->error("No Such Template","Template path could not be found: $this->template_name"));
		}
		
	return $this->template_contents;
	}
	
	function create_statbar($max_hp, $min_hp, $max_mp, $min_mp, $max_sp, $min_sp, $config){
		if($min_hp > $max_hp){ $max_hp = $min_hp; }
		if($min_mp > $max_mp){ $max_mp = $min_mp; }
		if($min_sp > $max_sp){ $max_sp = $min_sp; }
		$stathp = ceil($min_hp / $max_hp * 100); 
		$statsp = ceil($min_sp / $max_sp * 100);
		if ($min_mp != 0) { $statmp = ceil($min_mp / $max_mp * 100); } 
		else { $statmp = 0; }
		
		$stattable = "<label><table width=\"100%\" cellspacing=\"0\" cellpadding=\"0\"><tr><tr><td align=\"left\"><table cellspacing=\"2\" cellpadding=\"0\" title=\"Hit Points: ".$min_hp." / ".$max_hp."\"><tr><td style=\"width: 25px; height:10px; font: 10px Verdana; vertical-align: middle;\">HP:</td><td style=\"padding:0px; width:100px; height:10px; border:solid 1px black; background-image:url(http://".$config['link']."/images/bars_grey.gif);\">\n";
	    	if ($stathp >= 66){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$stathp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_green.gif);\"><img src=\"http://".$config['link']."/images/bars_green.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_greenend.gif\" alt=\"\" /></div>"; }
			if ($stathp < 66 && $stathp >= 33){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$stathp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_yellow.gif);\"><img src=\"http://".$config['link']."/images/bars_yellow.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_yellowend.gif\" alt=\"\" /></div>"; }
			if ($stathp < 33 && $stathp >= 4){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$stathp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_red.gif);\"><img src=\"http://".$config['link']."/images/bars_red.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_redend.gif\" alt=\"\" /></div>"; }
			if ($stathp < 3){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$stathp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_red.gif);\"><img src=\"http://".$config['link']."/images/bars_red.gif\" alt=\"\" /></div>"; }
		$stattable .="</td><td style=\"width: 60px; height:10px; font: 10px Verdana; vertical-align: middle;\">&nbsp;".$min_hp." / ".$max_hp."</td></tr></table></td></tr>";
		$stattable .= "<tr><td align=\"left\"><table cellspacing=\"2\" cellpadding=\"0\" title=\"Mana Points: ".$min_mp." / ".$max_mp."\"><tr><td style=\"width: 25px; height:10px; font: 10px Verdana; vertical-align: middle;\">MP:</td><td style=\"padding:0px; width:100px; height:10px; border:solid 1px black; background-image:url(http://".$config['link']."/images/bars_grey.gif);\">\n";
	        if ($statmp >= 66){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statmp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_green.gif);\"><img src=\"http://".$config['link']."/images/bars_green.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_greenend.gif\" alt=\"\" /></div>"; }
	        if ($statmp < 66 && $statmp >= 33){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statmp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_yellow.gif);\"><img src=\"http://".$config['link']."/images/bars_yellow.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_yellowend.gif\" alt=\"\" /></div>";}
			if ($statmp < 33 && $statmp >= 4){ $stattable .= "<div style=\" text-align:right; padding:0px; width:".$statmp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_red.gif);\"><img src=\"http://".$config['link']."/images/bars_red.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_redend.gif\" alt=\"\" /></div>"; }
			if ($statmp < 3){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statmp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_red.gif);\"></div>"; }
		$stattable .="</td><td style=\"width: 60px; height:10px; font: 10px Verdana; vertical-align: middle;\">&nbsp;".$min_mp." / ".$max_mp."</td></tr></table></td></tr>";
		$stattable .= "<tr><td align=\"left\"><table cellspacing=\"2\" cellpadding=\"0\" title=\"Stamina: ".$min_sp." / ".$max_sp."\"><tr><td style=\"width: 25px; height:10px; font: 10px Verdana; vertical-align: middle;\">SP:</td><td style=\"padding:0px; width:100px; height:10px; border:solid 1px black; background-image:url(http://".$config['link']."/images/bars_grey.gif);\">\n";
	        if ($statsp >= 66){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statsp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_green.gif);\"><img src=\"http://".$config['link']."/images/bars_green.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_greenend.gif\" alt=\"\" /></div>"; }
	        if ($statsp < 66 && $statsp >= 33){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statsp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_yellow.gif);\"><img src=\"http://".$config['link']."/images/bars_yellow.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_yellowend.gif\" alt=\"\" /></div>"; }
	        if ($statsp < 33 && $statsp >= 4){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statsp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_red.gif);\"><img src=\"http://".$config['link']."/images/bars_red.gif\" alt=\"\" /><img src=\"http://".$config['link']."/images/bars_redend.gif\" alt=\"\" /></div>"; }
			if ($statsp < 3){ $stattable .= "<div style=\"text-align:right; padding:0px; width:".$statsp."px; border-top:solid 0px black; background-image:url(http://".$config['link']."/images/bars_red.gif);\"><img src=\"http://".$config['link']."/images/bars_red.gif\" alt=\"\" /></div>"; }
		$stattable .="</td><td style=\"width: 60px; height:10px; font: 10px Verdana; vertical-align: middle;\">&nbsp;".$min_sp." / ".$max_sp."</td></tr></table></td></tr></table></label>";	
		$finalizetable = $stattable;
		
	return $finalizetable;
	}
	
	/*
	-------------------------------------------------------------
	Create Navigation
	-------------------------------------------------------------
	@param	string	$userrow		Contains user information, if there is any.
	@param	array		$links		Contains the Links that will be shown in the navigation
	@param	int		$online		Number of users online.
	@return	mixed				When parsed, it checks whether the user is logged in, If not sends guest navigation. If it was it sends the users navigation.
	-------------------------------------------------------------
	*/
	function create_nav($userinfo,$links,$online)
	{
	$this->userinfo = $userinfo;
	$this->link = $links;
	$this->online = $online;
	
		if ($this->userinfo == false) 
		{ 
		    $this->username = "Guest";
			$this->navigation = 'Welcome '.$this->username.'&nbsp;&bull;&nbsp;
				<a href="'.$this->link['register'].'">Register Page</a>&nbsp;&bull;&nbsp;
				<a href="'.$this->link['users'].'">Users Page</a>&nbsp;&bull;&nbsp;
				<a href="'.$this->link['login'].'">Login</a>&nbsp;&bull;&nbsp;
				<a href="'.$this->link['online'].'">Users Online: '.$this->online.'</a>';
		} 
		else 
		{
			$this->username = $this->userinfo["name"];
			$this->navigation = 'Welcome '.$this->username.'&nbsp;&bull;&nbsp;
				<a href="'.$this->link['main'].'">Profile Page</a>&nbsp;&bull;&nbsp;
				<a href="'.$this->link['users'].'">Users Page</a>&nbsp;&bull;&nbsp;
				<a href="'.$this->link['logout'].'">Logout</a>&nbsp;&bull;&nbsp;
				<a href="'.$this->link['online'].'">Users Online: '.$this->online.'</a>';
		}
		
	return $this->navigation;
	}
	
	/*
	-------------------------------------------------------------
	Error Template
	-------------------------------------------------------------
	@param	string	$title 		Contains the Error Condensed into a title
	@param	string	$reason		Contains the Base Error
	-------------------------------------------------------------
	*/
	function error($title,$reason)
	{
	return '<html>
				<head>
					<title>vbGORESp - '.$title.'</title>
					<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
					<link rel="stylesheet" type="text/css" href="template/vbgoresp/css/template_error.css" />
				</head>
				<body>
					<div id="container">
						<h1 style="border-bottom: 1px solid #ddd; margin-bottom: 4px;">vbGORE<span style="color:limegreen">SP</span> &rarr; '.$title.'</h1><br />
						<div class="error">'.$reason.'</div><br />
						<div class="note"><font style="color:#990000;">*Note:</font> If you cannot figure out how to get rid of this message or do not understand the error, please message "DarkGrave" <a href="http://vbgoresp.animenetworx.info/forums/">Here</a>. Thank you very much.</div>
				</body>
			</html>';
	}

}

?>