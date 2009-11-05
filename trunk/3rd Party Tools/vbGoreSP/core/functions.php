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

// Functions - For Random Things!
	function stripslashes_deep($value)
	{
	    
	   $value = is_array($value) ?
	               array_map('stripslashes_deep', $value) :
	               stripslashes($value);
	   return $value;
	   
	}

	function addslashes_deep($value)
	{
	    
	   $value = is_array($value) ?
	               array_map('addslashes_deep', $value) :
	               addslashes($value);
	   return $value;
	   
	}

	function html_deep($value)
	{
	    
	   $value = is_array($value) ?
	               array_map('html_deep', $value) :
	               htmlspecialchars($value);
	   return $value;
	   
	}
	
	function check_status($value)
	{
		return 
		(
			$value == 1 ? 
				"Online" : 
				"Offline"
		);
	}
	
	function check_description($value)
	{
		return
		(
			$value == "" ? 
				"None" : ($value == null ?
				"None" :
				$value
		));
	}
	
	function check_gm($value)
	{
		return
		(
			$value == 0 ? 
				"User" : 
				"GM"
		);
	}
	
	function str_cut($string, $max_length)
	{
		return 
		(
			strlen($string) > $max_length ? 
				substr
				(
					$string, 0, $max_length-strlen
					(
						'...'
					)
				).'...' : 
			$string
		); 
	}

?>