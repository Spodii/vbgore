<?php
$base = '<html>
	<head>
		<title>vbGORE SP - vbGORE PHP Panel</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link rel="stylesheet" type="text/css" href="http://{base_url}/template/{template}/css/default_style.css" media="screen" />
	</head>
<body>
<div align="right">
	<h5>
		{navigation}
		<form action="search.php" method="GET">
			Quick Search: <input name="q" type="text" size="20" alt="Enter Term and press Enter" title="Enter Term and press Enter"> <input type="submit" name="s" value="process" width="0" height="0" style="display:none;visibility:hidden;">
		</form>
	</h5>
</div>';
?>