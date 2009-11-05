<?php
$base = '<div class="SiteContainer Apply">
<h1>Vbgore <span style="color:limegreen">Registration</span> Form</h1>
	<div class="About">	
		<h2>Why Register?</h2>
		<p>You get more benefits, A profile, game stuff. ect w/e you want can go here. This is just a demo.</p>
	</div>
	<div id="Form" class="ApplyForm">
    <fieldset>
	<form id="frmSignIn" action="register.php" method="post">
		<ul>
			<li>
				<label for="name">Name: </label>
				<input type="text" name="name" size="30" maxlength="32" alt="Name must be 4 to 30 characters long + alphanumerical" title="Name must be 4 to 30 characters long & alphanumerical" />
			</li>
			<li>
				<label for="desc">Description: </label>
				<input type="text" name="desc" size="30" maxlength="32" alt="Description must be 32 characters long." title="Description must be 32 characters long." />
			</li>
            <li>
				<label for="class">Class: </label>
				<select name="class" style="width:56%">';
?>