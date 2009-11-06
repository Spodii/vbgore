<?php
$base = '<div class="SiteContainer Apply">
<h1>vbGORE <span style="color:limegreen">Login</span> Form</h1>
	<div class="About">
		<h2>About Logging In</h2>
			<p><strong>This form will not log you in INGAME.</strong><br />Only to check your profile, edit description, etc.</p>
			<p><a href="{r_link}">Back to the register</a></p>
	</div>
	<div id="Form" class="ApplyForm">
		<fieldset>
			<form id="ApplicationForm" action="login.php" method="post">
			<ul>
				<li>
					<label for="username">Username:</label>
					<input type="text" size="30" class="Input" name="username" />
				</li>
				<li>
					<label for="password">Password:</label>
					<input type="password" size="30" class="Input" name="password" />
				</li>
				<li>
					<label for="rememberme">Remember me: </label>
					<input type="checkbox" name="rememberme" value="yes" />
				</li>
			</ul>
			<div class="Submit"><input type="submit" name="submit" class="Button" value="Log In" /></div>
			</form>
        </fieldset>
	</div>
</div>';
?>