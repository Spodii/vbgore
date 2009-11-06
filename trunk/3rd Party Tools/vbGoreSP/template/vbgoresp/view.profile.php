<?php
$base = '<div class="SiteContainer Apply">
<h1>Viewing <span style="color:limegreen">{name}\'s</span> Profile</h1>
    <div class="About">
		<h2>Profile For: {name}</h2>
		<p><small>Description: {descr}</small><br /><br />{statbar}<br /><br />User Is: <b><a href="{link_online}">{server_status}</a></b></p>	  
	</div>
	<div id="Form" class="ApplyForm">
        <fieldset>
			<label>User Level:</label> {stat_elv}<br />
			<label>Gm Level:</label> {gm}<br />
			<label>Gold:</label> {stat_gold}<br />
			<label>Exp:</label> {stat_exp}<br />
			<label>Strength:</label> {stat_str}<br />
			<label>Agility:</label> {stat_agi}<br />
			<label>Defense:</label> {stat_def}<br />
			<label>Magic Def:</label> {stat_mag}<br />
		</fieldset>
	</div>
</div>';
?>