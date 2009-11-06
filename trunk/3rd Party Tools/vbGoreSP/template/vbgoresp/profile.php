<?php
$base = '<script text="text/javascript">
function blocking(nr){if (document.layers){current = (document.layers[nr].display == \'none\') ? \'block\' : \'none\';document.layers[nr].display = current;}else if (document.all){current = (document.all[nr].style.display == \'none\') ? \'block\' : \'none\';document.all[nr].style.display = current;}else if (document.getElementById){vista = (document.getElementById(nr).style.display == \'none\') ? \'block\' : \'none\';document.getElementById(nr).style.display = vista;}}
// -->
</script>
<div class="SiteContainer Apply">
<h1>Your <span style="color:limegreen">vbGORE</span> Profile</h1>
    <div class="About">
		<h2>Profile For: {name}</h2>
		<p>
			<small>
				Description: {descr} <a href="#" onClick="blocking(\'editform\'); return false;" id="edit">Edit?</a>
			</small>
			<form id="editform" action="index.php" method="post" style="margin:0px;padding:0px;display:none;">
				<input type="text" maxlength="62" name="desc" alt="Enter Description and Press Enter!" title="Enter Description and Press Enter!">
				<input type="hidden" name="cd">
				<input type="submit" name="submit" value="submit" style="border:0px;padding:0px;color:limegreen;cursor:pointer;" width="0px" height="0px"> <a href="#" style="color: red" onClick="blocking(\'editform\'); return false;">x</a>
			</form>
			<br />
			{statbar}
		</p>  
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