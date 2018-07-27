<?php

date_default_timezone_set('Europe/Amsterdam');

$setting = file_get_contents("https://infoscherm.gewis.nl/backoffice/api.php?query=alcoholtijd");
if(!$setting) {
	$setting = "14:00";
}
//$currenttime = time();
//$setting = strtotime($setting) + 7200;

if($_GET['quirk']) {
	$script = "C:\quirkmode.ps1";
	echo "Quirk mode";
	$query = shell_exec("powershell -command $script < NUL");
	die();
}

//$difference = $currenttime - $setting;
//var_dump($currenttime );

$currenttime = date('G');
$setting = substr($setting, 0, 2);

$lm_begin = ($setting * 3600) - 60;
$lm_end = $setting * 3600;

$current_exact = (date('G') * 3600) + (date('i') * 60) + date('s');

var_dump($currenttime);
var_dump($setting);
var_dump($current_exact);
//die();

if ($currenttime >= $setting) {
	$script = "C:\borreltijd_ja.ps1";
	echo 'Groen';
} elseif (($current_exact >= $lm_begin) && ($current_exact < $lm_end)) {
	$script = "C:\borreltijd_lastminute.ps1";
	echo 'Laatste minuut';
} else {
	$script = "C:\borreltijd_nee.ps1";
	echo 'Rood';
} 

$query = shell_exec("powershell -command $script < NUL");

?>