<?php

header("Cache-Control: no-store, no-cache, must-revalidate, max-age=0");
header("Cache-Control: post-check=0, pre-check=0", false);
header("Pragma: no-cache");

$ringstatus = file_get_contents('ringing.txt');

//var_dump($ringstatus);

if (strpos($ringstatus,'0') !== false) {
	echo 0;
} else {
	$caller = file_get_contents('caller.txt');

	$caller = (explode("\r",$caller));

	echo substr($caller[1], 10);
	$fp = fopen('ringing.txt', 'w');
	fwrite($fp, '0');
	fclose($fp);
}

?>