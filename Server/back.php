<?php
$ip = $_SERVER["REMOTE_ADDR"];
$paraValue = $_POST['MENU'];
$nfs=fopen("result.txt","ab");
flock($nfs,LOCK_EX);
fwrite($nfs,"$ip|$paraValue\r\n",strlen("$ip|$paraValue\r\n"));
flock($nfs,LOCK_UN);
fclose($nfs);
echo("love $paraValue");
?>
loveyou