<html>
<form action="sink.php" method="post">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<head>
<title>serve</title>
</head>
<?php
if (file_exists("install.php")) {
	header("location:install.php");
}
?>
<p>Please enter your information to complete this class'information</p>
</br>
<tr>
    <td>序号</td>
    <td align="center"><input type="text"name="number"size="14"></td>
</tr>
</br>
<tr>
    <td>时间</td>
    <td align="center"><input type="text"name="time"size="14"></td>
</tr>
</br>
<tr>
    <td>命令</td>
    <td align="center"><input type="text"name="command"size="24"></td>
</tr>
</br>
<tr>
    <td>语句</td>
    <td align="center"><input type="text"name="shell"size="24"></td>
</tr>
</br>
<tr>
<input type="submit" value="Submit">
</tr>
</form>
</html>