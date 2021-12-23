<?php

header("content-type:text/html;charset=utf-8");

exec("python3 /var/www/html/RecvSend/DminusABC.py 2>&1",$out,$ret);
print_r($out);
print_r($ret);


#$shell="python3 preForA.py";
#$a=exec($shell."2>error.txt",$array,$ret);
#print_r($a);
#echo $return_var;


header("Location: http://47.114.178.105/RecvSend/uploadD");
?>
