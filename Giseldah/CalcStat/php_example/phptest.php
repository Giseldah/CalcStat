<html>
 <head>
  <title>CalcStat PHP Test</title>
 </head>
 <body>
<?php
include 'calcstat.php';

$csn = 'OutHealPRatP';
$cl = 130;
$cn = 479664;
echo '<p>Calculate Outgoing Healing percentage from '.$cn.' Outgoing Healing Rating at player level '.$cl.':</p>'.Chr(13).Chr(10);
echo '<p>result: CalcStat("'.$csn.'",'.$cl.','.$cn.') = '.CalcStat($csn,$cl,$cn).'</p>';
?> 
 </body>
</html>