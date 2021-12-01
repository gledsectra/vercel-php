<style>body{ font-family: arial }</style>
<?php
if(isset($_POST['query'])) 
{
$sql = $_POST['query'];

//echo $sql;
$sql = str_replace(',',"\r\n,",$sql);
$sql = str_replace('FROM',"\r\nFROM",$sql);
$sql = str_replace(' AND ',"\r\nAND ",$sql);

$aSql = explode("\r\n",$sql);

foreach ($aSql as $line){
    $naSql[] = "'".trim($line)." '".'+';
}
$sql = implode("\r\n",$naSql);
$sql = str_replace(",\r\n",", ",$sql);
$sql = str_replace(",'+\r\n'",", ",$sql);
$sql = str_replace(", '+\r\n'",", ",$sql);
$sql = str_replace(", '+\r\n'",", ",$sql);
$sql = str_replace("' '+\r\n","",$sql);
$sql = str_replace("'","' ",$sql);
$sql = str_replace("' ,","'      ,",$sql);
$sql = str_replace(",  ",", ",$sql);
$sql = str_replace("( ","(",$sql);
$sql = str_replace("' FROM","'   FROM",$sql);
$sql = str_replace("' WHERE","'  WHERE",$sql);
$sql = str_replace("' AND","'    AND",$sql);
$sql = str_replace("' ORDER","'  ORDER",$sql);
//$sql = str_replace("' )","'        )",$sql);
$sql = "  g_SqlText :=\r\n".substr($sql,0);
$sql = substr($sql,0,-2).';';
$sql = str_replace("\'","''",$sql);
$sql = str_replace(" '' "," ''",$sql);
$sql = str_replace("'' )","'')",$sql);

$aSql = explode("\r\n",$sql);
$sql = implode("\r\n  ",$aSql);
//echo "<pre>".print_r($sql,1). "</pre>";
//die;
//print_r($naSql);
}

if(isset($sql)) { ?>
<label>QUERY FORMATADA:</label>
<textarea id="result" style="width:100%; height:600px;font-family:Courier New"><?php echo $sql; ?></textarea>
<input type="button" value="COPIAR" onclick="jQuery('#result').select();document.execCommand('copy');">
<br><br>
<?php } ?>
<form method="post">
<label>QUERY SQL:</label>
<textarea id="query" name="query" style="width:100%; height:200px;font-family:Courier New"><?php if(isset($sql)) echo str_replace("\'","'",$_POST['query']);  ?></textarea>
<input type="submit" value="FORMATAR">
<input type="button" onclick="jQuery('#query').val('')" value="LIMPAR">
<input type="button" value="COPIAR" onclick="jQuery('#query').select();document.execCommand('copy');">
</form>