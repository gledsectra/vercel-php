<style> body { font-family: arial; }</style>
<?php
if(isset($_POST['query'])) 
{
$code = $_POST['query'];
$code=str_replace('&#34;','',$code);
$ncode = $code;

$ncode = explode('Text=',$ncode);
array_shift($ncode);
foreach($ncode as $item)
{
    $itemx = explode('"',$item)[1]; 
    $tipo = 'AsString;';
    if(stristr($itemx,'VLR')) $tipo = 'AsFloat;'; 
    if(stristr($itemx,'Campos.')) {
      $itemx = str_replace('[Campos.',"lQueryA.FieldByName(&#39;", $itemx)."&#39;).".$tipo;
      $itemx = str_replace("]","",$itemx);
    }
    $arr[] = str_replace('\\','',$itemx);
}
$ncode = implode("\r\n",$arr); 

$arr = null;

}

if(isset($ncode)) { ?>

<label>CAMPOS ENCONTRADOS:</label>
<textarea id="result" style="width:100%; height:600px;font-family:Courier New"><?php echo $ncode; ?></textarea>
<input type="button" value="COPIAR" onclick="jQuery('#result').select();document.execCommand('copy');">
<br><br>
<?php } ?>
<form method="post">
<label>COLE OS ITENS DO LAYOUT:</label>
<textarea id="query" name="query" style="width:100%; height:200px;font-family:Courier New"><?php if(isset($ncode)) echo str_replace('\"','"',$code);  ?></textarea>
<input type="submit" value="EXTRAIR">
<input type="button" onclick="jQuery('#query').val('')" value="LIMPAR">
<input type="button" value="COPIAR" onclick="jQuery('#query').select();document.execCommand('copy');">
</form>
