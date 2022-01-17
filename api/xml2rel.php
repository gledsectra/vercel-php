<?php

function object2array($object) { return @json_decode(@json_encode($object),1); }

//$_FILES['userfile']['tmp_name'] = 'TPL REL.xml';
//echo file_get_contents($_FILES['userfile']['tmp_name']);
//exit;
if(isset($_FILES['userfile'])) 
{
//echo file_get_contents($_FILES['userfile']['tmp_name']);//print_r($_FILES);

$abc = str_split('ABCDEFGHIJKLMNOPQRSTUVWXYZ');
//print_r($abc); //die;
$myXMLData = file_get_contents($_FILES['userfile']['tmp_name']);

$myXMLData = str_replace('ss:','sss',$myXMLData);
$xml=simplexml_load_string($myXMLData) or die("Error: Cannot create object");
$axml = object2array($xml);
//print_r($xml);
//echo '---'."\r\n";
$r = 0;
$rr = 0;

$code = <<<'CODE'
const
  xlCenter = -4108;
  xlRight = -4152;
  xlLeft = -4131;
  xlDown   = -4121;
  xlEdgeLeft  = $00000007;
  xlEdgeRight = $0000000A;
  xlEdgeTop    = $00000008;
  xlEdgeBottom = $00000009;
  xlThin     = $00000002;
  xlInsideHorizontal = $0000000C;
  xlInsideVertical = $0000000B;
  xlSolid = 1;

var
  lQueryA, lQueryB : TADOQuery;  
  lTpRelatorio : string;
  i: Integer;
  lExcel, lSheets: Variant;
  lLinha : Integer;

begin  
  //if f_Contido(lTpRelatorio, ['2']) then
    begin
      //lQueryA := f_CreateADOQuery(FrDmGr.ADOSistema,1);
 
      lExcel := CreateOleObject('Excel.Application');
      lExcel.Workbooks.Add;
      lSheets := lExcel.WorkBooks[1].Sheets[1];
 
      //f_OpenQueryTrans(lQueryA, g_SqlText);
      try
        try
          lLinha := 1;
          lSheets.Cells[lLinha,1] := FrGeradorV2.Relatorios.FieldByName('DESCRICAO').AsString;
CODE;
$dbg = '1';
$log = '';
$negritos = array();
$bordas = array();
while($xml->Styles->Style[$r] != null)
{
  if($xml->Styles->Style[$r]->attributes()['sssID']){
    $estiloId = trim($xml->Styles->Style[$r]->attributes()['sssID']);
    $bold = 0;
    //echo '******';
    //print_r($xml->Styles->Style[$r]->Font->attributes());
    if(isset($xml->Styles->Style[$r]->Font) && isset($xml->Styles->Style[$r]->Font->attributes()['sssBold']))
    {
      $bold = $xml->Styles->Style[$r]->Font->attributes()['sssBold'];
    }
    if($bold) $negritos[] = $estiloId;
    if($dbg)  $log .= '
    verificando estilo ' . $estiloId . ' negrito: '.$bold;

    if(isset($xml->Styles->Style[$r]->Borders) && isset($xml->Styles->Style[$r]->Borders[0]->Border))
    {
        $i = 0;
        //
        foreach($xml->Styles->Style[$r]->Borders[0]->Border as $xxx)
        {
            $borda = ($xml->Styles->Style[$r]->Borders[0]->Border[$i]->attributes()['sssPosition']); 
            $bordas[$estiloId][$i] = $borda; 
            $i++;
            if($dbg)  $log .= '
    verificando estilo ' . $estiloId . ' borda: '.$borda;
        }         
    }
    if(isset($xml->Styles->Style[$r]->Interior))
    {
        $interior[$estiloId] = $xml->Styles->Style[$r]->Interior->attributes()['sssColor'];
        if($dbg)  $log .= '    
    verificando estilo ' . $estiloId . ' interior: '.$interior[$estiloId];    
    }
    if(isset($xml->Styles->Style[$r]->Alignment))
    {
        $alinhamento[$estiloId] = $xml->Styles->Style[$r]->Alignment->attributes()['sssHorizontal'];
        if($dbg)  $log .= '    
    verificando estilo ' . $estiloId . ' alinhamento: '.$alinhamento[$estiloId];    
    }

  }
  $r++;
}
//print_r($bordas);
//exit;
//$log .= print_r($negritos,1);
$r = 0;
while($xml->Worksheet->Table->Row[$r] != null)
{
  $c = 0;
  $cc = 0;
  if($xml->Worksheet->Table->Row[$r]->attributes()['sssIndex']) $rr = $xml->Worksheet->Table->Row[$r]->attributes()['sssIndex'][0] - 1;
  if($dbg) $log .= ('
  processando linha '.($rr + 1));
  while($xml->Worksheet->Table->Row[$r]->Cell[$c] != null){
    if($xml->Worksheet->Table->Row[$r]->Cell[$c]->attributes()['sssIndex']) $cc = $xml->Worksheet->Table->Row[$r]->Cell[$c]->attributes()['sssIndex'][0] - 1;
    $cdata = $xml->Worksheet->Table->Row[$r]->Cell[$c]->Data;
    if($cdata) {
      if($dbg)  $log .= ('
    processando coluna '.$abc[$cc]. ' = ' . $cdata);
      $code .= '
          lSheets.Cells[lLinha,'.($cc + 1).'] := \''.$cdata.'\';';
    }
    $merge = $xml->Worksheet->Table->Row[$r]->Cell[$c]->attributes()['sssMergeAcross'];
    if($merge){
      if($dbg) $log .= ('
      mesclando celulas '.$abc[$cc].($r+1).':'.$abc[($cc+$merge)].($r+1));
      $mergeini = $abc[$cc];
      $mergeend = $abc[($cc+$merge)];
      if($cdata) $code .='
          lExcel.Range[\''.$abc[$cc].'\'+IntToStr(lLinha),\''.$abc[($cc+$merge)].'\'+IntToStr(lLinha)].MergeCells := True; //mesclar celulas';
      else $code .= '
          lExcel.Range[\''.$abc[$cc].'\'+IntToStr(lLinha),\''.$abc[($cc+$merge)].'\'+IntToStr(lLinha)].MergeCells := True; //mesclar celulas';
      $cc+=$merge;
    } 
    $estilo =  trim($xml->Worksheet->Table->Row[$r]->Cell[$c]->attributes()['sssStyleID']);
    if($dbg) $log .= '
    aplicando estilo '. $estilo;
    if( in_array($estilo,$negritos)){
        if($merge)
        {
          $code .= '
          lExcel.Range[\''.$mergeini.'\'+IntToStr(lLinha),\''.$mergeend.'\'+IntToStr(lLinha)].Font.Bold := True;      ';
        }
        else
          $code .= '
          lSheets.Cells[lLinha,'.($cc + 1).'].Font.Bold := True;      ';
    }
    if(isset($bordas[$estilo]))
    {
        foreach($bordas[$estilo] as $borda)
        {
          if($merge)
          {
          $code .= '
          lExcel.Range[\''.$mergeini.'\'+IntToStr(lLinha),\''.$mergeend.'\'+IntToStr(lLinha)].Borders.Item[xlEdge'.$borda.'].Weight := xlThin; ';
          }
          else
          $code .= '  
          lSheets.Cells[lLinha,'.($cc + 1).'].Borders.Item[xlEdge'.$borda.'].Weight := xlThin; ';
        }
    }
    if(isset($interior[$estilo]))
    {
        $hex = str_replace('#','',$interior[$estilo]);
        $invhex = substr($hex,4,2).substr($hex,2,2).substr($hex,0,2);
        //list($red, $green, $blue) = sscanf('#'.$hex, "#%02x%02x%02x");

        if($merge)
        {
          $code .= '
          lExcel.Range[\''.$mergeini.'\'+IntToStr(lLinha),\''.$mergeend.'\'+IntToStr(lLinha)].Interior.Color := $'.$invhex.'; ';
        }
        else
          $code .= '            
          lSheets.Cells[lLinha,'.($cc + 1).'].Interior.Color := $'.$invhex.'; ';
          //lSheets.Cells[lLinha,'.($cc + 1).'].Interior.Pattern := xlSolid ;
          //lSheets.Cells[lLinha,'.($cc + 1).'].Interior.PatternColor := RGB('."$red,$green,$blue".'); ';
    }
    if(isset($alinhamento[$estilo]))
    {

        if($merge)
        {
          $code .= '    
          lExcel.Range[\''.$mergeini.'\'+IntToStr(lLinha),\''.$mergeend.'\'+IntToStr(lLinha)].HorizontalAlignment := xl'.$alinhamento[$estilo].'; ';
        }
        else
          $code .= '  
          lSheets.Cells[lLinha,'.($cc + 1).'].HorizontalAlignment := xl'.$alinhamento[$estilo].'; ';
    }
    $c++;
    $cc++;
  }
  $r++;
  $rr++;
  $code .= "
          inc(lLinha,1);
  ";
}

$code .= <<<'CODE'
      
          {while not lQueryA.Eof do
            begin
            ...
            lQueryA.Next;
            end;}
        except
          g_Abort := True;
          f_Mensagem([ExceptionMessage],0);
          Exit;
        end;
      finally
        //lQueryA.Free;
        //lQueryB.Free;
        lExcel.Columns.AutoFit;
        lExcel.Visible := True;
        if lTpRelatorio = '2' then
        g_Abort := True;
      end;
    end;
end. 
CODE;


?>
<label>Código:</label>
<textarea id="result" style="width:100%; height:600px;font-family:Courier New"><?php echo $code; ?></textarea>
<input type="button" value="COPIAR" onclick="jQuery('#result').select();document.execCommand('copy');">
<br><br>
<label>Log:</label>
<textarea id="log" style="width:100%; height:600px;font-family:Courier New"><?php echo $log; ?></textarea>
<?php

//die;
/*
echo ($xml->Worksheet->Table->Row[0]->Cell[0]->attributes()['sssMergeAcross']);
echo ($xml->Worksheet->Table->Row[1]->Cell[0]->attributes()['sssMergeAcross']);
print_r($xml->Worksheet->Table->Row[2]->Cell[0]->attributes());
print_r($xml->Worksheet->Table->Row[2]->Cell[1]);
print_r($xml->Worksheet->Table->Row[2]->Cell[2]);
print_r($xml->Worksheet->Table->Row[2]->Cell[3]);
print_r($xml->Worksheet->Table->Row[2]->Cell[4] == null);
print_r($xml->Worksheet->Table->Row[5]->Cell[0]);
print_r($xml->Worksheet->Table->Row[5]->Cell[1] == null); //ESTA EM BCO MAS NAO NULO
print_r($xml->Worksheet->Table->Row[5]->Cell[2] == null);
print_r($xml->Worksheet->Table->Row[5]->Cell[3]);
print_r($xml->Worksheet->Table->Row[5]->Cell[4]);
print_r($xml->Worksheet->Table->Row[5]->Cell[5]);
print_r($xml->Worksheet->Table->Row[5]->Cell[6] == NULL);
print_r($xml->Worksheet->Table->Row[8]->Cell[0]);
print_r($xml->Worksheet->Table->Row[8]->attributes()['sssIndex'][0]);
*/
}
//https://drive.google.com/uc?export=download&id=1HTMkJEAtbqG7fWHmEdQ_yKgVVhWldECF //nao esta deixando baixar direto, considera ameaca
//https://drive.google.com/file/d/1HTMkJEAtbqG7fWHmEdQ_yKgVVhWldECF/view?usp=sharing
?>
<style>
  body{ font-family: arial; }
  .file, #fname {
        width: 400px;
        height: 50px;
        background: #fff;
        padding: 4px;
        border: 1px dashed #333;
        position: relative;
        cursor: pointer;
    }

    .file::before {
        content: '';
        position: absolute;
        background: #fff;
        font-size: 20px;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 100%;
        height: 100%;
    }

    .file::after {
        content: 'Arraste aqui o arquivo XML';
        position: absolute;
        color: #000;
        font-size: 20px;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
    }
</style>  
<a href="https://1drv.ms/u/s!ApVG5zwTAUQ5gQJWAlXghwclotaS?e=ZVaTxi" target="_blank">Exemplo de XML (abrir no EXCEL, pode alterar a extensão para xls, para facilitar a abertura, caso desejar)</a>
<br><br>
<form enctype="multipart/form-data" method="post">
<div>
<input name="userfile" type="file" name="file" id="userfile" class="file">
<span id="fname" style="display:none"></span>
</div>
	<br>
<input type="submit">
</form>
<script>
  window.onload=function(){
let file = document.getElementById('userfile');
file.addEventListener('change', function() {
    if(file && file.value) {
        let val = file.files[0].name;
        document.getElementById('fname').innerHTML = "Arquivo selecionado: " + val;
			  document.getElementById('fname').style.display = 'block';
    }
	//alert(val)
});
}
</script> 
