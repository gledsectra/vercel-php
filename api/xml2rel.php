<?php

function object2array($object) { return @json_decode(@json_encode($object),1); }

//$_FILES['userfile']['tmp_name'] = 'TPL REL.xml.xls';
//echo file_get_contents($_FILES['userfile']['tmp_name']);
//exit;
if(isset($_FILES['userfile'])){
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
  cCenter = -4108;
  cLeft = -4131;
  cRight = -4152;
  cDown = -4121;
  cTop = -4160;
  cEdgeBottom = $00000009;
  cEdgeLeft = $00000007;
  cEdgeRight = $0000000A;
  cEdgeTop = $00000008;
  cHairline = $00000001;
  cMedium = $FFFFEFD6;
  cThick = $00000004;
  cThin = $00000002;
  cContinuous = 1;
  cDot = -4118;
  cDashDotDot = 5;
  cDashDot = 4;
  cDash = -4115;
  cSlantDashDot = 13;
  cDouble = -4119;
  cLineStyleNone = -4142;
  cInsideHorizontal = $0000000C;
  cInsideVertical = $0000000B;
  cSolid = 1;

var
  lQueryA, lQueryB: TADOQuery;  
  lExcel, lSheets: Variant;
  i, lLinha: Integer;
  lTpRelatorio: string;

procedure f_Setar(pSheet: Variant; pColunaIni, pColunaFim: String; pLinha: Integer; pValor: Variant);
  begin
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)] := pValor; 
  end;

procedure f_Mesclar(pSheet: Variant; pColunaIni, pColunaFim: String; pLinha: Integer);
  begin
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)].MergeCells := True;
  end;

procedure f_Negritar(pSheet: Variant; pColunaIni, pColunaFim: String; pLinha: Integer);
  begin
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)].Font.Bold := True;
  end;

procedure f_Alinhar(pSheet: Variant; pColunaIni, pColunaFim: String; pLinha: Integer; pAlinhamento: Variant);
  begin
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)].HorizontalAlignment := pAlinhamento; //ex: cCenter
  end;

procedure f_Interior(pSheet: Variant; pColunaIni, pColunaFim: String; pLinha: Integer; pCor: Variant);
  begin
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)].Interior.Color := pCor; 
  end;  

procedure f_SetarBorda(pSheet: Variant; pColunaIni, pColunaFim: String; pLinha: Integer; pAlinhamento, pEstilo, pEspessura: Variant);
  begin
    // Format. Contorno // cContinuous / cDot / cDashDot / cDashDotDot / cDash / cSlantDashDot / cDouble / cLineStyleNone
    // Format. Espessura // cThin / cMedium / cThick
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)].Borders.Item[pAlinhamento].LineStyle := pEstilo;
    pSheet.Range[pColunaIni+IntToStr(pLinha), pColunaFim+IntToStr(pLinha)].Borders.Item[pAlinhamento].Weight := pEspessura;
  end;  

begin 

  lVetorSistema :=
    VarArrayOf([
                 VarArrayOf(['', 'T', 'Tp. Relatório', 'N', 'Igual', 'E3;1=Relatório;2=Excel;3=Ambos','']),
                 //VarArrayOf(['', 'T', 'Aaaa', 'S', 'Contido', '', '']),
                 //VarArrayOf(['', 'D', 'Dt.', 'S', 'Entre', '', '']),
                 //VarArrayOf(['', 'T', 'Cliente', 'S', 'Contido', 'P@Tabela de Clientes@FN_FORNECEDORES@CODIGO@RAZAO@Código@Razão', '(CKCLIENTE=''S'') AND (ATCLIENTE = ''S'')']),
                ]);


  if not f_AssistenteFiltro(g_Filtro, g_Ordem, g_Legenda, lVetorSistema, True, 'R', False, 'VE_PEDIDO',
                             g_NroRelat, '', g_Id) then
    begin
      g_Abort := True;
      Exit;
    end;

  lTpRelatorio := f_ExtractValue(lVetorSistema[0]);

  {
    Titulo := FrGeradorV2.Relatorios.FieldByName('DESCRICAO').AsString;
    (...).NumberFormat := 'R$ ###.##0,00';
    lSheets.Columns('B:B').ColumnWidth := 13.15;
  }

  if f_Contido(lTpRelatorio, ['2']) then
    begin
      //lQueryA := f_CreateADOQuery(FrDmGr.ADOSistema,1);
 
      lExcel := CreateOleObject('Excel.Application');
      lExcel.Workbooks.Add;
      lSheets := lExcel.WorkBooks[1].Sheets[1];

      try
        //f_OpenQueryTrans(lQueryA, g_SqlText);
        try
          //if lQueryA.RecordCount = 0 then Exit;
 
          lLinha := 1;

CODE;
$dbg = '1';
$log = '';
$negritos = array();
$bordas = array();
while($xml->Styles->Style[$r] != null)
{
    //echo '<pre>---';
    //var_dump($xml->Styles->Style[$r]);
  if($xml->Styles->Style[$r]->attributes()['sssID']){
    $estiloId = trim($xml->Styles->Style[$r]->attributes()['sssID']);
    $bold = 0;
    //echo '******';
    //print_r($xml->Styles->Style[$r]->Font->attributes());
    if(isset($xml->Styles->Style[$r]->Font) && isset($xml->Styles->Style[$r]->Font->attributes()['sssBold']))    {
      $bold = $xml->Styles->Style[$r]->Font->attributes()['sssBold'];
    }
    if($bold) $negritos[] = $estiloId;
    if($dbg)  $log .= '
    verificando estilo ' . $estiloId . ' negrito: '.$bold;

    if(isset($xml->Styles->Style[$r]->Borders) && isset($xml->Styles->Style[$r]->Borders[0]->Border))    {
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
    if(isset($xml->Styles->Style[$r]->Interior))    {
        $interior[$estiloId] = $xml->Styles->Style[$r]->Interior->attributes()['sssColor'];
        if($dbg)  $log .= '    
    verificando estilo ' . $estiloId . ' interior: '.$interior[$estiloId];    
    }
    if(isset($xml->Styles->Style[$r]->Alignment))    {
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
          f_Setar(lSheets, \''.$abc[$cc].'\', \''.$abc[$cc].'\', lLinha, \''.$cdata.'\');'; 
          //lSheets.Cells[lLinha,'.($cc + 1).'] := \''.$cdata.'\';';
    }
    $merge = $xml->Worksheet->Table->Row[$r]->Cell[$c]->attributes()['sssMergeAcross'];
    if($merge){
      if($dbg) $log .= ('
      mesclando celulas '.$abc[$cc].($r+1).':'.$abc[($cc+$merge)].($r+1));
      $mergeini = $abc[$cc];
      $mergeend = $abc[($cc+$merge)];
      $code .='
          f_Mesclar(lSheets, \''.$mergeini.'\', \''.$mergeend.'\', lLinha);';
      $cc+=$merge;
    } 
    else $mergeini = $mergeend = $abc[$cc];
    $estilo =  trim($xml->Worksheet->Table->Row[$r]->Cell[$c]->attributes()['sssStyleID']);
    if($dbg) $log .= '
    aplicando estilo '. $estilo;
    if( in_array($estilo,$negritos)){       
          $code .= '
          f_Negritar(lSheets, \''.$mergeini.'\', \''.$mergeend.'\', lLinha);';
    }
    if(isset($bordas[$estilo]))    {
        foreach($bordas[$estilo] as $borda)
        {       
          $code .= '
          f_SetarBorda(lSheets, \''.$mergeini.'\', \''.$mergeend.'\', lLinha, cEdge'.$borda.', cContinuous, cThin);';          
        }
    }
    if(isset($interior[$estilo]))    {
        $hex = str_replace('#','',$interior[$estilo]);
        $invhex = substr($hex,4,2).substr($hex,2,2).substr($hex,0,2);
        //list($red, $green, $blue) = sscanf('#'.$hex, "#%02x%02x%02x");
       
        $code .= '
          f_Interior(lSheets, \''.$mergeini.'\', \''.$mergeend.'\', lLinha, $'.$invhex.');';           
    }
    if(isset($alinhamento[$estilo]))    {
          $code .= '  
          f_Alinhar(lSheets, \''.$mergeini.'\', \''.$mergeend.'\', lLinha, c'.$alinhamento[$estilo].');';           
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
<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.6.0/jquery.min.js" integrity="sha512-894YE6QWD5I59HgZOGReFYm4dnWc1Qt5NtvYSaNcOP+u1T9qYdvdihz0PPSiiqn/+/3e7Jo4EaG7TubfWGUrMQ==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<label>Código:</label>
<textarea id="result" style="width:100%; height:600px;font-family:Courier New"><?php echo $code; ?></textarea>
<input type="button" value="COPIAR" onclick="jQuery('#result').select();document.execCommand('copy');">
<br><br>
<label>Log:</label>
<textarea id="log" style="width:100%; height:600px;font-family:Courier New"><?php echo $log; ?></textarea>
<?php

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
