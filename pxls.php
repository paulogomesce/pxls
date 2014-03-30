<?php
class pxls{

	public $bgColor; //Cor de fundo da celula
	public $size;    //tamanho da font
	public $color;   //Cor da font
	
	//Define a formataÃ§Ã£o para o texto e para a cÃ©lula
	function setFormatacao($bgColor,$size,$color){
		$this->bgColor = $bgColor;
		$this->size    = $size;
		$this->color   = $color;
	}

	//Inicia o arquivo Excel, recomendÃ¡vel utilizar undeline(_) ao invÃ©s de espaÃ§o
	function abrirArquivo($nmArquivo){
		header("Content-type: application/msexcel");
		header("Content-Disposition: attachment; filename={$nmArquivo}.xls");
	}

	//Insere quebras de linha, o parÃ¢metro Ã© a quantidade de linhas que serï¿½o quebradas
	function quebraLinha($qtdQuebras="0"){
		for($i=0;$i<=$qtdQuebras;$i++){
			echo "<br />";
		}
	}
	
	//Adiciona um tÃ­rulo para a tabela
	function addTitulo($txtTitulo, $pAlign="L"){
		$align = "L";
		
		switch(strtoupper($pAlign)){
			case "C":
				$align = "center";
				break;
			case "R":
				$align = "right";
				break;
			default:
				$align = "left";
		}
	
		echo "<div align={$align}><b><font size={$this->size} color={$this->color}>".$txtTitulo."</font></b></div>";
		$this->quebraLinha();
	}

	//adiciona subtitulos para a tabela
	function addSubTitulo($txtTitulo, $pAlign="L"){
		$align = "L";
	
		switch(strtoupper($pAlign)){
			case "C":
				$align = "center";
				break;
			case "R":
				$align = "right";
				break;
			default:
				$align = "left";
		}
	
		echo "<div align={$align}><font size={$this->size} color={$this->color}>".$txtTitulo."</font></div>";
		$this->quebraLinha();
	}
	
	//Adiciona o cabeÃ§alho da tabela
	function addCabecalho($valor, $pAlign="C", $colspan=0){
		$align = "C";
		
		switch(strtoupper($pAlign)){
			case "L":
				$align = "left";
				break;
			case "R":
				$align = "right";
				break;
			default:
				$align = "center";
		}
		
		if(is_array($valor)){
			foreach ($valor as $coluna){
				echo "<th align={$align} bgcolor={$this->bgColor} colspan={$colspan} ><font size={$this->size} color={$this->color} >".$coluna."</font></th>";
			}
		}else{
			echo "<th align={$align} bgcolor={$this->bgColor} colspan={$colspan} ><font size={$this->size} color={$this->color} >".$valor."</font></th>";
		}
	}
	
	//Inicia a abertura de uma tabela
	function abreTabela($border="0"){
		echo "<table border=".$border.">";
	}

	//Indica o fim da tabela
	function fechaTabela(){
		echo "</table>";
	}
	
	//Inicia uma linha para tabela
	function abreLinha(){
		echo "<tr>";
	}

	//Indica o fim de uma linha aberta anteriormente
	function fechaLinha(){
		echo "</tr>";
	}

	//Adiciona uma cÃ©lula
	function addCel($valor="", $pAlign="C"){
		$align = "C";
		switch(strtoupper($pAlign)){
			case "L":
				$align = "left";
				break;
			case "R":
				$align = "right";
				break;
			default:
				$align = "center";
		}
		if(is_array($valor)){
			foreach ($valor as $valorCelula){
				echo "<td align={$align}><font size={$this->size} color={$this->color}>".$valorCelula."</font></td>";
			}
		}else{
			echo "<td align={$align}><font size={$this->size} color={$this->color}>".$valor."</font></td>";
		}
	}

	//Mesclar CÃ©lulas
	function addColSpan($valor, $qtdColl, $pAlign="C"){
		$align = "C";		
		switch(strtoupper($pAlign)){
			case "L":
				$align = "left";
				break;
			case "R":
				$align = "right";
				break;
			default:
				$align = "center";
		}		
		echo "<td colspan={$qtdColl} align={$align}><font size={$this->size} color={$this->color}>".$valor."</font></td>";
	}
}
?>