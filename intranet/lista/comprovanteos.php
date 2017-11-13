<?php 

try {
    $conn = new PDO('sqlsrv:Server=SMARTNEW;Database=vdlap', 'sa', '@fgh55qdy');
	$conn->exec("SET CHARACTER SET utf8");
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch(PDOException $e) {
    echo 'ERROR: ' . $e->getMessage();
}



$consulta = $conn->query("SELECT * FROM TBL_OrdemServico2 WHERE num_os = 2366");
$linha = $consulta->fetch(PDO::FETCH_ASSOC);

echo "Numero da OS: {$linha['num_os']}<br>";
 echo "Entregue por: {$linha['nm_solicitante']}<br>";
  echo "Nome da unidade: {$linha['nm_unidade']}<br>" ;
    echo "Data da OS: {$linha['dt_os']}<br>" ;
	 echo "Data da entrega:{$linha['dt_entrada']}<br>" ;
	 echo "Equipamento: {$linha['cd_equipamento']}";