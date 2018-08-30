<?php
/*
# Requerimiento: Necesario descargar la librería PHPExcel
*/
require_once ("Classes/PHPExcel.php");
require_once ("ODBCExcel/ODBCExcelReport.php");
#Genera el excel con el resultado arrojado de la consulta 
$sql ="
	SELECT
		*
	 FROM NombreTabla
	 
	";
#Configuracion de  la base de datos
$arrayConfigDB=array("serverDB"=>"Servidor de base","nameDB"=>"NombreBase","userDB"=>"UsuarioBase","passwordDB"=>"contraseñaUsuarioBase");
$reporteExcel=new ODBCExcelReport("File Name.xlsx","Titulo Tabla","Titulo Hoja",$arrayConfigDB,1);
ob_start();
#Genera Excel de la consulta realizada
$reporteExcel->ExcelReportGenerate($sql);
$xlsData = ob_get_contents();
ob_end_clean();
$resultado = array('code'=>0, 'msg'=>'','filename'=>'File Name','filedata' =>"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,".base64_encode($xlsData));
echo json_encode($resultado);
?>