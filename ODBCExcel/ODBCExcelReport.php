<?php
/**
 * @autor       Juan Carlos Hdz Hdz
 * @category    PHPExcel
 * @version     1.0, 2016-05-02
 */
require_once("ODBC.php");
class ODBCExcelReport extends PHPExcel
{
	private $campos=array();
	private $nametitle="";
	private $fontcolortitle="";
	private $fontsizetitle="";
	private $fonttype="";
	private $fontcolorcamapos="";
	private $fontsizecamapos="";
	private $fonttypecamapos="";
	private $numcampos="";
	private $filename="";
	private $objPHPExcel=null;
	private $rownum=0;
	private $row=0;
	private $arrColumn;
	private $titlehoja;
	private $table_striped;
	private $msg;
	private $ArrayMsg;
	private $tipomsg;
	private $configDataBase;
	private $arrayRecordExterno;
	
	public function __construct($filename="Reporte",$titletable="",$titlehoja="",$configDataBase=0,$rownum=1,$ArrayMsg=array(0=>"Error",1=>"Warning",2=>"Sin Resultados",3=>"Success"))
	{
		//Creates a new object PHPExcel
		parent::__construct(); 
		$this->arrColumn=array("1"=>"A","2"=>"B","3"=>"C","4"=>"D","5"=>"E","6"=>"F","7"=>"G","8"=>"H","9"=>"I","10"=>"J","11"=>"K","12"=>"L","13"=>"M","14"=>"N","15"=>"O","16"=>"P","17"=>"Q","18"=>"R","19"=>"S","20"=>"T","21"=>"U","22"=>"V","23"=>"W","24"=>"X","25"=>"Y","26"=>"Z");
		$this->rownum=$rownum;
		$this->titlehoja=$titlehoja;
		$this->msg="";
		$this->ArrayMsg=$ArrayMsg;
		$this->tipomsg=0;
		$this->filename=$filename;
		$this->nametitle=$titletable;
		$this->configDataBase=$configDataBase;
		$this->arrayRecordExterno=0;
		$this->row=1;
	}
	/****************************************
	
	*****************************************/
	public function initialize($filename,$titlehoja,$titletable,$campos)
	{
		$this->campos=$campos;
		$this->nametitle=$titletable;
		$this->numcampos=count($campos);
		$this->filename=$filename;
		//$this->filename=($filename)?((get_extension($filename)=="xlsx"||(get_extension($filename)=="xls"))?$filename:$filename.".xlsx"):"SinNombre.xlsx";
		
		$this->titlehoja=$titlehoja;
		//get_extension($filename);
		// Establecer propiedades
		$this->getProperties()
		->setCreator("Juan Carlos Hdz")
		->setLastModifiedBy("Juan Carlos Hdz")
		->setTitle("Documento Excel de Reporte")
		->setSubject("Documento Excel de Reporte")
		->setDescription("Creacion de archivos de Excel.")
		->setKeywords("Excel Office 2007 openxml php")
		->setCategory("Reportes de Excel");
		// Renombrar Hoja
		$this->getActiveSheet()->setTitle($this->titlehoja);
		// Establecer la hoja activa, para que cuando se abra el documento se muestre primero.
		$this->setActiveSheetIndex(0);
		$this->theadExcel();
		
	}
	/****************************************
	
	*****************************************/
	private function propertiesExcel()
	{
		$this->numcampos=0;
		// Establecer propiedades
		$this->getProperties()
		->setCreator("Juan Carlos Hdz")
		->setLastModifiedBy("Juan Carlos Hdz")
		->setTitle("Documento Excel de Reporte")
		->setSubject("Documento Excel de Reporte")
		->setDescription("Creacion de archivos de Excel.")
		->setKeywords("Excel Office 2007 openxml php")
		->setCategory("Reportes de Excel");
		// Renombrar Hoja
		$this->getActiveSheet()->setTitle($this->titlehoja);
		// Establecer la hoja activa, para que cuando se abra el documento se muestre primero.
		$this->setActiveSheetIndex(0);
	}
	/****************************************
	
	*****************************************/
	private function AddColumn($rsdate)
	{
		if(!count($this->campos))
		{
			$this->campos["#"]=array("ColumName"=>"N°","Type"=>"");
			foreach($rsdate as $nombrecampo=>$infodat) 
			{
				$arrnombre=explode('_',$nombrecampo);
				if(count($arrnombre)>1)
				{
					$nombre="";
					foreach($arrnombre as $clave=>$valname)
					{
						if($nombre)
							$nombre.=" ";
						$nombre.=$valname;
					}
					$this->campos[$nombrecampo]=array("ColumName"=>$nombre,"Type"=>$infodat);
				}
				else
					$this->campos[$nombrecampo]=array("ColumName"=>$nombrecampo,"Type"=>$infodat);
			}
			///para agregar subtitle
			/*$bandtitlecol=false;
			foreach($rsdate as $nombrecampo=>$infodat) 
			{
				$nameCampo=$nombrecampo;
				$colExt="";
				$arrnombreSubTitle=explode('@',$nombrecampo);
				if(count($arrnombreSubTitle)>1)
				{
					$colExt=$arrnombreSubTitle[0];
					$nameCampo=$arrnombreSubTitle[1];
					$bandtitlecol=true;
				}
				else if(count($arrnombreSubTitle))
				{
					$nameCampo=$arrnombreSubTitle[0];
				}
				$arrnombre=explode('_',$nameCampo);
				if(count($arrnombre)>0)
				{
					$nombre="";
					foreach($arrnombre as $clave=>$valname)
					{
						if($nombre)
							$nombre.=" ";
						$nombre.=$valname;
					}
					$this->campos[$nombrecampo]=array("ColExt"=>$colExt,"ColumName"=>$nombre,"Type"=>$infodat);
				}
				else
					$this->campos[$nombrecampo]=array("ColExt"=>$colExt,"ColumName"=>$nombrecampo,"Type"=>$infodat);
			}
			if(!$bandtitlecol){
				foreach($this->campos as $clave=>$arr)
				{
					unset($this->campos[$clave]["ColExt"]);
				}
			}*/
		}
		$this->numcampos=count($this->campos);
		$this->theadExcel();
	}
	/****************************************
	
	*****************************************/
	private function get_extension($str) 
	{
		//return end(explode(".", $str));
		$arrarch=explode(".", $str);
		return $arrarch[count($arrarch)-1];
	}
	/****************************************
	
	*****************************************/
	private function GetColumn($id)
	{
		if(!$id)
			return "A";
		return $this->arrColumn[$id];
	}
	/***************************************
	
	****************************************/
	private function titleTable()
	{
		$this->setActiveSheetIndex(0)
		->setCellValue($this->GetColumn(1).$this->rownum, $this->nametitle);
		//color font
		$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos).$this->rownum)->getFont()->setBold(true)
		->setName('Arial')
		->setSize(15)
		->getColor()->setRGB('ffffff');
		///color celda
		$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos).$this->rownum)
		->getFill()
		->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
		->getStartColor()
		->setRGB('333399');

		$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos).$this->rownum)->getBorders()
		->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
		$this->setActiveSheetIndex(0)->mergeCells($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos).$this->rownum);
		
		$this->getActiveSheet()->getStyle(($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos)).$this->rownum)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
		
		$this->rownum++;
	}
	/**************************************************
	
	***************************************************/
	private function theadExcel()
	{
		
		$this->titleTable();
		$this->setActiveSheetIndex(0);
		$col=1;
		
		foreach($this->campos as $key=>$nombrecampo) {
			$namecampo="";
			if(is_array($nombrecampo)&&count($nombrecampo))
				$namecampo=$nombrecampo["ColumName"];
			else
				$namecampo=$nombrecampo;
			$this->getActiveSheet()->setCellValueByColumnAndRow($col-1,$this->rownum,$namecampo);
			$this->getActiveSheet()->getColumnDimension($this->GetColumn($col))->setAutoSize(true);
			
			$this->getActiveSheet()->getStyle($this->GetColumn($col).$this->rownum)->getFont()->setBold(true)
			->setName('Arial')
			->setSize(10)
			->getColor()->setRGB('FFFFFF');
			$col++; 
		}
		///color cell
		$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)
		->getFill()
		->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
		->getStartColor()
		->setRGB('333399');
		
		$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
		$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)->getBorders()->getAllBorders()->getColor()->setARGB('FFFFFF') ; // Color del marco 
		
		$this->rownum++;
	}
	/************************************************
	Add Mensagge
	tipo
	0=error
	1=warning
	2=info
	3=success
	*************************************************/
	private function AddMsg($msg="",$tipo=0)
	{
		$this->msg=$msg;
		$this->tipomsg=$tipo;
	}
	/************************************************
	
	*************************************************/
	private function tbodyExcel($rsdate)
	{
	}
	/**************************************************
	
	***************************************************/
	///$styleCell=array("fontcolor"=>"000","FillColor"=>"FFFF","fontBlonde"=>"True")
	private function AddRowtbodyExcel($rsdate,$fontcolor="",$rowcolor="",$styleCell=0)//,$stylecampos,$stylerow
	{
		if(count($rsdate))
		{
			if(count($this->campos))
			{
				$col=1;
				foreach($this->campos as $idcampo=>$datcampo)
				{
					$typecampo="";
					$valdato="";
					if(is_array($datcampo)&&count($datcampo))
						$typecampo=$datcampo["Type"];
					if($typecampo=="money")
					{
						if($rsdate[$idcampo]>=1000)
							$valdato="$".number_format($rsdate[$idcampo], 2, '.', ',');
						else
							$valdato="$".number_format($rsdate[$idcampo], 2, '.','');
					}
					else if($typecampo=="datetime")
					{
						if($rsdate[$idcampo])
						{
							$valdato=date("Y-m-d h:i:s A",strtotime($rsdate[$idcampo]));
						}
					}
					else if($typecampo=="date")
					{
						if($rsdate[$idcampo])
						{
							$valdato=$rsdate[$idcampo];//date("Y-m-d",strtotime($rsdate[$idcampo]));
						}
					}
					else if(strpos($idcampo,'_')!==false && strpos($idcampo,'_')==0)//si es Consulta externa
					{   //$idcampo[0]=='_'
						$valdato=$rsdate[$idcampo];
						if(isset($this->arrayRecordExterno[$idcampo][$rsdate[$idcampo]]))
						{
							if($this->arrayRecordExterno[$idcampo][$rsdate[$idcampo]])
								$valdato=$this->arrayRecordExterno[$idcampo][$rsdate[$idcampo]];
						}
						else
							if(isset($this->arrayRecordExterno[$rsdate[$idcampo]]))
							{
								if($this->arrayRecordExterno[$rsdate[$idcampo]])
									$valdato=$this->arrayRecordExterno[$rsdate[$idcampo]];
							}
					}
					else if($idcampo=="#")
					{
						$valdato=$this->row;
					}
					else
						$valdato=$rsdate[$idcampo];
					$this->getActiveSheet()->setCellValueByColumnAndRow($col-1,$this->rownum,utf8_encode($valdato));
					//$this->getActiveSheet()->getColumnDimension($this->GetColumn($col))->setAutoSize(true);
					$col++;
				}
				///color font
				
				$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)->getFont()
				->setName('Arial')
				->setSize(10)
				->getColor()->setRGB(($fontcolor)?$fontcolor:'000000');
				///color celda
				$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)
				->getFill()
				->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
				->getStartColor()
				->setRGB(($rowcolor)?$rowcolor:((($this->rownum%2)==0)?'E8E8E8':'FFFFFF'));
				///Border
				$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN); 
				$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum.':'.$this->GetColumn($this->numcampos).$this->rownum)->getBorders()->getAllBorders()->getColor()->setARGB('000000') ; // Color del marco 
				
				$this->rownum++;
				$this->row++;
			}
			else
			{
				$this->getActiveSheet()->setCellValueByColumnAndRow(0,$this->rownum,"No Se han definidos los Campos");
			
				$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum)->getFont()
				->setName('Arial')
				->setSize(10)
				->getColor()->setRGB('FE0202');
				$this->getActiveSheet()->getStyle(($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos)).$this->rownum)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				
				$this->rownum++;
			}
		}
		else
		{
			$this->getActiveSheet()->setCellValueByColumnAndRow(0,$this->rownum,"No contiene informacion");
			
			$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum)->getFont()
			->setName('Arial')
			->setSize(10)
			->getColor()->setRGB('FE0202');
			$this->getActiveSheet()->getStyle(($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos)).$this->rownum)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			
			$this->rownum++;
		}
	}
	private function AddTbodyExcel($sql,$dB)
	{
		@odbc_free_result($record);
		$record=$dB->ExecSQL($sql,$msgerror);
		if(!$record)
		{
			$this->AddMsg($msgerror,0);
		}
		else
		{
			if(odbc_num_rows($record))
			{	
				if(!count($this->campos))
				{
					$this->AddColumn($dB->getFieldsNameResult($record,$msgerror));	
				}
				while($reg=odbc_fetch_array($record))
				{
					$this->AddRowtbodyExcel($reg);
				}
			}
			else
			{
				$this->AddMsg($this->ArrayMsg[2],2);
				//$this->AddMsg('Sin registros',3);
			}
			
		}
	}
	public function getarrayRecordExterno()
	{
		return $this->arrayRecordExterno;
	}
	/********************************
	
	*********************************/
	public function ExcelReportGenerate($sql,$arrRegExt=0)
	{
		$this->arrayRecordExterno=$arrRegExt;
		$this->propertiesExcel();
		if(!count($this->configDataBase))
		{
			$this->AddMsg("No se ha definido la configuracion para el recurso de datos",0);
		}
		else
		{
			$msgerror="";
			$dB=new ODBC($this->configDataBase["serverDB"],$this->configDataBase["nameDB"],$this->configDataBase["userDB"],$this->configDataBase["passwordDB"]);
			if(!is_array($sql))
			{
				$this->AddTbodyExcel($sql,$dB);
			}
			else
			{
				foreach($sql as $key=>$scriptsql)
				{
					if(!is_array($scriptsql))//si no es arreglo
					{
						if($this->campos)
						{
							$this->rownum+=2;
							$this->row=1;
							unset($this->campos);
						}
						$this->AddTbodyExcel($scriptsql,$dB);
					}
					else//si es areglo
					{
						foreach($scriptsql as $keys=>$arraysqls)
						{
							if($this->campos)
							{
								$this->rownum+=2;
								$this->row=1;
								unset($this->campos);
							}
							$this->AddTbodyExcel($datsqls["sql"],$dB);
						}
					}
				}
			}
		}
		$this-> ExcelGenerate();
	}
	/**************************************************
	
	***************************************************/
	private function ExcelGenerate()
	{
		//si no tiene campos o tiene mensaje
		if(!count($this->campos)||($this->msg))
		{
			$this->getActiveSheet()->setCellValueByColumnAndRow(0,$this->rownum,$this->msg);
			$this->getActiveSheet()->getColumnDimension($this->GetColumn(1))->setAutoSize(true);
			$this->getActiveSheet()->getStyle($this->GetColumn(1).$this->rownum)->getFont()
			->setBold(true)
			->setName('Arial')
			->setSize(10)
			->getColor()->setRGB(($this->tipomsg==0)?'FE0202':'0000FF');
			$this->getActiveSheet()->getStyle(($this->GetColumn(1).$this->rownum .':'.$this->GetColumn($this->numcampos)).$this->rownum)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			
			$this->rownum++;
		}
		// Se modifican los encabezados del HTTP para indicar que se envia un archivo de Excel.
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="'.$this->filename.'"');
		header('Cache-Control: max-age=0');
		$objWriter = PHPExcel_IOFactory::createWriter($this, 'Excel2007');
		$objWriter->save('php://output');
		//exit();
	}
	
}
?>