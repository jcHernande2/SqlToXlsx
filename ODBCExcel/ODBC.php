<?php
/*
 * @autor       Juan Carlos Hdz Hdz
 * @category    PHPExcel
 * @version     1.0, 2016-05-02
 */
class ODBC
{
	private $server_db="";
	private $name_db="";
	private $user_db="";
	private $password_db="";
	
	public function __construct($server_db="",$name_db="",$user_db="",$password_db="")
	{
		$this->server_db=$server_db;
		$this->name_db=$name_db;
		$this->user_db=$user_db;
		$this->password_db=$password_db;
	}
	private function DBConnect(&$msgerror)
	{
		$strc ="Driver={SQL Server Native Client 10.0};Server=".$this->server_db.";Database=".$this->name_db.";";
		$c = odbc_connect($strc, $this->user_db,$this->password_db);
		if(!$c)
		{ 
			$msgerror="Error al intentar realizar la Conexion. ";
			return FALSE;
		}
		return $c;
	}
	public function tableNameDB(&$msgerror)
	{
		$c=$this->DBConnect($msgerror);
		$result = odbc_tables($c);
		$tables = array();
		while (odbc_fetch_row($result))
		{
			if(odbc_result($result,"TABLE_TYPE")=="TABLE")
				$tables[]=odbc_result($result,"TABLE_NAME");
		}
		return $tables;
	}
	public function fieldTDB($tName,&$msgerror)
	{
		$c=$this->DBConnect($msgerror);
		$outval = odbc_columns($c, $this->name_db, "%", $tName, "%");
		$pages = array();
		$fieldName= array();
		while (odbc_fetch_into($outval, $pages)) 
		{
			$fieldName=$pages[3];
		}
		return $fieldName;
	}
	public function getFieldsNameResult($record,&$msgerror)
	{
		$arrayFields=array();
		for($i = 1;$i <= odbc_num_fields($record);$i++)
		{
			$arrayFields[odbc_field_name($record,$i)]=odbc_field_type ($record,$i);
		}
		return $arrayFields;
	}
	public function ExecSQL($sql,&$msgerror)
	{
		$c=$this->DBConnect($msgerror);
		if($c)
		{
			$recordset = odbc_exec($c,$sql);
			if (!$recordset) 
			{
				$msgerror="Error al realizar la consulta. ";
				return FALSE;
			}
			else
			{
				return $recordset;
			}
		}
		return FALSE;
	}
}
?>