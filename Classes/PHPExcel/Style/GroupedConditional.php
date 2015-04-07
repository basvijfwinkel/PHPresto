<?PHP

class PHPExcel_Style_GroupedConditional extends PHPExcel_Style_Supervisor
{
	/**
	* cellReference definition : (eg:A1:C5)
	*
	* @var string
	*/	
	protected $_cellReference;

	
	public function __construct()
	{
		// set these cfvo values by default because without them no IconSet is shown
		$this->_cfvos = array(PHPExcel_Style_CFVOType::fromString('min'),PHPExcel_Style_CFVOType::fromString('max'));
		// uniq id
		$this->_id = uniqid('',true);
	}
	
	
	public function removeCellReference($position)
	{
		if ((is_array($this->_cellReference)) && (in_array($position, $this->_cellReference)))
		{
			$index = array_search($position, $this->_cellReference);
			if ($index !== FALSE ) { unset($this->_cellReference[$index]); }
		}
		elseif ($this->_cellReference == $position)
		{
			$this->_cellReference = null;
		}
	}

	public function updateCellReference($oldreference, $newreference)
	{
		if ((is_array($this->_cellReference)) && (in_array($oldreference, $this->_cellReference)))
		{
			$index = array_search($oldreference, $this->_cellReference);
			if($index !== FALSE) 
			{ 
				$this->_cellReference[$index] = $newreference; 
			}
		}
		elseif ($this->_cellReference == $oldreference)
		{
			$this->_cellReference = $newreference;
		}
	}
	
		/*
	 * set the group of cells that this colorscale setting applies to
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) 
	 * { 
	 *	 $worksheetstyles[0]->setCellReference('A1:A5') ; 
	 * }
	 * </code>
	 *
	 * @return string
	*/
	public function setCellReference($cellReference = null) 
	{
		$this->_cellReference = array();
		list($startColumn, $endColumn, $startRow, $endRow) = $this->_extractRowColumns($cellReference);
		for($row = $startRow;$row <= $endRow;$row++)
		{
			for($column = $startColumn; $column <= $endColumn; $column++)
			{
				$position = PHPExcel_Cell::stringFromColumnIndex($column-1).$row;
				array_push($this->_cellReference, $position);
			}
		}
   		return $this;
    }
	
	protected function _extractRowColumns($cellReference)
	{
		$parts = explode(':',$cellReference);
		// start
		list($startColumChar, $startRow) = PHPExcel_Cell::coordinateFromString($parts[0]);
		$startColumn = PHPExcel_Cell::columnIndexFromString($startColumChar);
		if (count($parts) > 1)
		{
			// range
			list($endColumChar, $endRow) = PHPExcel_Cell::coordinateFromString($parts[1]);
			$endColumn = PHPExcel_Cell::columnIndexFromString($endColumChar);
		}
		else
		{
			// single reference
			$endColumn = $startColumn;
			$endRow = $startRow;
		}
		
		return array($startColumn, $endColumn, $startRow, $endRow);
	}
	

	/*
	 * get the group of cells that this IconSet setting applies to
	 * NOTE : THE RANGE IS DETERMINED ON THE MIX/MAX ROW/COLUMNS AND PRESUMES THE CELLS IN BETWEEN
	 *       ARE ALSO SUBJECT TO THIS CONDITIONAL FORMATTING
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_IconSet) { $color = $worksheetstyles[0]->getCellReference() ; }
	 * </code>
	 *
	 * @return string
	*/
    public function getCellReference() 
	{
		$minColumn = null; $maxColumn = null; $minRow = null; $maxRow = null;
		foreach($this->_cellReference as $cellRef)
		{
			list($columnChar, $row) = PHPExcel_Cell::coordinateFromString($cellRef);
			$column = PHPExcel_Cell::columnIndexFromString($columnChar);
			$minColumn = (($minColumn == null)||($column < $minColumn))?$column:$minColumn;
			$maxColumn = (($maxColumn == null)||($column > $maxColumn))?$column:$maxColumn;
			$minRow = (($minRow == null)||($row < $minRow))?$row:$minRow;
			$maxRow = (($maxRow == null)||($row > $maxRow))?$row:$maxRow;
		}
		
		if (($minRow == $maxRow) && ($minColumn == $maxColumn))
		{
			// single reference
			$result = PHPExcel_Cell::stringFromColumnIndex($minColumn-1).$minRow;
		}
		else
		{
			// range
			$result = PHPExcel_Cell::stringFromColumnIndex($minColumn-1).$minRow.':'.PHPExcel_Cell::stringFromColumnIndex($maxColumn-1).$maxRow;
		}
		return $result;
    }

	
}
?>