<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package	PHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Style_ColorScale
 *
 * @category   PHPExcel
 * @package	PHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @author	Bas Vijfwinkel
 */
class PHPExcel_Style_ColorScale extends PHPExcel_Style_Supervisor implements PHPExcel_IComparable
{

	/**
	* CFVO Array
	*
	* @var array
	*/
	
	protected $_cfvos;
	
	/**
	* colors
	*
	* @var array of PHPExcel_Style_Color
	*/
	
	protected $_colors;
	
	/**
	* uniq id for this object
	*
	* @var string 
	*/
	protected $_id;
	
	/**
	 * Create a new PHPExcel_Style_Border
	 *
	 */
	public function __construct() 
	{
		// set these cfvo values by default because without them no colorscale is shown
		$this->_cfvos = array(PHPExcel_Style_CFVOType::fromString('min'),PHPExcel_Style_CFVOType::fromString('max'));
		// uniq id
		$this->_id = uniqid('',true);
	}	
	
	/**
	 * Apply styles from array
	 *
	 * <code>
	 * $pValue = new PHPExcel_Style_ColorScale();
	 *	$pValue->applyFromArray(
	 *		array(
	 *			'cellgroup' => 'A1:A5',
	 *			'colors' => array(array('rgb'=>'00FF0000','rgb'=>'FFFF0000','rgb'=>'00FFFF00')),
	 *          'cfvos' => array(array('type'=>'min'),array('type'=>'max'))
	 *			), true);
	 * $objPHPExcel->getActiveSheet()->setConditionalStyles('A3',$pValue)
	 * </code>
	 *
	 * @param	array	$pStyles	Array containing style information
	 * @return	PHPExcel_Style_ColorScale
	 * Note : cfvos and color parameter will override all existing cfvo and color values; If you want to preserve them, add them manually with addCfvo or addColor
	 */
	public function applyFromArray($pStyles = null, $checkInput=true, $isExtLstData= false) {
		if (is_array($pStyles)) 
		{
			if (array_key_exists('cellgroup', $pStyles))    
			{ 
				$this->setCellGroup($pStyles['cellgroup']); 
			}
			if (array_key_exists('colors', $pStyles))
			{ 
				foreach($pStyles['colors'] as $color)
				{
					$this->setColor((new PHPExcel_Style_Color())->applyFromArray($color)); 
				}
			}
			if (array_key_exists('cfvos', $pStyles)) 
			{
				$resultcfvos = array();
				foreach ($pStyles['cfvos'] as $cfvotype)
				{
					$resultcfvos[] = PHPExcel_Style_CFVOType::fromArray($cfvotype);
				}
				$this->addCfvos($resultcfvos); // add cfvo to ext_cfvo to preserve both
			}
		}
		else
		{
			throw new PHPExcel_Exception("ColorScale : invalid input applyFromArray :".var_export($pStyles,true));
		}
		return $this;
	}
	
	
	/*
	* Create an array of elements containing all the information of the input xml object
	*
	* @param	string	$ref	cell reference (e.g A1:A5)
	* @param	SimpleXML $cfRule	cfRule xml structure containing the colorscale section
	* @return	Array	array containing all the information of the input xml object 
	* @throws PHPExcel_Exception
	*/
	public function applyFromXML($ref, $cfRule)
	{
		// these default properties must exist
		if (isset($cfRule->colorScale->color) &&
			isset($cfRule->colorScale->cfvo))
		{
			
			// add default properties
			$this->setCellGroup($ref);
			$this->_colors = array(); // clear out colors bfore adding new ones
			foreach($cfRule->colorScale->color as $color)
			{
				$this->setColor(new PHPExcel_Style_Color((string)$color['rgb']));
			}
			$this->_cfvos = array(); // clear our the list before adding new ones
			foreach($cfRule->colorScale->cfvo as $cfvo)
			{
				$this->addCfvo(PHPExcel_Style_CFVOType::fromXML($cfvo));
			}
		}
		else
		{
			// missing property
			throw new PHPExcel_Exception("ColorScale : missing color or cvfo setting");
		}
		
		return $this;
	}
	
		
	/*
	/*
	* Create an array of elements containing all the information of the object
	*  in order to pass it on to the xmlwriter
	*
	* @return	Array	array containing all the information of this colorscale
	* NOTE : The order in which the elements are written into the array should not matter but for Microsoft Excel it apparently _DOES MATTER_
	*/
	public function getElementsAsArray()
	{
		$result = array();
		
		// 1. cfvo's
		if ($this->_cfvos)
		{
			foreach($this->_cfvos as $cfvotype)
			{
				$result[] = $cfvotype->toArray();
			}
		}
		// 2. colors
		if ($this->_colors)  
		{ 
			foreach($this->_colors as $color)
			{
				$result[] = array('name' => 'color', 'attributes' => array(array('name' => 'rgb', 'attributes' => $color->getARGB()))); 
			}
		}
		
		// return the resulting array
		return $result;
	}
	
	/*
	 * get the color for the colorScale
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) { $color = $worksheetstyles[0]->getColor(1) ; }
	 * </code>
	 *
	 * @params	int	index of the requested color
	 * @return PHPExcel_Style_Color
	 * @throws	PHPExcel_Exception	if index not found
	*/
	public function getColor($index) 
	{
		if ($this->_colors && isset($this->_colors[$index]))
		{
			return $this->_colors[$index];
		}
		else
		{
			throw new PHPExcel_Exception("ColorScale : no colors or index of requested color does not exist");
		}
    }
	
	/*	
	 * Set the fill color of the colorscale
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) 
	 * { 
	 * 	 $worksheetstyles[0]->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_RED,2)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	fill color of the colorscale
	*/
	public function setColor($color,$index = null) 
	{
		if (is_null($this->_colors)) { $this->_colors = array(); }
		if (is_null($index)) { $index = count($this->_colors); }
   		$this->_colors[$index] = $color;
   		return $this;
    }
	
	/*
	 * Set all colors
	 *
	 * @param	array	array of PHPExcel_Color objects
	 * @return PHPExcel_Style_ColorScale
	 */
	public function setColors($colors)
	{
		$this->_colors = $colors;
		return $this;
	}
	
	/*
	 * Get all colors
	 *
	 * @returns	array	array of PHPExcel_Color objects
	 */
	public function getColors()
	{
		return $this->_colors;
	}

	/*
	 * get the group of cells that this colorscale setting applies to
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) { $color = $worksheetstyles[0]->getCellGroup() ; }
	 * </code>
	 *
	 * @return string
	*/
    public function getCellGroup() 
	{
    	return $this->_cell_group;
    }
	
	/*
	 * set the group of cells that this colorscale setting applies to
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) 
	 * { 
	 *	 $worksheetstyles[0]->setCellGroup('A1:A5') ; 
	 * }
	 * </code>
	 *
	 * @return string
	*/
	public function setCellGroup($cell_group = null) {
   		$this->_cell_group = $cell_group;
   		return $this;
    }

	/*
	 * Get the cfvo settings
	 * 
	 * @params	boolean	set to true if the extlst cfvo settings must be used
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) { $cfvo = $worksheetstyles[0]->getCfvo() ; }
	 * </code>
	 *
	 * @return array of CFVOType	all CFVOTypes for this colorscale	
	*/
    public function getCfvos($use_extlst_cfvos=false) 
	{
		if ($use_extlst_cfvos)
		{
			return $this->_extlst_cfvos;
		}
		else
		{
			return $this->_cfvos;
		}
    }	
	
	/*
	* add a cfvo type
	*
	* @param	cfvotype
	* @param	boolean	set to true is the data must be entered to the extlst cvfo array
	* <code>
	* $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	* if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) 
	* { 
	*	$worksheetstyles[0]->addCfvoType(PHPExcel_Style_CFVOType::fromString('min'));
	*   $worksheetstyles[0]->addCfvoType(PHPExcel_Style_CFVOType::fromString('max'));
	* }
	* </code>
	*/
	public function addCfvo($cfvotype = null, $use_extlst_cfvos=false) 
	{
		if ($cfvotype)
		{
			if (!$use_extlst_cfvos)
			{
				array_push($this->_cfvos,$cfvotype);
			}
			else
			{
				if (!$this->_extlst_cfvos) { $this->_extlst_cfvos = array();}
				array_push($this->_extlst_cfvos,$cfvotype);
			}
		}
   		return $this;
    }
	/*
	 * Add a list of cfvos (Note: existing entries will be destroyed
	 *
	 * @param	array	list of cfvotypes
	 * @param	boolean	set to true is the data must be entered to the extlst cvfo array
	 *
	 *
	 */
	public function addCfvos($cfvo = null, $use_extlst_cfvos=false) 
	{
		if ($use_extlst_cfvos)
		{
			$this->_extlst_cfvos = $cfvo;
		}
		else
		{
			$this->_cfvos = $cfvo;
		}
   		return $this;
    }
	
	
	/*
	 * Get the hashcode for this object
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * $hashcode = $worksheetstyles[0]->getHashCode(); 
	 * </code>
	 *
	 * @return string containing md5 hashcode of the object
	*/
	public function getHashCode()
	{
		return strtoupper(md5($this->_id));
	}
	
	/*
	 * create a datastructure for creating the colorscale element with default properties
	 *
	 * @param	array	array with properties 
	 */
	public function getDefaultData()
	{
		// add cfvo
		$cfvos = $this->getCfvos();
		$result = array('name' => 'colorScale', 'attributes' => array());
		foreach ($cfvos as $cfvo)
		{
				$result['attributes'][] = $cfvo->toArray();
		}
		// add colors
		$colors = $this->getColors();
		foreach ($colors as $color)
		{
			$result['attributes'][] =  array('name' => 'color', 'attributes' => array(array('name' => 'rgb', 'attributes' => $color->getARGB())));
		}
		return $result;
	}
	
	/*
	 *  Indicate whether this object needs a extlst data entry
	 *  
	 * @return	bool	always false : ColorScale has no extlst data
	 */
	public function needsExtLstEntry()
	{
		return false;
	}

}
