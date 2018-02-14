<?php
/*
TODO : 

- check : if priority of 2 databars work correctly
- consider : preserve CLASSID?

/*<code>
       $conditional = new PHPExcel_Style_Conditional();
       $conditionType = PHPExcel_Style_Conditional::CONDITION_DATABAR;
       $settingsArray =  array(
                          'cellReference' => 'A1:A5',
                          'color' => ['rgb' => '557DBC'],
                          //'fillColor' => ['rgb' =>'00FF00'],
                          //'border' => true,
                          //'borderColor' => ['rgb' =>'5279BA'],
                          //'negativeBarColorSameAsPositive' => 0,
                          //'negativeFillColor' => ['rgb' =>'FF0000'],
                          //'negativeBarBorderColorSameAsPositive' => 0,
                          //'negativeBorderColor' => ['rgb' =>'FF00FF'],
                          //'axisColor' => ['rgb' =>'7F7F7F'],
                          //'minLength' => 20,
                          //'maxLength' => 70,
                          //'showValue' => true,
                          //'direction' => 'context',
                          //'axisPosition' => 'middle',
                          //'cfvos' => array(array('type'=>'min'),array('type'=>'max'))
                      );
       $conditional->setConditionObjectFromArray($conditionType, $settingsArray);
       $phpExcelObj->getActiveSheet->getStyle('A1')->addConditionalStyle($conditional);
</code>
*/

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
 * PHPExcel_Style_DataBar
 *
 * @category   PHPExcel
 * @package	PHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @author	Bas Vijfwinkel
 */
class PHPExcel_Style_DataBar extends PHPExcel_Style_GroupedConditional implements PHPExcel_IComparable
{

	/**
	* CFVO Array for Databar
	*
	* @var array
	*/
	
	protected $_cfvos;
	
	/**
	* ExtLstCFVO Array for Databar (cfvo data in the default databar block and the extlst data block might differ)
	*
	* @var array
	*/
	
	protected $_extlst_cfvos;
	
	/**
	* color for Databar
	*
	* @var PHPExcel_Style_Color
	*/
	
	protected $_color;
	
	
	/**
	* fill color
	*
	* @var PHPExcel_Style_Color
	*/
	protected $_fillColor;
	
	/**
	* border color
	*
	* @var PHPExcel_Style_Color
	*/
	protected $_borderColor;

	/**
	* 	negative fill color
	*
	* @var PHPExcel_Style_Color
	*/
	protected $_negativeFillColor;
	
	/**
	* negative border color
	*
	* @var PHPExcel_Style_Color
	*/
	protected $_negativeBorderColor;
	
	/**
	* axis color
	*
	* @var PHPExcel_Style_Color
	*/
	protected $_axisColor;
	
	/**
	* min langth
	*
	* @var unsigned int (default : 10)
	*/
	protected $_minLength;
	
	/**
	* max length
	*
	* @var unsigned int (default : 90)
	*/
	protected $_maxLength;
	
	/**
	* show value in the cell (0=false;1=true)
	*
	* @var Integer (default : 1)
	*/
	protected $_showValue;

	/**
	* show border (Note : 0=false; 1=true)
	*
	* @var Integer	(default : 0)
	*/
	protected $_border;
	
	/**
	* show gradient (Note : 0=false; 1=true)
	*
	* @var Integer (default : 0)
	*/
	protected $_gradient;
	
	/**
	* direction of the databar 
	*
	* @var PHPExcel_Style_DataBarDirection (default 'context')
	*/
	protected $_direction;
	
	/**
	* negativeBarColorSameAsPositive  (0=false;1=true)
	*
	* @var Integer (default : 0)
	*/
	protected $_negativeBarColorSameAsPositive;
	
	/**
	* negativeBarBorderColorSameAsPositive
	*
	* @var Boolean (default : true)
	*/
	protected $_negativeBarBorderColorSameAsPositive;
	
	/**
	* axisPosition
	*
	* @var PHPExcel_Style_DataBarAxisPosition (default : automatic)
	*/
	protected $_axisPosition;
	
	/**
	* namespace for extLst entry (optional)
	*
	* @var string (default : http://schemas.microsoft.com/office/spreadsheetml/2009/9/main)
	*/
	protected $_namespace;
	
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
		parent::__construct();
		// default namespace
		$this->_namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
	}
	
	/**
	 * Check if the object needs a extLst entry 
	 * if besides color and cfvo any other property is used, Excel2010 only needs to add a conditional format entry in the extLst structure.
	 *
	 *  @return : boolean	true if an extLst entry needs to be generated. 
	 */
	 public function needsExtLstEntry()
	 {
		return ($this->_extlst_cfvos ||
				$this->_fillColor || 
				$this->_borderColor ||
				$this->_negativeFillColor ||
				$this->_negativeBorderColor ||
				$this->_axisColor ||
				!is_null($this->_minLength) ||
				!is_null($this->_maxLength) ||
				//!is_null($this->_showValue) ||
				!is_null($this->_border) ||
				!is_null($this->_gradient) ||
				$this->_direction ||
				!is_null($this->_negativeBarColorSameAsPositive) ||
				!is_null($this->_negativeBarBorderColorSameAsPositive) ||
				$this->_axisPosition)?true:false;
	 }

	/**
	 * Apply styles from array
	 *
	 * @param	array	$pStyles	Array containing style information
	 * @param   boolean	$checkInput	set to true if the validity of the data must be checked (Excel2010 does not seem to follow the specifications...)
	 * @param	boolean	$isExtLstData	set to true if the data is from an extlst block (cfvo data will be stored separately)
	 * @return PHPExcel_Style_DataBar
	 * Note : cfvos parameter will override all existing cfvo values; If you want to preserve them, add them manually with addCfvo
	 */
	public function applyFromArray($pStyles = null, $checkInput=true, $isExtLstData= false) {
		if (is_array($pStyles)) 
		{
			if (array_key_exists('cellReference', $pStyles))       { $this->setCellReference($pStyles['cellReference']); }
			if (array_key_exists('color', $pStyles))               { $this->setColor(new PHPExcel_Style_Color($pStyles['color']['rgb'])); }
			if (array_key_exists('fillColor', $pStyles))           { $this->setFillColor(new PHPExcel_Style_Color($pStyles['fillColor']['rgb'])); }	
			if (array_key_exists('borderColor', $pStyles))         { $this->setBorderColor(new PHPExcel_Style_Color($pStyles['borderColor']['rgb'])); }
			if (array_key_exists('negativeFillColor', $pStyles))   { $this->setNegativeFillColor(new PHPExcel_Style_Color($pStyles['negativeFillColor']['rgb'])); }
			if (array_key_exists('negativeBorderColor', $pStyles)) { $this->setNegativeBorderColor(new PHPExcel_Style_Color($pStyles['negativeBorderColor']['rgb'])); }
			if (array_key_exists('minLength', $pStyles))           { $this->setMinLength((int)$pStyles['minLength']); }
			if (array_key_exists('maxLength', $pStyles))           { $this->setMaxLength((int)$pStyles['maxLength']); }
			if ($checkInput && (array_key_exists('minLength', $pStyles)) && (array_key_exists('maxLength', $pStyles)) && ($pStyles['minLength'] > $pStyles['maxLength']))
			{
				throw new PHPExcel_Exception("DataBar : minLength should be smaller or equal to maxLength");
			}
			if (array_key_exists('showValue', $pStyles))           { $this->setShowValue((int)$pStyles['showValue']); }
			if (array_key_exists('border', $pStyles))              { $this->setBorder((int)$pStyles['border']); }
			if (array_key_exists('gradient', $pStyles))            { $this->setGradient((int)$pStyles['gradient']); }
			if (array_key_exists('negativeBarColorSameAsPositive', $pStyles)) { $this->setNegativeBarColorSameAsPositive((int)$pStyles['negativeBarColorSameAsPositive']); }
			if ($checkInput && (array_key_exists('negativeBarColorSameAsPositive', $pStyles)) && (!(array_key_exists('negativeFillColor', $pStyles))))
			{
				throw new PHPExcel_Exception("DataBar : negativeFillColor should be set");
			}
			if (array_key_exists('negativeBarBorderColorSameAsPositive', $pStyles)) { $this->setNegativeBarBorderColorSameAsPositive((int)$pStyles['negativeBarBorderColorSameAsPositive']); }
			if ($checkInput && (array_key_exists('negativeBarBorderColorSameAsPositive', $pStyles)) && (!(array_key_exists('negativeBorderColor', $pStyles))))
			{
				throw new PHPExcel_Exception("DataBar : negativeBorderColor should be set");
			}
			if (array_key_exists('direction', $pStyles))          { $this->setDirection(PHPExcel_Style_DataBar_DataBarDirection::fromString($pStyles['direction'])); }
			if (array_key_exists('axisPosition', $pStyles))       { $this->setAxisPosition(PHPExcel_Style_DataBar_DataBarAxisPosition::fromString($pStyles['axisPosition'])); }
			if (array_key_exists('cfvos', $pStyles)) 
			{
				$resultcfvos = array();
				foreach ($pStyles['cfvos'] as $cfvotype)
				{
					$resultcfvos[] = PHPExcel_Style_CFVOType::fromArray($cfvotype);
				}
				$this->addCfvos($resultcfvos,$isExtLstData); // add cfvo to ext_cfvo to preserve both
			}
			
			if (array_key_exists('axisColor', $pStyles)) 
			{ 
				if ((!$checkInput) || ((array_key_exists('axisPosition', $pStyles))	&& ($this->_axisPosition->toString() != PHPExcel_Style_DataBar_DataBarAxisPosition::NONE)))
				{
					$this->setAxisColor(new PHPExcel_Style_Color($pStyles['axisColor']['rgb'])); 
				}
				else
				{
					throw new PHPExcel_Exception("DataBar : in order to set axisColor, axisPosition must be defined and not set to 'NONE'");
				}
			}			
		}
		else
		{
			throw new PHPExcel_Exception("DataBar : invalid input applyFromArray :".var_export($pStyles,true));
		}
		return $this;
	}
	
	/*
	* Create an array of elements containing all the information of the input xml object
	*
	* @param	string	$ref	cell reference (e.g A1:A5)
	* @param	SimpleXML $cfRule	cfRule xml structure containing the databar section
	* @param	SimpleXML	$extLst	extLst structure of the worksheet the databar object is defined for
	* @return	Array	array containing all the information of the input xml object 
	* @throws PHPExcel_Exception
	*/
	public function applyFromXML($ref, $cfRule, $extLst)
	{
		// these default properties must exist
		if (isset($cfRule->dataBar->color[0]['rgb']) &&
			isset($cfRule->dataBar->cfvo[0]['type']) &&
			isset($cfRule->dataBar->cfvo[1]['type'])
			)		
		{
			// add default properties (if they are defined)
			$this->setCellReference($ref);
			$this->setColor(new PHPExcel_Style_Color((string)$cfRule->dataBar->color[0]['rgb']));
			$this->_cfvos = array(); // clear our the list before adding new ones
			$this->addCfvo(PHPExcel_Style_CFVOType::fromXML($cfRule->dataBar->cfvo[0]));
			$this->addCfvo(PHPExcel_Style_CFVOType::fromXML($cfRule->dataBar->cfvo[1]));
			if (isset($cfRule->dataBar['showValue']))
			{
				$this->setShowValue((int)$cfRule->dataBar['showValue']) ;
			}
			
			// check if an extLst object is used to mark up this databar
			if (isset($cfRule->extLst) &&
			    isset($cfRule->extLst->ext['uri']) &&
				($cfRule->extLst->ext['uri'] == "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}") // ID for ext uri : http://msdn.microsoft.com/en-us/library/dd905242%28v=office.12%29.aspx
				)
			{
				// extract all data
				foreach ($cfRule->extLst->ext[0]->getNamespaces(true) as $ns_name => $ns_uri)
				{
					if ($ns_name)
					{
						// save this namespace
						$this->setNamespace($ns_uri);
						// look up the id of this extLst entry
						$children = $cfRule->extLst->ext[0]->children($ns_name,TRUE);
						$linkid = (string)$children->id;
						// search the data in the extLst structure
						if ($extLst)
						{
							// find the object that matches the linkid
							$linkeddatabarelements = $extLst->xpath('//*[@id="'.$linkid.'"]');
							if ($linkeddatabarelements)
							{
								// create an array with all the relevant information and 
								// add it to the databar object
								$linkeddatabarelement = $linkeddatabarelements[0]->children($ns_name,TRUE);
								$databar = $linkeddatabarelement->dataBar;
								$databar_array = $this->xml2array($databar,$ns_name);
								// apply all the setting from the array
								$isExtLstData = true;
								$this->applyFromArray($databar_array, false, $isExtLstData);
							}
							else
							{
								// missing databar element with ID
								throw new PHPExcel_Exception("DataBar : missing databar element entry with id ".$linkid." in extLst");
							}
						}
						else
						{
							throw new PHPExcel_Exception("DataBar : missing extLst entry");
						}
					}
				}
				
			}
		}
		else
		{
			// missing property
			throw new PHPExcel_Exception("DataBar : missing color or cvfo setting");
		}
		
		return $this;
	}
	
	/*
	* Create an array of elements containing all the information of the input xml object
	*
	* @param	SimpleXML	xml object to convert to an array
	* @param	String	namespace of the elements
	* @return	Array	array containing all the information of the input xml object 
	*/
	protected function xml2array($inputxml, $namespace=NULL)
	{
		$result = array();
		// add attributes
		foreach ($inputxml->attributes() as $attr_name => $attr_value)
		{
			$result[$attr_name] = (string)$attr_value;
		}
		// add child objects 
		if ($namespace)
		{
			$children = $inputxml->children($namespace, TRUE);
		}
		else
		{
			$children = $inputxml->children();
		}
		
		foreach ($children as $prop_name => $prop)
		{
			$prop_array = array();
			foreach ($prop->attributes() as $attr_name => $attr_value)
			{
				$prop_array[$attr_name] = (string)$attr_value;
			}
			
			if ($prop_name != 'cfvo')
			{
				$result[$prop_name] = $prop_array;
			}
			else
			{
				if (!isset($result['cfvos'])) { $result['cfvos'] = array(); }
				
				// there might be a child element with the value
				$cfvo_formula = $prop->children('xm', TRUE);
				if ($cfvo_formula)
				{
					$prop_array['xm:f'] = (int)$cfvo_formula->f;
				}
				$result['cfvos'][] = $prop_array;
			}
		}		
		return $result;
	}
	
	/*
	* Create an array of elements containing all the information of the object
	*  in order to pass it on to the xmlwriter
	*
	* @return	Array	array containing all the information of this databar
	* NOTE : The order in which the elements are written into the array should not matter but for Microsoft Excel it apparently _DOES MATTER_
	*/
	public function getElementsAsArray($forExtLst=false)
	{
		// 1. cfvo's (do not add to the extlst ouput)
		$result = array();
		if (($this->_cfvos) && (!$forExtLst))
		{
			foreach($this->_cfvos as $cfvotype)
			{
				$result[] = $cfvotype->toArray($forExtLst);
			}
		}
		// 2. extlst_cfvo's (add to the extlst ouput with name 'cfvo'; else 'extlst_cfvo')
		if ($this->_extlst_cfvos)
		{
			foreach($this->_extlst_cfvos as $cfvotype)
			{				
				$result[] = $cfvotype->toArray($forExtLst,(($forExtLst)?'cfvo':'extlst_cfvo'));
			}
		}
		// 3. color (do not add to the extlst ouput)
		if (($this->_color) && (!$forExtLst)) { $result[] = array('name' => 'color', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->_color->getARGB()))); }
		// 4. fillColor
		if ($this->_fillColor)          { $result[] = array('name' => 'fillColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->_fillColor->getARGB()))); }
		// 5. borderColor (only if border = true)
		if ((!is_null($this->_border)) && ($this->_borderColor))
		{
			$result[] = array('name' => 'borderColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->_borderColor->getARGB()))); 
		}
		// 6. negativeFillColor (only if negativeBarColorSameAsPositive = true)
		//if (($this->_negativeBarColorSameAsPositive === 0) && ($this->_negativeFillColor)) // Excel ignores specification : negativeBarColorSameAsPositive not needed ?
		if ($this->_negativeFillColor)
		{
			$result[] = array('name' => 'negativeFillColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->_negativeFillColor->getARGB()))); 
		}
		// 7. negativeBorderColor (only if negativeBarBorderColorSameAsPositive and border are true)
		if (($this->_negativeBarBorderColorSameAsPositive === 0) && ($this->_negativeBorderColor) && (!is_null($this->_border)))
		{
			$result[] = array('name' => 'negativeBorderColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->_negativeBorderColor->getARGB()))); 
		}
		// 8. axisColor ( only if axis position is not NONE)
		//if ($this->_axisPosition)
		//{
			//$axisPosition = $this->_axisPosition->toString();
			//if (($axisPosition != PHPExcel_Style_DataBar_DataBarAxisPosition::NONE) && ($this->_axisColor))
			if ($this->_axisColor) // it seems that excel does not follow the definition here and defines an axisColor without an axisPosition
			{
				$result[] = array('name' => 'axisColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->_axisColor->getARGB())));
			}
		//}
		// 9. minLength
		if (!is_null($this->_minLength))
		{
			$result[] = array('name' => 'minLength', 'attributes' => $this->_minLength);
		}
		// 10. maxLength
		if (!is_null($this->_maxLength))
		{
			$result[] = array('name' => 'maxLength', 'attributes' => $this->_maxLength);
		}
		// 11. showValue
		// Excel stores this attribute in the dataBar element
		/*if (!is_null($this->_showValue))
		{
			$result[] = array('name' => 'showValue', 'attributes' => $this->_showValue);
		}*/
		// 12. borderColor (only id borderColor also exists)
		if (($this->_border) && ($this->_borderColor))
		{
			$result[] = array('name' => 'border', 'attributes' => $this->_border);
		}
		// 13. gradient
		if (!is_null($this->_gradient))
		{
			$result[] = array('name' => 'gradient', 'attributes' => $this->_gradient);
		}
		// 14. direction
		if ($this->_direction)
		{
			$result[] = array('name' => 'direction', 'attributes' => $this->getDirection()->toString());
		}
		// 14. negativeBarColorSameAsPositive (only if negativeFillColor is set )
		if (!is_null($this->_negativeBarColorSameAsPositive) && ($this->_negativeFillColor))
		{
			$result[] = array('name' => 'negativeBarColorSameAsPositive', 'attributes' => $this->_negativeBarColorSameAsPositive);
		}
		// 15. negativeBarBorderColorSameAsPositive (only if negativeBorderColor and border are set )
		if (!is_null($this->_negativeBarBorderColorSameAsPositive) && ($this->_negativeBorderColor) && ($this->_border))
		{
			$result[] = array('name' => 'negativeBarBorderColorSameAsPositive', 'attributes' => $this->_negativeBarBorderColorSameAsPositive);
		}
		// 16. axisPosition
		if ($this->_axisPosition)
		{
			$axisPosition = $this->_axisPosition->toString();
			if ($axisPosition == PHPExcel_Style_DataBar_DataBarAxisPosition::NONE)
			{
				$result[] = array('name' => 'axisPosition', 'attributes' => $axisPosition);
			}
			else
			{
				$result[] = array('name' => 'axisPosition', 'attributes' => $axisPosition);
			}
		}
		
		// return the resulting array
		return $result;
	}
	
	/*
	 * get the color for the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $color = $worksheetstyles[0]->getColor() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_Color
	*/
	public function getColor() 
	{
    	return $this->_color;
    }
	
	/*	
	 * Set the fill color of the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_RED)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	fill color of the databar
	*/
	public function setColor($color = null) 
	{
   		$this->_color = $color;
   		return $this;
    }


	/*
	 * Get the cfvo settings
	 * 
	 * @params	boolean	set to true if the extlst cfvo settings must be used
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $cfvo = $worksheetstyles[0]->getCfvo() ; }
	 * </code>
	 *
	 * @return array of CFVOType	all CFVOTypes for this databar	
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
	* if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
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
	 * Get the fill color
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $fillColor = $worksheetstyles[0]->getFillColor() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_Color
	*/
	public function getFillColor() 
	{ 
		return $this->_fillColor;
	}
	
	/*
	 * Set the fill color of the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setFillColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_GREEN)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	fill color of the databar
	*/
	public function setFillColor($value) 
	{ 
		$this->_fillColor = $value;
		return $this;
	}
	
	/*
	 * Get the border color
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $borderColor = $worksheetstyles[0]->getBorderColor() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_Color
	*/
	public function getBorderColor() 
	{ 
		return $this->_borderColor;
	}
	
	/*
	 * Set the border color of the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setBorderColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_BLUE)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	border color of the databar
	*/
	public function setBorderColor($value) 
	{ 
		$this->_borderColor = $value;
		return $this;
	}
	
	/*
	 * Get the negative fill color
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $negativeFillColor = $worksheetstyles[0]->getNegativeFillColor() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_Color
	*/
	public function getNegativeFillColor() 
	{ 
		return $this->_negativeFillColor;
	}
	
	/*
	 * Set the fill color of the negative part of databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setNegativeFillColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_GREEN)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	fill color of the negative part of databar
	*/
	public function setNegativeFillColor($value) 
	{ 
		$this->_negativeFillColor = $value;
		return $this;
	}
	
	/*
	 * Get the negative border color
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $negativeBorderColor = $worksheetstyles[0]->getNegativeBorderColor() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_Color
	*/
	public function getNegativeBorderColor() 
	{ 
		return $this->_negativeBorderColor;
	}
	
	/*
	 * Set the axis color of the negative part of databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setNegativeBorderColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_BLUE)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	axis color of the negative part of databar
	*/
	public function setNegativeBorderColor($value) 
	{ 
		$this->_negativeBorderColor = $value;
		return $this;
	}
	
	/*
	 * Get the axis color
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $axisColor = $worksheetstyles[0]->getAxisColor() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_Color
	*/
	public function getAxisColor() 
	{ 
		return $this->_axisColor;
	}
	
	/*
	 * Set the axis color of the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setAxisColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_BLUE)); 
	 * }
	 * </code>
	 *
	 * @param	PHPExcel_Style_Color	axis color
	*/
	public function setAxisColor($value) 
	{ 
		$this->_axisColor = $value;
		return $this;
	}
	
	/*
	 * Get the minimum length of the bar
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $minLength = $worksheetstyles[0]->getMinLength() ; }
	 * </code>
	 *
	 * @return unsigned int
	*/
	public function getMinLength() 
	{ 
		return $this->_minLength;
	}
	
	/*
	 * Set the minimum length of the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setMinLength(50); 
	 * }
	 * </code>
	 *
	 * @param	unigned int	minimum length of the databar
	*/
	public function setMinLength($value) 
	{ 
		if (($value >= 0)&&($value > 100))
		{
			throw new PHPExcel_Exception("DataBar : minLength should be in the range of 0 to 100");
		}
		$this->_minLength = $value;
		return $this;
	}
	
	/*
	 * Get the maximum length of the bar
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $maxLength = $worksheetstyles[0]->getMaxLength() ; }
	 * </code>
	 *
	 * @return unsigned int
	*/
	public function getMaxLength() 
	{ 
		return $this->_maxLength;
	}
	
	/*
	 * Set the maximum length of the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setMaxLength(50); 
	 * }
	 * </code>
	 *
	 * @param	unigned int	maximum length of the databar
	*/
	public function setMaxLength($value) 
	{ 
		if (($value >= 0)&&($value > 100))
		{
			throw new PHPExcel_Exception("DataBar : maxLength should be in the range of 0 to 100");
		}
		$this->_maxLength = $value;
		return $this;
	}
	
	/*
	 * Get whether the cell value will be shown
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $showValue = $worksheetstyles[0]->getShowValue() ; }
	 * </code>
	 *
	 * @return integer (0=false;1=true)
	*/
	public function getShowValue() 
	{ 
		return $this->_showValue;
	}
	
	/*
	 * Indicate whether the to show the value of the cell
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setShowValue(1); 
	 * }
	 * </code>
	 *
	 * @param	integer	1 if the value of the cell must be shown
	*/
	public function setShowValue($value) 
	{ 
		$this->_showValue = $value;
		return $this;
	}
	
	/*
	 * Get whether the show the border
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $showBorder = $worksheetstyles[0]->getBorder() ; }
	 * </code>
	 *
	 * @return integer
	*/
	public function getBorder() 
	{ 
		return $this->_border;
	}
	
	/*
	 * Indicate whether the to use a border for the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setBorder(1); 
	 * }
	 * </code>
	 *
	 * @param	integer	1 if a border must be used for the databar
	*/
	public function setBorder($value) 
	{ 
		$this->_border = $value;
		return $this;
	}
	
	/*
	 * Get whether to use a gradient
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $showGradient = $worksheetstyles[0]->getGradient() ; }
	 * </code>
	 *
	 * @return bool
	*/
	public function getGradient() 
	{ 
		return $this->_gradient;
	}
	
	/*
	 * Indicate whether the to use a gradient for the databar
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setGradient(0); 
	 * }
	 * </code>
	 *
	 * @param	integer	0 if no gradient must be used
	*/
	public function setGradient($value) 
	{ 
		$this->_gradient = $value;
		return $this;
	}
	
	/*
	 * Get the direction of the databar
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $direction = $worksheetstyles[0]->getDirection() ; }
	 * </code>
	 *
	 * @return PHPExcel_Style_DataBar_DataBarAxisDirection
	*/
	public function getDirection() 
	{ 
		return $this->_direction;
	}
	
	/*
	 * Set the axis direction setting
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setDirection(new PHPExcel_Style_DataBar_DataBarDirection(PHPExcel_Style_DataBar_DataBarDirection::fromString('context'))); 
	 * }
	 * </code>
	 *
	 * @param PHPExcel_Style_DataBar_DataBarDirection
	*/
	public function setDirection($value) 
	{ 
		$this->_direction = $value;
		return $this;
	}
	
	/*
	 * Get whether to use the same color for the negative part of the databar
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 *   $useSameColorForNegative = $worksheetstyles[0]->getNegativeBarColorSameAsPositive() ; 
	 * }
	 * </code>
	 *
	 * @return integer 0=false; 1=true
	*/
	public function getNegativeBarColorSameAsPositive() 
	{ 
		return $this->_negativeBarColorSameAsPositive;
	}
	
	/*
	 * Indicate whether the negative bar color needs to be the same as the positive bar color
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setNegativeBarColorSameAsPositive(true); 
	 * }
	 * </code>
	 *
	 * @param	integer	1=true if the negative bar color needs to be the same as the positive bar color
	*/
	public function setNegativeBarColorSameAsPositive($value) 
	{ 
		$this->_negativeBarColorSameAsPositive = $value;
		return $this;
	}
	
	/*
	 * Get whether to use the same color for the border of the negative part of the databar
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 *   $useSameColorForNegativeBorder = $worksheetstyles[0]->getNegativeBarBorderColorSameAsPositive() ; 
	 * }
	 * </code>
	 *
	 * @return integer (0=false;1=true)
	*/
	public function getNegativeBarBorderColorSameAsPositive() 
	{ 
		return $this->_negativeBarBorderColorSameAsPositive;
	}
	
	/*
	 * Indicate whether the negative bar border color needs to be the same as the positive bar border color
	 * 
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setNegativeBarBorderColorSameAsPositive(1); 
	 * }
	 * </code>
	 *
	 * @param	integer	1(=true) if the negative bar border color needs to be the same as the positive bar border color
	*/
	public function setNegativeBarBorderColorSameAsPositive($value) 
	{ 
		$this->_negativeBarBorderColorSameAsPositive = $value;
		return $this;
	}

	/*
	 * Get the axis position setting
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) { $axisPosition = $worksheetstyles[0]->getAxisPosition(); }
	 * </code>
	 *
	 * @return PHPExcel_Style_DataBar_DataBarAxisPosition
	*/
	public function getAxisPosition() 
	{ 
		return $this->_axisPosition;
	}
	
	/*
	 * Set the axis position setting
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * if ($worksheetstyles[0]->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) 
	 * { 
	 * 	 $worksheetstyles[0]->setAxisPosition(new PHPExcel_Style_DataBar_DataBarAxisPosition(PHPExcel_Style_DataBar_DataBarAxisPosition::fromString('automatic'))); 
	 * }
	 * </code>
	 *
	 * @param PHPExcel_Style_DataBar_DataBarAxisPosition
	*/
	public function setAxisPosition($value) 
	{ 
		$this->_axisPosition = $value;
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
	 * Set the namespace setting for creating the extLst entry
	 * Default already set to http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * $worksheetstyles[0]->setNamespace('http://schemas.microsoft.com/office/spreadsheetml/2009/9/main)'); 
	 * </code>
	 * 
	 * @return string containing namespace setting
	*/
	public function setNamespace($namespace)
	{
		$this->_namespace = $namespace;
	}
	
	/*
	 * Get the namespace setting for creating the extLst entry
	 *
	 * <code>
	 * $worksheetstyles = $objPHPExcel->getActiveSheet()->getConditionalStyles();
	 * $namespace = $worksheetstyles[0]->getNamespace(); 
	 * </code>
	 *
	 * @return string containing namespace setting
	*/
	public function getNamespace()
	{
		return $this->_namespace;
	}	
	
	/*
	 * Get a unique classID for this object
	 *
	 * @return	string	CLASSID v3 string : e.g. {1546058F-5A25-4334-85AE-E68F2A44BBAF}
	 *
	 */
	public function getClassID()
	{
		$hash = $this->getHashCode();
		return sprintf('{%08s-%04s-%04x-%04x-%12s}',
						substr($hash, 0, 8),
						substr($hash, 8, 4),
						(hexdec(substr($hash, 12, 4)) & 0x0fff) | 0x3000,
						 (hexdec(substr($hash, 16, 4)) & 0x3fff) | 0x8000,
						  substr($hash, 20, 12));
	}
	
	/*
	 * Get an array containing the databar data that must be written to the extLst entry of the worksheet
	 *
	 * @params priority not used
	 * @return	array	array containing all the data that must be written to the extLst entry of the worksheert
	 *
	 */
	public function getExtLstData($priority)
	{
		$forExtLst = true;
		$data = $this->getElementsAsArray($forExtLst);
			// add id andtype to 
		$result = array('name' => 'cfRule',
						'cellReference' => $this->getCellReference(),
						'attributes' => array(array('name' => 'type',   'attributes' => 'dataBar'),
											   array('name' => 'id',     'attributes' => $this->getClassID()),
											   array('name' => 'dataBar','attributes' => $data)));
		
		return $result;
	}
	
	/*
	 * create a datastructure for creating the databar element with default properties
	 *
	 * @param	array	array with properties 
	 */
	public function getDefaultData()
	{
		$cfvos = $this->getCfvos();
		$result = array('name' => 'dataBar',
		                'attributes' => array($cfvos[0]->toArray(),
											  $cfvos[1]->toArray(),
											  array('name' => 'color', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->getColor()->getARGB())))
											  )
						);
		if (!is_null($this->_showValue)) { $result['attributes'][] = array('name' => 'showValue' , 'attributes' => $this->_showValue); }
		return $result;
	}
	
	/*
	 * Indicates whether this object needs a reference to the entry in the extLst section
	 *
	 * @returns	bool	true is such a reference is needed
	 *
	 */
	 public function needsExtLstReference()
	 {
		return true;
	 }

}
