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
 * @package    PHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Style_Conditional
 *
 * @category   PHPExcel
 * @package    PHPExcel_Style
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Style_Conditional implements PHPExcel_IComparable
{
    /* Condition types */
    const CONDITION_NONE                    = 'none';
    const CONDITION_CELLIS                    = 'cellIs';
    const CONDITION_CONTAINSTEXT            = 'containsText';
    const CONDITION_EXPRESSION                 = 'expression';
    const CONDITION_DATABAR                 = 'dataBar';
    const CONDITION_COLORSCALE                 = 'colorScale';
    const CONDITION_ICONSET                    = 'iconSet';
    const CONDITION_TIMEPERIOD                 = 'timePeriod';
    const CONDITION_DUPLICATEVALUES             = 'duplicateValues';
    const CONDITION_TOP10                     = 'top10';
    const CONDITION_ABOVEAVERAGE             = 'aboveAverage';
    

    /* Operator types */
    const OPERATOR_NONE                        = '';
    const OPERATOR_BEGINSWITH                = 'beginsWith';
    const OPERATOR_ENDSWITH                    = 'endsWith';
    const OPERATOR_EQUAL                    = 'equal';
    const OPERATOR_GREATERTHAN                = 'greaterThan';
    const OPERATOR_GREATERTHANOREQUAL        = 'greaterThanOrEqual';
    const OPERATOR_LESSTHAN                    = 'lessThan';
    const OPERATOR_LESSTHANOREQUAL            = 'lessThanOrEqual';
    const OPERATOR_NOTEQUAL                    = 'notEqual';
    const OPERATOR_CONTAINSTEXT                = 'containsText';
    const OPERATOR_NOTCONTAINS                = 'notContains';
    const OPERATOR_BETWEEN                    = 'between';

    /**
     * Condition type
     *
     * @var int
     */
    private $_conditionType;

    /**
     *
     * Priority
     *
     * @var int
     */
    private $_priority;
     
    /**
     * Operator type
     *
     * @var int
     */
    private $_operatorType;

    /**
     * Text
     *
     * @var string
     */
    private $_text;

    /**
     * Condition
     *
     * @var string[]
     */
    private $_condition = array();

    /**
     * Style
     *
     * @var PHPExcel_Style
     */
    private $_style;

    /**
     *
     * aboveAverage (setting for CONDITION_ABOVEAVERAGE)
     *
     * @var string
     */
    private $_aboveAverage;

    /**
     *
     * percent  (setting for CONDITION_TOP10)
     *
     * @var string
     */
    private $_percent;

    /**
     *
     * rank  (setting for CONDITION_TOP10)
     *
     * @var string
     */
    private $_rank;
    
    /**
     *
     * timePeriod  (setting for CONDITION_TIMEPERIOD)
     *
     * @var string
     */
    private $_timePeriod;

    /**
     *
     * bottom  (setting for CONDITION_TOP10)
     *
     * @var string
     */
    private $_bottom;
    
    /**
     * ConditionObject
     *
     * @var linked ConditionObject
     */
    private $_conditionObject;

    /**
     * Cell reference (in case a conditionObject is defined, that cell reference will be used instead)
     *
     * @var
     */
    private $_cellReference;    
    
    /**
     * Create a new PHPExcel_Style_Conditional
     */
    public function __construct()
    {
        // Initialise values
        $this->_conditionType        = PHPExcel_Style_Conditional::CONDITION_NONE;
        $this->_operatorType        = PHPExcel_Style_Conditional::OPERATOR_NONE;
        $this->_text                = null;
        $this->_priority            = 0;
        $this->_condition            = array();
        $this->_style                = new PHPExcel_Style(FALSE, TRUE);
    }

    /**
     * Get Condition type
     *
     * @return string
     */
    public function getConditionType() {
        return $this->_conditionType;
    }

    /**
     * Set Condition type
     *
     * @param string $pValue    PHPExcel_Style_Conditional condition type
     * @return PHPExcel_Style_Conditional
     */
    public function setConditionType($pValue = PHPExcel_Style_Conditional::CONDITION_NONE) {
        $this->_conditionType = $pValue;
        return $this;
    }
    
     /**
     * Get priority 
     *
     * @return int (0 if not set)
     */
    public function getPriority() {
        return $this->_priority;
    }

    /**
     * Set priority
     *
     * @param int $pValue    priority
     * @return PHPExcel_Style_Conditional
     */
    public function setPriority($pValue = 0) {
        $this->_priority = $pValue;
        return $this;
    }
    
         /**
     * Get aboveAverage 
     *
     * @return string
     */
    public function getAboveAverage() {
        return $this->_aboveAverage;
    }

    /**
     * Set aboveAverage
     *
     * @param string $pValue    aboveAverage setting
     * @return PHPExcel_Style_Conditional
     */
    public function setAboveAverage($pValue) {
        $this->_aboveAverage = $pValue;
        return $this;
    }
    
     /**
     * Get rank 
     *
     * @return string
     */
    public function getRank() {
        return $this->_rank;
    }

    /**
     * Set rank
     *
     * @param string $pValue    rank
     * @return PHPExcel_Style_Conditional
     */
    public function setRank($pValue) {
        $this->_rank = $pValue;
        return $this;
    }

     /**
     * Get timePeriod 
     *
     * @return string
     */
    public function getTimePeriod() {
        return $this->_timePeriod;
    }

    /**
     * Set timePeriod
     *
     * @param string $pValue    timePeriod
     * @return PHPExcel_Style_Conditional
     */
    public function setTimePeriod($pValue) {
        $this->_timePeriod = $pValue;
        return $this;
    }
    
     /**
     * Get bottom 
     *
     * @return string
     */
    public function getBottom() {
        return $this->_bottom;
    }

    /**
     * Set bottom
     *
     * @param string $pValue    bottom
     * @return PHPExcel_Style_Conditional
     */
    public function setBottom($pValue) {
        $this->_bottom = $pValue;
        return $this;
    }
    
    /**
     * Get percent 
     *
     * @return string
     */
    public function getPercent() {
        return $this->_percent;
    }

    /**
     * Set percent
     *
     * @param string $pValue    percent
     * @return PHPExcel_Style_Conditional
     */
    public function setPercent($pValue) {
        $this->_percent = $pValue;
        return $this;
    }

     /**
     * Get cellReference 
     *
     * @return string
     */
    public function getCellReference() {
        if ($this->_conditionObject)
        {
            // get it from the conditional object
            return $this->_conditionObject->getCellReference();
        }
        else
        {
            return $this->_cellReference;
        }
    }

    public function removeCellReference($position)
    {
        if ($this->_conditionObject)
        {
            // remove the position from the conditionObject's list of cellReferences
            $this->_conditionObject->removeCellReference($position);
        }
        else
        {
            $this->_cellReference = null;
        }
        return $this;
    }
        
    /**
     * Set cellReference
     *
     * @param string $pValue    cellReference
     * @return PHPExcel_Style_Conditional
     */
    public function setCellReference($pValue) {
        if ($this->_conditionObject)
        {
            // set it in the conditional object
            $this->_conditionObject->setCellReference($pValue);
        }
        else
        {
            $this->_cellReference = $pValue;
        }
        return $this;
    }


    /**
     * Get Operator type
     *
     * @return string
     */
    public function getOperatorType() {
        return $this->_operatorType;
    }

    /**
     * Set Operator type
     *
     * @param string $pValue    PHPExcel_Style_Conditional operator type
     * @return PHPExcel_Style_Conditional
     */
    public function setOperatorType($pValue = PHPExcel_Style_Conditional::OPERATOR_NONE) {
        $this->_operatorType = $pValue;
        return $this;
    }

    /**
     * Get text
     *
     * @return string
     */
    public function getText() {
        return $this->_text;
    }

    /**
     * Set text
     *
     * @param string $value
     * @return PHPExcel_Style_Conditional
     */
    public function setText($value = null) {
           $this->_text = $value;
           return $this;
    }

    /**
     * Get Condition
     *
     * @deprecated Deprecated, use getConditions instead
     * @return string
     */
    public function getCondition() {
        if (isset($this->_condition[0])) {
            return $this->_condition[0];
        }

        return '';
    }

    /**
     * Set Condition
     *
     * @deprecated Deprecated, use setConditions instead
     * @param string $pValue    Condition
     * @return PHPExcel_Style_Conditional
     */
    public function setCondition($pValue = '') {
        if (!is_array($pValue))
            $pValue = array($pValue);

        return $this->setConditions($pValue);
    }

    /**
     * Get Conditions
     *
     * @return string[]
     */
    public function getConditions() {
        return $this->_condition;
    }

    /**
     * Set Conditions
     *
     * @param string[] $pValue    Condition
     * @return PHPExcel_Style_Conditional
     */
    public function setConditions($pValue) {
        if (!is_array($pValue))
            $pValue = array($pValue);

        $this->_condition = $pValue;
        return $this;
    }

    /**
     * Add Condition
     *
     * @param string $pValue    Condition
     * @return PHPExcel_Style_Conditional
     */
    public function addCondition($pValue = '') {
        $this->_condition[] = $pValue;
        return $this;
    }

    /**
     * Get Style
     *
     * @return PHPExcel_Style
     */
    public function getStyle() {
        return $this->_style;
    }

    /**
     * Set Style
     *
     * @param     PHPExcel_Style $pValue
     * @throws     PHPExcel_Exception
     * @return PHPExcel_Style_Conditional
     */
    public function setStyle(PHPExcel_Style $pValue = null) {
           $this->_style = $pValue;
           return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode() {
        return md5(
              $this->_conditionType
            . $this->_priority
            . $this->_operatorType
            . $this->_aboveAverage
            . $this->_percent
            . $this->_rank
            . $this->_timePeriod
            . $this->_bottom
            . implode(';', $this->_condition)
            . $this->_style->getHashCode()
            . __CLASS__
        );
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }

    /*
    * return the Condition object that is related to this conditional
    *
    * @return     PHPExcel_Style_DataBar or PHPExcel_Style_ColorScale    object
    * @throws    PHPExcel_Exception    in case condition type is not  CONDITION_DATABAR or CONDITION_COLORSCALE or CONDITION_ICONSET
    */
    public function getConditionalObject()
    {
        if (($this->_conditionType == PHPExcel_Style_Conditional::CONDITION_DATABAR) ||
            ($this->_conditionType == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) ||
            ($this->_conditionType == PHPExcel_Style_Conditional::CONDITION_ICONSET))
        {
            return $this->_conditionObject;
        }
        else
        {
            throw new PHPExcel_Exception("ConditionType is not a databar or colorscale.");
        }
    }

    /*
    * set conditional object (Databar, Colorscale, IconSet) that is related to this conditional
    *
    */
    public function setConditionalObject($ref, $rule, $extLst)
    {
        // create the object
        if ($this->_conditionType == PHPExcel_Style_Conditional::CONDITION_DATABAR)
        {
            $this->_conditionObject = new PHPExcel_Style_DataBar();
        }
        elseif    ($this->_conditionType == PHPExcel_Style_Conditional::CONDITION_COLORSCALE)
        {
            $this->_conditionObject = new PHPExcel_Style_ColorScale();
        }
        elseif    ($this->_conditionType == PHPExcel_Style_Conditional::CONDITION_ICONSET)
        {
            $this->_conditionObject = new PHPExcel_Style_IconSet();
        }
        else
        {
            throw new PHPExcel_Exception("type is not a databar, colorscale or iconset.");
        }
        // initialize the object
        $this->_conditionObject->applyFromXML($ref, $rule, $extLst);
    }

    public function updateCellReference($oldCoordinates, $newCoordinates)
    {
        if ($this->_conditionObject)
        {
            // set it in the conditional object
            $this->_conditionObject->updateCellReference($oldCoordinates, $newCoordinates);
        }
        else
        {
            $this->_cellReference = $newCoordinates;
        }
        return $this;
    }

    public function setConditionObjectFromArray($conditionType, $arr)
    {
        // initialize the object
        if ($conditionType == PHPExcel_Style_Conditional::CONDITION_DATABAR)
        {
            $this->_conditionObject = new PHPExcel_Style_DataBar();
        }
        elseif ($conditionType == PHPExcel_Style_Conditional::CONDITION_COLORSCALE)
        {
            $this->_conditionObject = new PHPExcel_Style_ColorScale();
        }
        elseif ($conditionType == PHPExcel_Style_Conditional::CONDITION_ICONSET)
        {
            $this->_conditionObject = new PHPExcel_Style_IconSet();
        }
        else
        {
            throw new PHPExcel_Exception("type is not a databar, colorscale or iconset.");
        }
        // set the type
        $this->_conditionType = $conditionType;
        // apply the array
        $this->_conditionObject->applyFromArray($arr,true);
    }
}
