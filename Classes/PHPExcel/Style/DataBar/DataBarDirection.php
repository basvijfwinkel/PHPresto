<?PHP
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
 * @copyright Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version ##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Style_DataBar_DataBarDirection
 *
 * @category   PHPExcel
 * @package	PHPExcel_Style_DataBar
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @author	Bas Vijfwinkel
 */
class PHPExcel_Style_DataBar_DataBarDirection
{
	const CONTEXT = 'context';
	const LEFTTORIGHT = 'leftToRight';
	const RIGHTTOLEFT = 'rightToLeft';
	
	protected $_direction;
	
	/*
	* constructor : PHPExcel_Style_DataBar_DataBarDirection object should be created with 'fromString' method
	*
	* @params	string 	$direction	valid databar axis direction
	* 
	*/
	protected function __construct($direction)
	{
		$this->_direction = $direction;
	}
	
	/**
	 * check if the databar axis direction is correct and return a PHPExcel_Style_DataBar_DataBarDirection object representing this direction
	 * in case an unknown value is passed, an exception will be thrown
	 *
	 * <code>
	 * $databardirection = PHPExcel_Style_DataBar_DataBarDirection::fromString($stringValue)
	 * </code>
	 *
	 * @param	string	$type	string containing databar direction information
	 * @throws	PHPExcel_Exception
	 * @return string	databar direction as a string value
	 */
	public static function fromString($type)
	{
		if (is_string($type))
		{
			switch($type)
			{
				case 'context':
									return new PHPExcel_Style_DataBar_DataBarDirection(PHPExcel_Style_DataBar_DataBarDirection::CONTEXT);
									break;
				case 'leftToRight':
									return new PHPExcel_Style_DataBar_DataBarDirection(PHPExcel_Style_DataBar_DataBarDirection::LEFTTORIGHT);
									break;
				case 'rightToLeft':
									return new PHPExcel_Style_DataBar_DataBarDirection(PHPExcel_Style_DataBar_DataBarDirection::RIGHTTOLEFT);
									break;
			}
		}
		// unknown type
		throw new PHPExcel_Exception("Invalid DataBarDirection passed.:".$type);
	}
	
	/*
	* Return the databar axis direction as a string
	*
	* @return	string	databar axis direction 
	*/
	public function toString()
	{
		return "$this->_direction";
	}
	
}
?>