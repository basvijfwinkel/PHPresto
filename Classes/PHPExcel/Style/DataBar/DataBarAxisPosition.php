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
 * PHPExcel_Style_DataBar_DataBarAxisPosition
 *
 * @category   PHPExcel
 * @package	PHPExcel_Style_DataBar
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @author	Bas Vijfwinkel
 */
class PHPExcel_Style_DataBar_DataBarAxisPosition
{
	const AUTOMATIC= 'automatic';
	const MIDDLE = 'middle';
	const NONE = 'none';
	
	protected $_position;
	
	/*
	* constructor : PHPExcel_Style_DataBar_DataBarAxisPosition object should be created with 'fromString' method
	*
	* @params	string 	$position	valid databar axis position
	* 
	*/
	protected function __construct($position)
	{
		$this->_position = $position;
	}
	
	/**
	 * check if the databar axis position is correct and return a PHPExcel_Style_DataBar_DataBarAxisPosition that represents this position
	 * in case an unknown value is passed, an exception will be thrown
	 *
	 * <code>
	 * $databaraxisposition = PHPExcel_Style_DataBar_DataBarAxisPosition::fromString($typevalue)
	 * </code>
	 *
	 * @param	string	$type	string containing databar axis position information
	 * @throws	PHPExcel_Exception
	 * @return PHPExcel_Style_DataBar_DataBarAxisPosition	databar axis position
	 */
	public static function fromString($type)
	{
		if (is_string($type))
		{
			switch($type)
			{
				case 'automatic':
									return new PHPExcel_Style_DataBar_DataBarAxisPosition(PHPExcel_Style_DataBar_DataBarAxisPosition::AUTOMATIC);
									break;
				case 'middle':
									return new PHPExcel_Style_DataBar_DataBarAxisPosition(PHPExcel_Style_DataBar_DataBarAxisPosition::MIDDLE);
									break;
				case 'none':
									return new PHPExcel_Style_DataBar_DataBarAxisPosition(PHPExcel_Style_DataBar_DataBarAxisPosition::NONE);
									break;
			}
		}
		// unknown type
		throw new PHPExcel_Exception("Invalid DataBarAxisPosition string passed.");
	}
	
	/*
	* Return the databar axis position as a string
	*
	* @return	string	databar axis position 
	*/
	public function toString()
	{
		return "$this->_position";
	}
}
?>