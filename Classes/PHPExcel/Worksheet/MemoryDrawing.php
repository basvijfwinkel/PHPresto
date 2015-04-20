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
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Worksheet_MemoryDrawing
 *
 * @category   PHPExcel
 * @package    PHPExcel_Worksheet
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Worksheet_MemoryDrawing extends PHPExcel_Worksheet_BaseDrawing implements PHPExcel_IComparable
{
	/* Rendering functions */
	const RENDERING_DEFAULT					= 'imagepng';
	const RENDERING_PNG						= 'imagepng';
	const RENDERING_GIF						= 'imagegif';
	const RENDERING_JPEG					= 'imagejpeg';

	/* MIME types */
	const MIMETYPE_DEFAULT					= 'image/png';
	const MIMETYPE_PNG						= 'image/png';
	const MIMETYPE_GIF						= 'image/gif';
	const MIMETYPE_JPEG						= 'image/jpeg';

	/**
	 * Image resource
	 *
	 * @var resource
	 */
	private $_imageResource;

	/**
	 * Rendering function
	 *
	 * @var string
	 */
	private $_renderingFunction;

	/**
	 * Mime type
	 *
	 * @var string
	 */
	private $_mimeType;

	/**
	 * Unique name
	 *
	 * @var string
	 */
	private $_uniqueName;
	
	/**
	 * Hashtag for the image we referencing
	 *
	 * @var string
	 */
	protected $_referenceHashTag;

	/**
	 * rId of the image
	 *
	 * @var string
	 */
	protected $_rId;
	
    /**
     * Create a new PHPExcel_Worksheet_MemoryDrawing
     */
    public function __construct()
    {
    	// Initialise values
    	$this->_imageResource		= null;
    	$this->_renderingFunction 	= self::RENDERING_DEFAULT;
    	$this->_mimeType			= self::MIMETYPE_DEFAULT;
    	$this->_uniqueName			= md5(rand(0, 9999). time() . rand(0, 9999));
		$this->_referenceHashTag	= null;
		$this->_rId					= null;

    	// Initialize parent
    	parent::__construct();
    }

    /**
     * Get image resource
     *
     * @return resource
     */
    public function getImageResource() {
    	return $this->_imageResource;
    }

    /**
     * Set image resource
     *
     * @param	$value resource
     * @return PHPExcel_Worksheet_MemoryDrawing
     */
    public function setImageResource($value = null) {
    	$this->_imageResource = $value;

    	if (!is_null($this->_imageResource)) {
	    	// Get width/height
	    	$this->_width	= imagesx($this->_imageResource);
	    	$this->_height	= imagesy($this->_imageResource);
    	}
    	return $this;
    }

    /**
     * Get rendering function
     *
     * @return string
     */
    public function getRenderingFunction() {
    	return $this->_renderingFunction;
    }

    /**
     * Set rendering function
     *
     * @param string $value
     * @return PHPExcel_Worksheet_MemoryDrawing
     */
    public function setRenderingFunction($value = PHPExcel_Worksheet_MemoryDrawing::RENDERING_DEFAULT) {
    	$this->_renderingFunction = $value;
    	return $this;
    }

    /**
     * Get mime type
     *
     * @return string
     */
    public function getMimeType() {
    	return $this->_mimeType;
    }

    /**
     * Set mime type
     *
     * @param string $value
     * @return PHPExcel_Worksheet_MemoryDrawing
     */
    public function setMimeType($value = PHPExcel_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT) {
    	$this->_mimeType = $value;
    	return $this;
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename($index=null) 
	{
		$extension 	= strtolower($this->getMimeType());
		$extension 	= explode('/', $extension);
		$extension 	= $extension[1];
		if (is_null($index)) 
		{
			if (!is_null($this->_rId))
			{
				$index = $this->_rId;
			}
			else
			{
				die('Media should already have been saved');
			}
		}
		return $this->_uniqueName . $index . '.' . $extension;
    }

	/**
	 * Get hash code
	 *
	 * @return string	Hash code
	 */
	public function getHashCode() {
    	return md5(
    		  $this->_renderingFunction
    		. $this->_mimeType
    		. $this->_uniqueName
    		. parent::getHashCode()
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
	* set the hashcode of the image we are referencing
	*/

	public function setReferenceHashTag($hashTag)
	{
		$this->_referenceHashTag = $hashTag;
	}
	
	/*
	* get the hashcode of the image we are referencing
	*/
	public function getReferenceHashTag()
	{
		return $this->_referenceHashTag ;
	}
	
	/*
	* set the rId of the image
	*/
	
	public function setMediaReferenceId($rId)
	{
		$this->_rId = $rId;
	}
	
	/*
	* get the rId of the image
	*/
	public function getMediaReferenceId()
	{
		return $this->_rId;
	}
	
	/*
	* explicitly set the width and height of the resource
	* This is only used by images that reference another image
	*/
	public function setWidthAndHeightForced($width, $height)
	{
		$this->_width = $width;
		$this->_height = $height;
	}
}
