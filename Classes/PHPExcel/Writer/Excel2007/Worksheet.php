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
 * @package	PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license	http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version	##VERSION##, ##DATE##
 */


/**
 * PHPExcel_Writer_Excel2007_Worksheet
 *
 * @category   PHPExcel
 * @package	PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_Excel2007_Worksheet extends PHPExcel_Writer_Excel2007_WriterPart
{
	const EXTLST_CONDITIONALFORMATTINGID = 14;

	/**
	* extLst data
	* nested array, data grouped per grouptype (e.g. conditionalformatting rules)
	*
	* @var array
	*/
	private $_extlst = null;
	
	/**
	 * Write worksheet to XML format
	 *
	 * @param	PHPExcel_Worksheet		$pSheet
	 * @param	string[]				$pStringTable
	 * @param	boolean					$includeCharts	Flag indicating if we should write charts
	 * @return	string					XML Output
	 * @throws	PHPExcel_Writer_Exception
	 */
	public function writeWorksheet($pSheet = null, $pStringTable = null, $includeCharts = FALSE)
	{
		if (!is_null($pSheet)) {
			// Create XML writer
			$objWriter = null;
			if ($this->getParentWriter()->getUseDiskCaching()) {
				$objWriter = new PHPExcel_Shared_XMLWriter(PHPExcel_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
			} else {
				$objWriter = new PHPExcel_Shared_XMLWriter(PHPExcel_Shared_XMLWriter::STORAGE_MEMORY);
			}

			// Clear the extLst
			$this->clearExtLst();
			
			// XML header
			$objWriter->startDocument('1.0','UTF-8','yes');

			// Worksheet
			$objWriter->startElement('worksheet');
			$objWriter->writeAttribute('xml:space', 'preserve');
			$objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
			$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

				// sheetPr
				$this->_writeSheetPr($objWriter, $pSheet);

				// Dimension
				$this->_writeDimension($objWriter, $pSheet);

				// sheetViews
				$this->_writeSheetViews($objWriter, $pSheet);

				// sheetFormatPr
				$this->_writeSheetFormatPr($objWriter, $pSheet);

				// cols
				$this->_writeCols($objWriter, $pSheet);

				// sheetData
				$this->_writeSheetData($objWriter, $pSheet, $pStringTable);

				// sheetProtection
				$this->_writeSheetProtection($objWriter, $pSheet);

				// protectedRanges
				$this->_writeProtectedRanges($objWriter, $pSheet);

				// autoFilter
				$this->_writeAutoFilter($objWriter, $pSheet);

				// mergeCells
				$this->_writeMergeCells($objWriter, $pSheet);

				// conditionalFormatting
				$this->_writeConditionalFormatting($objWriter, $pSheet);

				// dataValidations
				$this->_writeDataValidations($objWriter, $pSheet);

				// hyperlinks
				$this->_writeHyperlinks($objWriter, $pSheet);

				// Print options
				$this->_writePrintOptions($objWriter, $pSheet);

				// Page margins
				$this->_writePageMargins($objWriter, $pSheet);

				// Page setup
				$this->_writePageSetup($objWriter, $pSheet);

				// Header / footer
				$this->_writeHeaderFooter($objWriter, $pSheet);

				// Breaks
				$this->_writeBreaks($objWriter, $pSheet);

				// Drawings and/or Charts
				$this->_writeDrawings($objWriter, $pSheet, $includeCharts);

				// LegacyDrawing
				$this->_writeLegacyDrawing($objWriter, $pSheet);

				// LegacyDrawingHF
				$this->_writeLegacyDrawingHF($objWriter, $pSheet);
				
				// extLst entries
				$this->_writeExtLstEntries($objWriter, $pSheet);

			$objWriter->endElement();

			
			// Return
			$result = $objWriter->getData();
			/* debug */
			//echo('<xmp style="white-space: pre-wrap">'.$result.'</xmp>');
			/* debug */
			return $result;
		} else {
			throw new PHPExcel_Writer_Exception("Invalid PHPExcel_Worksheet object passed.");
		}
	}

	/**
	 * Write SheetPr
	 *
	 * @param	PHPExcel_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet				$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeSheetPr(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// sheetPr
		$objWriter->startElement('sheetPr');
		//$objWriter->writeAttribute('codeName',		$pSheet->getTitle());
		if($pSheet->getParent()->hasMacros()){//if the workbook have macros, we need to have codeName for the sheet
			if($pSheet->hasCodeName()==false){
				$pSheet->setCodeName($pSheet->getTitle());
			}
			$objWriter->writeAttribute('codeName',		$pSheet->getCodeName());
		}
			$autoFilterRange = $pSheet->getAutoFilter()->getRange();
			if (!empty($autoFilterRange)) {
				$objWriter->writeAttribute('filterMode', 1);
				$pSheet->getAutoFilter()->showHideRows();
			}

			// tabColor
			if ($pSheet->isTabColorSet()) {
				$objWriter->startElement('tabColor');
				$objWriter->writeAttribute('rgb',	$pSheet->getTabColor()->getARGB());
				$objWriter->endElement();
			}

			// outlinePr
			$objWriter->startElement('outlinePr');
			$objWriter->writeAttribute('summaryBelow',	($pSheet->getShowSummaryBelow() ? '1' : '0'));
			$objWriter->writeAttribute('summaryRight',	($pSheet->getShowSummaryRight() ? '1' : '0'));
			$objWriter->endElement();

			// pageSetUpPr
			if ($pSheet->getPageSetup()->getFitToPage()) {
				$objWriter->startElement('pageSetUpPr');
				$objWriter->writeAttribute('fitToPage',	'1');
				$objWriter->endElement();
			}

		$objWriter->endElement();
	}

	/**
	 * Write Dimension
	 *
	 * @param	PHPExcel_Shared_XMLWriter	$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet			$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeDimension(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// dimension
		$objWriter->startElement('dimension');
		$objWriter->writeAttribute('ref', $pSheet->calculateWorksheetDimension());
		$objWriter->endElement();
	}

	/**
	 * Write SheetViews
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeSheetViews(PHPExcel_Shared_XMLWriter $objWriter = NULL, PHPExcel_Worksheet $pSheet = NULL)
	{
		// sheetViews
		$objWriter->startElement('sheetViews');

			// Sheet selected?
			$sheetSelected = false;
			if ($this->getParentWriter()->getPHPExcel()->getIndex($pSheet) == $this->getParentWriter()->getPHPExcel()->getActiveSheetIndex())
				$sheetSelected = true;


			// sheetView
			$objWriter->startElement('sheetView');
			$objWriter->writeAttribute('tabSelected',		$sheetSelected ? '1' : '0');
			$objWriter->writeAttribute('workbookViewId',	'0');

				// Zoom scales
				if ($pSheet->getSheetView()->getZoomScale() != 100) {
					$objWriter->writeAttribute('zoomScale',	$pSheet->getSheetView()->getZoomScale());
				}
				if ($pSheet->getSheetView()->getZoomScaleNormal() != 100) {
					$objWriter->writeAttribute('zoomScaleNormal',	$pSheet->getSheetView()->getZoomScaleNormal());
				}

				// View Layout Type
				if ($pSheet->getSheetView()->getView() !== PHPExcel_Worksheet_SheetView::SHEETVIEW_NORMAL) {
					$objWriter->writeAttribute('view',	$pSheet->getSheetView()->getView());
				}

				// Gridlines
				if ($pSheet->getShowGridlines()) {
					$objWriter->writeAttribute('showGridLines',	'true');
				} else {
					$objWriter->writeAttribute('showGridLines',	'false');
				}

				// Row and column headers
				if ($pSheet->getShowRowColHeaders()) {
					$objWriter->writeAttribute('showRowColHeaders', '1');
				} else {
					$objWriter->writeAttribute('showRowColHeaders', '0');
				}

				// Right-to-left
				if ($pSheet->getRightToLeft()) {
					$objWriter->writeAttribute('rightToLeft',	'true');
				}

				$activeCell = $pSheet->getActiveCell();

				// Pane
				$pane = '';
				$topLeftCell = $pSheet->getFreezePane();
				if (($topLeftCell != '') && ($topLeftCell != 'A1')) {
					$activeCell = $topLeftCell;
					// Calculate freeze coordinates
					$xSplit = $ySplit = 0;

					list($xSplit, $ySplit) = PHPExcel_Cell::coordinateFromString($topLeftCell);
					$xSplit = PHPExcel_Cell::columnIndexFromString($xSplit);

					// pane
					$pane = 'topRight';
					$objWriter->startElement('pane');
					if ($xSplit > 1)
						$objWriter->writeAttribute('xSplit',	$xSplit - 1);
					if ($ySplit > 1) {
						$objWriter->writeAttribute('ySplit',	$ySplit - 1);
						$pane = ($xSplit > 1) ? 'bottomRight' : 'bottomLeft';
					}
					$objWriter->writeAttribute('topLeftCell',	$topLeftCell);
					$objWriter->writeAttribute('activePane',	$pane);
					$objWriter->writeAttribute('state',		'frozen');
					$objWriter->endElement();

					if (($xSplit > 1) && ($ySplit > 1)) {
						//	Write additional selections if more than two panes (ie both an X and a Y split)
						$objWriter->startElement('selection');	$objWriter->writeAttribute('pane', 'topRight');		$objWriter->endElement();
						$objWriter->startElement('selection');	$objWriter->writeAttribute('pane', 'bottomLeft');	$objWriter->endElement();
					}
				}

				// Selection
//				if ($pane != '') {
					//	Only need to write selection element if we have a split pane
					//		We cheat a little by over-riding the active cell selection, setting it to the split cell
					$objWriter->startElement('selection');
					if ($pane != '') {
						$objWriter->writeAttribute('pane', $pane);
					}
					$objWriter->writeAttribute('activeCell', $activeCell);
					$objWriter->writeAttribute('sqref', $activeCell);
					$objWriter->endElement();
//				}

			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write SheetFormatPr
	 *
	 * @param	PHPExcel_Shared_XMLWriter $objWriter		XML Writer
	 * @param	PHPExcel_Worksheet		  $pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeSheetFormatPr(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// sheetFormatPr
		$objWriter->startElement('sheetFormatPr');

			// Default row height
			if ($pSheet->getDefaultRowDimension()->getRowHeight() >= 0) {
				$objWriter->writeAttribute('customHeight',		'true');
				$objWriter->writeAttribute('defaultRowHeight',	PHPExcel_Shared_String::FormatNumber($pSheet->getDefaultRowDimension()->getRowHeight()));
			} else {
				$objWriter->writeAttribute('defaultRowHeight', '14.4');
			}

			// Set Zero Height row
			if ((string)$pSheet->getDefaultRowDimension()->getZeroHeight()  == '1' ||
				strtolower((string)$pSheet->getDefaultRowDimension()->getZeroHeight()) == 'true' ) {
				$objWriter->writeAttribute('zeroHeight', '1');
			}

			// Default column width
			if ($pSheet->getDefaultColumnDimension()->getWidth() >= 0) {
				$objWriter->writeAttribute('defaultColWidth', PHPExcel_Shared_String::FormatNumber($pSheet->getDefaultColumnDimension()->getWidth()));
			}

			// Outline level - row
			$outlineLevelRow = 0;
			foreach ($pSheet->getRowDimensions() as $dimension) {
				if ($dimension->getOutlineLevel() > $outlineLevelRow) {
					$outlineLevelRow = $dimension->getOutlineLevel();
				}
			}
			$objWriter->writeAttribute('outlineLevelRow',	(int)$outlineLevelRow);

			// Outline level - column
			$outlineLevelCol = 0;
			foreach ($pSheet->getColumnDimensions() as $dimension) {
				if ($dimension->getOutlineLevel() > $outlineLevelCol) {
					$outlineLevelCol = $dimension->getOutlineLevel();
				}
			}
			$objWriter->writeAttribute('outlineLevelCol',	(int)$outlineLevelCol);

		$objWriter->endElement();
	}

	/**
	 * Write Cols
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeCols(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// cols
		if (count($pSheet->getColumnDimensions()) > 0)  {
			$objWriter->startElement('cols');

				$pSheet->calculateColumnWidths();

				// Loop through column dimensions
				foreach ($pSheet->getColumnDimensions() as $colDimension) {
					// col
					$objWriter->startElement('col');
					$objWriter->writeAttribute('min',	PHPExcel_Cell::columnIndexFromString($colDimension->getColumnIndex()));
					$objWriter->writeAttribute('max',	PHPExcel_Cell::columnIndexFromString($colDimension->getColumnIndex()));

					if ($colDimension->getWidth() < 0) {
						// No width set, apply default of 10
						$objWriter->writeAttribute('width',		'9.10');
					} else {
						// Width set
						$objWriter->writeAttribute('width',		PHPExcel_Shared_String::FormatNumber($colDimension->getWidth()));
					}

					// Column visibility
					if ($colDimension->getVisible() == false) {
						$objWriter->writeAttribute('hidden',		'true');
					}

					// Auto size?
					if ($colDimension->getAutoSize()) {
						$objWriter->writeAttribute('bestFit',		'true');
					}

					// Custom width?
					if ($colDimension->getWidth() != $pSheet->getDefaultColumnDimension()->getWidth()) {
						$objWriter->writeAttribute('customWidth',	'true');
					}

					// Collapsed
					if ($colDimension->getCollapsed() == true) {
						$objWriter->writeAttribute('collapsed',		'true');
					}

					// Outline level
					if ($colDimension->getOutlineLevel() > 0) {
						$objWriter->writeAttribute('outlineLevel',	$colDimension->getOutlineLevel());
					}

					// Style
					$objWriter->writeAttribute('style', $colDimension->getXfIndex());

					$objWriter->endElement();
				}

			$objWriter->endElement();
		}
	}

	/**
	 * Write SheetProtection
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeSheetProtection(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// sheetProtection
		$objWriter->startElement('sheetProtection');

		if ($pSheet->getProtection()->getPassword() != '') {
			$objWriter->writeAttribute('password',				$pSheet->getProtection()->getPassword());
		}

		$objWriter->writeAttribute('sheet',					($pSheet->getProtection()->getSheet()				? 'true' : 'false'));
		$objWriter->writeAttribute('objects',				($pSheet->getProtection()->getObjects()				? 'true' : 'false'));
		$objWriter->writeAttribute('scenarios',				($pSheet->getProtection()->getScenarios()			? 'true' : 'false'));
		$objWriter->writeAttribute('formatCells',			($pSheet->getProtection()->getFormatCells()			? 'true' : 'false'));
		$objWriter->writeAttribute('formatColumns',			($pSheet->getProtection()->getFormatColumns()		? 'true' : 'false'));
		$objWriter->writeAttribute('formatRows',			($pSheet->getProtection()->getFormatRows()			? 'true' : 'false'));
		$objWriter->writeAttribute('insertColumns',			($pSheet->getProtection()->getInsertColumns()		? 'true' : 'false'));
		$objWriter->writeAttribute('insertRows',			($pSheet->getProtection()->getInsertRows()			? 'true' : 'false'));
		$objWriter->writeAttribute('insertHyperlinks',		($pSheet->getProtection()->getInsertHyperlinks()	? 'true' : 'false'));
		$objWriter->writeAttribute('deleteColumns',			($pSheet->getProtection()->getDeleteColumns()		? 'true' : 'false'));
		$objWriter->writeAttribute('deleteRows',			($pSheet->getProtection()->getDeleteRows()			? 'true' : 'false'));
		$objWriter->writeAttribute('selectLockedCells',		($pSheet->getProtection()->getSelectLockedCells()	? 'true' : 'false'));
		$objWriter->writeAttribute('sort',					($pSheet->getProtection()->getSort()				? 'true' : 'false'));
		$objWriter->writeAttribute('autoFilter',			($pSheet->getProtection()->getAutoFilter()			? 'true' : 'false'));
		$objWriter->writeAttribute('pivotTables',			($pSheet->getProtection()->getPivotTables()			? 'true' : 'false'));
		$objWriter->writeAttribute('selectUnlockedCells',	($pSheet->getProtection()->getSelectUnlockedCells()	? 'true' : 'false'));
		$objWriter->endElement();
	}

	/**
	 * Write ConditionalFormatting
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeConditionalFormatting(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// Conditional id
		$id = 1;
		
		// Loop through styles in the current worksheet
		$processedCellReferences = array(); // conditional formats like databar/colorscale/iconset settings need to be applied as a group; not to each individual cell
		foreach ($pSheet->getConditionalStylesCollection() as $cellCoordinate => $conditionalStyles) {
			foreach ($conditionalStyles as $index => $conditional) {
				
				// WHY was this again?
				// if ($this->getParentWriter()->getStylesConditionalHashTable()->getIndexForHashCode( $conditional->getHashCode() ) == '') {
				//	continue;
				// }
				if ($conditional->getConditionType() != PHPExcel_Style_Conditional::CONDITION_NONE) 
				{
				    
					$cellReference = $conditional->getCellReference();
					// only process all rules for each cellReference once
					if(!in_array(str_replace(':','_',$cellReference),$processedCellReferences))
					{
						// conditionalFormatting
						if (($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CELLIS) ||
							($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT) ||
							($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_EXPRESSION) ||
							($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_TIMEPERIOD) ||
							($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_ABOVEAVERAGE) ||
							($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_TOP10) ||
							($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DUPLICATEVALUES)
							)
						{
							$cellReference = $conditional->getCellReference();
							$objWriter->startElement('conditionalFormatting');
								//$objWriter->writeAttribute('sqref',	$cellCoordinate);
								$objWriter->writeAttribute('sqref',	$cellReference);

									// cfRule
									$objWriter->startElement('cfRule');
									$objWriter->writeAttribute('type',		$conditional->getConditionType());
									$objWriter->writeAttribute('dxfId',		$this->getParentWriter()->getStylesConditionalHashTable()->getIndexForHashCode( $conditional->getHashCode() ));
									//$objWriter->writeAttribute('priority',	$id++);
									$objWriter->writeAttribute('priority',	$conditional->getPriority());
									
									if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_ABOVEAVERAGE)
									{
										$aboveAverage = $conditional->getAboveAverage();
										if (!is_null($aboveAverage)) { $objWriter->writeAttribute('aboveAverage', $aboveAverage); }
									}
									
									if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_TIMEPERIOD)
									{
										$timePeriod = $conditional->getTimePeriod();
										if (!is_null($timePeriod)) { $objWriter->writeAttribute('timePeriod', $timePeriod); }
									}
									
									if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_TOP10)
									{
										$percent = $conditional->getPercent();
										if (!is_null($percent)) { $objWriter->writeAttribute('percent',	$percent); }
										$bottom = $conditional->getBottom();
										if (!is_null($bottom)) { $objWriter->writeAttribute('bottom',	$bottom); }
										$rank = $conditional->getRank();
										if (!is_null($rank)) { $objWriter->writeAttribute('rank',	$rank); }

									}

									if (($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CELLIS
											||
										 $conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT)
										&& $conditional->getOperatorType() != PHPExcel_Style_Conditional::OPERATOR_NONE) {
										$objWriter->writeAttribute('operator',	$conditional->getOperatorType());
									}

									if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT
										&& !is_null($conditional->getText())) {
										$objWriter->writeAttribute('text',	$conditional->getText());
									}

									if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT
										&& $conditional->getOperatorType() == PHPExcel_Style_Conditional::OPERATOR_CONTAINSTEXT
										&& !is_null($conditional->getText())) {
										$objWriter->writeElement('formula',	'NOT(ISERROR(SEARCH("' . $conditional->getText() . '",' . $cellCoordinate . ')))');
									} else if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT
										&& $conditional->getOperatorType() == PHPExcel_Style_Conditional::OPERATOR_BEGINSWITH
										&& !is_null($conditional->getText())) {
										$objWriter->writeElement('formula',	'LEFT(' . $cellCoordinate . ',' . strlen($conditional->getText()) . ')="' . $conditional->getText() . '"');
									} else if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT
										&& $conditional->getOperatorType() == PHPExcel_Style_Conditional::OPERATOR_ENDSWITH
										&& !is_null($conditional->getText())) {
										$objWriter->writeElement('formula',	'RIGHT(' . $cellCoordinate . ',' . strlen($conditional->getText()) . ')="' . $conditional->getText() . '"');
									} else if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT
										&& $conditional->getOperatorType() == PHPExcel_Style_Conditional::OPERATOR_NOTCONTAINS
										&& !is_null($conditional->getText())) {
										$objWriter->writeElement('formula',	'ISERROR(SEARCH("' . $conditional->getText() . '",' . $cellCoordinate . '))');
									} else if ($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CELLIS
										|| $conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_CONTAINSTEXT
										|| $conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_TIMEPERIOD
										|| $conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_EXPRESSION) {
										foreach ($conditional->getConditions() as $formula) {
											// Formula
											$objWriter->writeElement('formula',	$formula);
										}
									}

								$objWriter->endElement();
							$objWriter->endElement();
						}
						elseif (($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_DATABAR) ||
								($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_COLORSCALE) ||
								($conditional->getConditionType() == PHPExcel_Style_Conditional::CONDITION_ICONSET))
						{
							// insert the element for the entire group
							$conditionalObj = $conditional->getConditionalObject();
							
					
							// databars/colorscales/iconsets are applied to a group of cells but their definition is just stored in one of the cells of the worksheet.
							// to prevent 'A1' from clogging up with all definition, the definition for each group is
							// assigned to the first cell of that group we envounter

							$defaultdata = $conditionalObj->getDefaultData();
							$classid = $conditionalObj->getClassID();
							$priority = $conditional->getPriority();
							
							if ($defaultdata)
							{
								// definition needs to be assigned to this cell
								// conditionalFormatting element
								$objWriter->startElement('conditionalFormatting');
								
								$realCellReference = $cellReference;
								$objWriter->writeAttribute('sqref',	$cellReference);
									// cfRule  element
									$objWriter->startElement('cfRule');
									$objWriter->writeAttribute('type', $conditional->getConditionType());
									$objWriter->writeAttribute('priority',	$priority);//$id++);
										
										// write out the dataBar/ColorScale/IconSet element with the default properties
										// some iconsets do not have default data (e.g custom icon sets)
										$this->_writeElement($objWriter, '' , $defaultdata);
										
										// check whether we need to add something to the extLst list
										if ($conditionalObj->needsExtLstEntry())
										{
											if ($conditionalObj->needsExtLstReference())
											{
												// add an extlst link for this element (only dataBar so far)
												$worksheet_cf_ns_id = 'x'.PHPExcel_Writer_Excel2007_Worksheet::EXTLST_CONDITIONALFORMATTINGID; 
												$objWriter->startElement('extLst');
													$objWriter->startElement('ext');
														$objWriter->writeAttribute('uri', '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'); //{B025F937-C7B1-47D3-B67F-A62EFF666E3E} = ext uri id : http://msdn.microsoft.com/en-us/library/dd905242%28v=office.12%29.aspx
														$objuri = $conditionalObj->getNamespace();
														$objWriter->writeAttribute('xmlns:'.$worksheet_cf_ns_id, $objuri);
														$objWriter->writeElement($worksheet_cf_ns_id.':id',$classid);
													$objWriter->endElement();
												$objWriter->endElement();
											}
										}
									$objWriter->endElement();
								$objWriter->endElement();
							}
							
							if ($conditionalObj->needsExtLstEntry())
							{
									// add an entry to the extlst list (to be written at the end of the worksheet by _writeExtLstEntries)
									$data = $conditionalObj->getExtLstData($priority);
									$this->addEntryToExtLstArray(PHPExcel_Writer_Excel2007_Worksheet::EXTLST_CONDITIONALFORMATTINGID, $cellReference, $classid, $data);
							}
						}
						
						// mark this cellReference as being processed
						$processedCellReferences[] = str_replace(':','_',$cellReference);
					}
				}
			}
		}
	}
	
	/**
	 * Write extLst entries
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeExtLstEntries(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// write all data in the extlst array
		$extlstarray = $this->getExtLstArray();
		if ($extlstarray)
		{
			// start marker
			$objWriter->startElement('extLst');
			
			// write it!
			foreach($extlstarray as $groupid => $groupdata)
			{
				switch($groupid)
				{
					case PHPExcel_Writer_Excel2007_Worksheet::EXTLST_CONDITIONALFORMATTINGID:
								$this->_writeExtLstConditionalFormattings($groupid, $groupdata, $objWriter, $pSheet);
							break;
					default:
								throw new PHPExcel_Writer_Exception("Unknown group .");
							
				}
			}
			
			// end marker
			$objWriter->endElement();
		}
	}
	
	private function _writeExtLstConditionalFormattings($groupid, $groupdata, PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		$worksheet_cf_ns_id = 'x'.$groupid;

		$objWriter->startElement('ext');
			$objWriter->writeAttribute('uri', "{78C0D931-6437-407d-A8EE-F0AAD7539E65}");
			$objWriter->writeAttribute('xmlns:'.$worksheet_cf_ns_id, "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
				$objWriter->startElement($worksheet_cf_ns_id.':conditionalFormattings');

					// write all rules for each cell group
					foreach($groupdata as $cellref => $elementdata)
					{
						$objWriter->startElement($worksheet_cf_ns_id.':conditionalFormatting');
							$objWriter->writeAttribute('xmlns:xm', "http://schemas.microsoft.com/office/excel/2006/main");
							// write each cfrule
							foreach($elementdata as $ruleid => $cfrule)
							{
								$this->_writeElement($objWriter, $worksheet_cf_ns_id ,$cfrule);
							}
							// write cell reference 
							$objWriter->writeElement('xm:sqref',str_replace('_',':',$cellref));
						$objWriter->endElement();
					}
					
				$objWriter->endElement();
		$objWriter->endElement();
		
		
	}
	
	private function _writeElement($objWriter, $prefix , $elementdata)
	{
		if (isset($elementdata['name']) && (isset($elementdata['attributes'])))
		{
			// if the element has a namespace defined, use it instead of the default namespace
			if (isset($elementdata['namespace'])) 
			{ 
				$prefix= $elementdata['namespace']; 
			} 
			// generate the name (+ namespace prefix) for this elemet
			$elementname = ($prefix=="")?$elementdata['name']:$prefix.':'.$elementdata['name'];

			if (isset($elementdata['attributes']['value']))
			{
				// single valued element e.g. <xm:f>5</xm:f>				
				$objWriter->writeElement($elementname, $elementdata['attributes']['value']);
			}
			else
			{
				// start this element
				$objWriter->startElement($elementname);
				
				// add attributes (e.g type in <cfvo type='3'>)
				foreach ($elementdata['attributes'] as $attributeproperties)
				{
					if (!is_array($attributeproperties['attributes']))
					{
						// attribute for this element
						$objWriter->writeAttribute($attributeproperties['name'], (string)$attributeproperties['attributes']);
					}
				}
				
				// write nested elements (e.g color element in <cfvo><color rgb="FFFFFF00"></cfvo>)
				foreach ($elementdata['attributes'] as $attributeproperties)
				{
					if (is_array($attributeproperties['attributes']))
					{
						// nested element
						$this->_writeElement($objWriter, $prefix, $attributeproperties);
					}
				}
				// close this element
				$objWriter->endElement();
			}
		}
		else
		{
			throw new PHPExcel_Writer_Exception("_writeElement : missing name or attributes property:".var_export($elementdata,true));
		}
	}
	
	/**
	 * Write DataValidations
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeDataValidations(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// Datavalidation collection
		$dataValidationCollection = $pSheet->getDataValidationCollection();

		// Write data validations?
		if (!empty($dataValidationCollection)) {
			$objWriter->startElement('dataValidations');
			$objWriter->writeAttribute('count', count($dataValidationCollection));

			foreach ($dataValidationCollection as $coordinate => $dv) {
				$objWriter->startElement('dataValidation');

				if ($dv->getType() != '') {
					$objWriter->writeAttribute('type', $dv->getType());
				}

				if ($dv->getErrorStyle() != '') {
					$objWriter->writeAttribute('errorStyle', $dv->getErrorStyle());
				}

				if ($dv->getOperator() != '') {
					$objWriter->writeAttribute('operator', $dv->getOperator());
				}

				$objWriter->writeAttribute('allowBlank',		($dv->getAllowBlank()		? '1'  : '0'));
				$objWriter->writeAttribute('showDropDown',		(!$dv->getShowDropDown()	? '1'  : '0'));
				$objWriter->writeAttribute('showInputMessage',	($dv->getShowInputMessage()	? '1'  : '0'));
				$objWriter->writeAttribute('showErrorMessage',	($dv->getShowErrorMessage()	? '1'  : '0'));

				if ($dv->getErrorTitle() !== '') {
					$objWriter->writeAttribute('errorTitle', $dv->getErrorTitle());
				}
				if ($dv->getError() !== '') {
					$objWriter->writeAttribute('error', $dv->getError());
				}
				if ($dv->getPromptTitle() !== '') {
					$objWriter->writeAttribute('promptTitle', $dv->getPromptTitle());
				}
				if ($dv->getPrompt() !== '') {
					$objWriter->writeAttribute('prompt', $dv->getPrompt());
				}

				$objWriter->writeAttribute('sqref', $coordinate);

				if ($dv->getFormula1() !== '') {
					$objWriter->writeElement('formula1', $dv->getFormula1());
				}
				if ($dv->getFormula2() !== '') {
					$objWriter->writeElement('formula2', $dv->getFormula2());
				}

				$objWriter->endElement();
			}

			$objWriter->endElement();
		}
	}

	/**
	 * Write Hyperlinks
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeHyperlinks(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// Hyperlink collection
		$hyperlinkCollection = $pSheet->getHyperlinkCollection();

		// Relation ID
		$relationId = 1;

		// Write hyperlinks?
		if (!empty($hyperlinkCollection)) {
			$objWriter->startElement('hyperlinks');

			foreach ($hyperlinkCollection as $coordinate => $hyperlink) {
				$objWriter->startElement('hyperlink');

				$objWriter->writeAttribute('ref', $coordinate);
				if (!$hyperlink->isInternal()) {
					$objWriter->writeAttribute('r:id',	'rId_hyperlink_' . $relationId);
					++$relationId;
				} else {
					$objWriter->writeAttribute('location',	str_replace('sheet://', '', $hyperlink->getUrl()));
				}

				if ($hyperlink->getTooltip() != '') {
					$objWriter->writeAttribute('tooltip', $hyperlink->getTooltip());
				}

				$objWriter->endElement();
			}

			$objWriter->endElement();
		}
	}

	/**
	 * Write ProtectedRanges
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeProtectedRanges(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		if (count($pSheet->getProtectedCells()) > 0) {
			// protectedRanges
			$objWriter->startElement('protectedRanges');

				// Loop protectedRanges
				foreach ($pSheet->getProtectedCells() as $protectedCell => $passwordHash) {
					// protectedRange
					$objWriter->startElement('protectedRange');
					$objWriter->writeAttribute('name',		'p' . md5($protectedCell));
					$objWriter->writeAttribute('sqref',	$protectedCell);
					if (!empty($passwordHash)) {
						$objWriter->writeAttribute('password',	$passwordHash);
					}
					$objWriter->endElement();
				}

			$objWriter->endElement();
		}
	}

	/**
	 * Write MergeCells
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeMergeCells(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		if (count($pSheet->getMergeCells()) > 0) {
			// mergeCells
			$objWriter->startElement('mergeCells');

				// Loop mergeCells
				foreach ($pSheet->getMergeCells() as $mergeCell) {
					// mergeCell
					$objWriter->startElement('mergeCell');
					$objWriter->writeAttribute('ref', $mergeCell);
					$objWriter->endElement();
				}

			$objWriter->endElement();
		}
	}

	/**
	 * Write PrintOptions
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writePrintOptions(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// printOptions
		$objWriter->startElement('printOptions');

		$objWriter->writeAttribute('gridLines',	($pSheet->getPrintGridlines() ? 'true': 'false'));
		$objWriter->writeAttribute('gridLinesSet',	'true');

		if ($pSheet->getPageSetup()->getHorizontalCentered()) {
			$objWriter->writeAttribute('horizontalCentered', 'true');
		}

		if ($pSheet->getPageSetup()->getVerticalCentered()) {
			$objWriter->writeAttribute('verticalCentered', 'true');
		}

		$objWriter->endElement();
	}

	/**
	 * Write PageMargins
	 *
	 * @param	PHPExcel_Shared_XMLWriter				$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet						$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writePageMargins(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// pageMargins
		$objWriter->startElement('pageMargins');
		$objWriter->writeAttribute('left',		PHPExcel_Shared_String::FormatNumber($pSheet->getPageMargins()->getLeft()));
		$objWriter->writeAttribute('right',		PHPExcel_Shared_String::FormatNumber($pSheet->getPageMargins()->getRight()));
		$objWriter->writeAttribute('top',		PHPExcel_Shared_String::FormatNumber($pSheet->getPageMargins()->getTop()));
		$objWriter->writeAttribute('bottom',	PHPExcel_Shared_String::FormatNumber($pSheet->getPageMargins()->getBottom()));
		$objWriter->writeAttribute('header',	PHPExcel_Shared_String::FormatNumber($pSheet->getPageMargins()->getHeader()));
		$objWriter->writeAttribute('footer',	PHPExcel_Shared_String::FormatNumber($pSheet->getPageMargins()->getFooter()));
		$objWriter->endElement();
	}

	/**
	 * Write AutoFilter
	 *
	 * @param	PHPExcel_Shared_XMLWriter				$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet						$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeAutoFilter(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		$autoFilterRange = $pSheet->getAutoFilter()->getRange();
		if (!empty($autoFilterRange)) {
			// autoFilter
			$objWriter->startElement('autoFilter');

			// Strip any worksheet reference from the filter coordinates
			$range = PHPExcel_Cell::splitRange($autoFilterRange);
			$range = $range[0];
			//	Strip any worksheet ref
			if (strpos($range[0],'!') !== false) {
				list($ws,$range[0]) = explode('!',$range[0]);
			}
			$range = implode(':', $range);

			$objWriter->writeAttribute('ref',	str_replace('$','',$range));

			$columns = $pSheet->getAutoFilter()->getColumns();
			if (count($columns) > 0) {
				foreach($columns as $columnID => $column) {
					$rules = $column->getRules();
					if (count($rules > 0)) {
						$objWriter->startElement('filterColumn');
							$objWriter->writeAttribute('colId',	$pSheet->getAutoFilter()->getColumnOffset($columnID));

							$objWriter->startElement( $column->getFilterType());
								if ($column->getJoin() == PHPExcel_Worksheet_AutoFilter_Column::AUTOFILTER_COLUMN_JOIN_AND) {
									$objWriter->writeAttribute('and',	1);
								}

								foreach ($rules as $rule) {
									if (($column->getFilterType() === PHPExcel_Worksheet_AutoFilter_Column::AUTOFILTER_FILTERTYPE_FILTER) &&
										($rule->getOperator() === PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_COLUMN_RULE_EQUAL) &&
										($rule->getValue() === '')) {
										//	Filter rule for Blanks
										$objWriter->writeAttribute('blank',	1);
									} elseif($rule->getRuleType() === PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_RULETYPE_DYNAMICFILTER) {
										//	Dynamic Filter Rule
										$objWriter->writeAttribute('type', $rule->getGrouping());
										$val = $column->getAttribute('val');
										if ($val !== NULL) {
											$objWriter->writeAttribute('val', $val);
										}
										$maxVal = $column->getAttribute('maxVal');
										if ($maxVal !== NULL) {
											$objWriter->writeAttribute('maxVal', $maxVal);
										}
									} elseif($rule->getRuleType() === PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_RULETYPE_TOPTENFILTER) {
										//	Top 10 Filter Rule
										$objWriter->writeAttribute('val',	$rule->getValue());
										$objWriter->writeAttribute('percent',	(($rule->getOperator() === PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_PERCENT) ? '1' : '0'));
										$objWriter->writeAttribute('top',	(($rule->getGrouping() === PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_COLUMN_RULE_TOPTEN_TOP) ? '1': '0'));
									} else {
										//	Filter, DateGroupItem or CustomFilter
										$objWriter->startElement($rule->getRuleType());

											if ($rule->getOperator() !== PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_COLUMN_RULE_EQUAL) {
												$objWriter->writeAttribute('operator',	$rule->getOperator());
											}
											if ($rule->getRuleType() === PHPExcel_Worksheet_AutoFilter_Column_Rule::AUTOFILTER_RULETYPE_DATEGROUP) {
												// Date Group filters
												foreach($rule->getValue() as $key => $value) {
													if ($value > '') $objWriter->writeAttribute($key,	$value);
												}
												$objWriter->writeAttribute('dateTimeGrouping',	$rule->getGrouping());
											} else {
												$objWriter->writeAttribute('val',	$rule->getValue());
											}

										$objWriter->endElement();
									}
								}

							$objWriter->endElement();

						$objWriter->endElement();
					}
				}
			}

			$objWriter->endElement();
		}
	}

	/**
	 * Write PageSetup
	 *
	 * @param	PHPExcel_Shared_XMLWriter			$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet					$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writePageSetup(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// pageSetup
		$objWriter->startElement('pageSetup');
		$objWriter->writeAttribute('paperSize',		$pSheet->getPageSetup()->getPaperSize());
		$objWriter->writeAttribute('orientation',	$pSheet->getPageSetup()->getOrientation());

		if ((!is_null($pSheet->getPageSetup()->getScale())) && ($pSheet->getPageSetup()->getScale() != 100)) {
			// don't write out the scale if it is 100%
			$objWriter->writeAttribute('scale',				 $pSheet->getPageSetup()->getScale());
		}
		if (!is_null($pSheet->getPageSetup()->getFitToHeight())) {
			$objWriter->writeAttribute('fitToHeight',		 $pSheet->getPageSetup()->getFitToHeight());
		} else {
			$objWriter->writeAttribute('fitToHeight',		 '0');
		}
		if (!is_null($pSheet->getPageSetup()->getFitToWidth())) {
			$objWriter->writeAttribute('fitToWidth',		 $pSheet->getPageSetup()->getFitToWidth());
		} else {
			$objWriter->writeAttribute('fitToWidth',		 '0');
		}
		if (!is_null($pSheet->getPageSetup()->getFirstPageNumber())) {
			$objWriter->writeAttribute('firstPageNumber',	$pSheet->getPageSetup()->getFirstPageNumber());
			$objWriter->writeAttribute('useFirstPageNumber', '1');
		}
		if (!is_null($pSheet->getPageSetup()->getHorizontalDpi())) {
			$objWriter->writeAttribute('horizontalDpi',		 $pSheet->getPageSetup()->getHorizontalDpi());
		}
		if (!is_null($pSheet->getPageSetup()->getVerticalDpi())) {
			$objWriter->writeAttribute('verticalDpi',		 $pSheet->getPageSetup()->getVerticalDpi());
		}
	
		$objWriter->endElement();
	}

	/**
	 * Write Header / Footer
	 *
	 * @param	PHPExcel_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet				$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeHeaderFooter(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// headerFooter
		$oddHeader = $pSheet->getHeaderFooter()->getOddHeader();
		$oddFooter = $pSheet->getHeaderFooter()->getOddFooter();
		$evenHeader = $pSheet->getHeaderFooter()->getEvenHeader();
		$evenFooter = $pSheet->getHeaderFooter()->getEvenFooter();
		$firstHeader = $pSheet->getHeaderFooter()->getFirstHeader();
		$firstFooter = $pSheet->getHeaderFooter()->getFirstFooter();
		// check if any footer/header information needs to be written
		if ($oddHeader.$oddFooter.$evenHeader.$evenFooter.$firstHeader.$firstFooter != '')
		{
			$objWriter->startElement('headerFooter');
			$objWriter->writeAttribute('differentOddEven',	($pSheet->getHeaderFooter()->getDifferentOddEven() ? 'true' : 'false'));
			$objWriter->writeAttribute('differentFirst',	($pSheet->getHeaderFooter()->getDifferentFirst() ? 'true' : 'false'));
			$objWriter->writeAttribute('scaleWithDoc',		($pSheet->getHeaderFooter()->getScaleWithDocument() ? 'true' : 'false'));
			$objWriter->writeAttribute('alignWithMargins',	($pSheet->getHeaderFooter()->getAlignWithMargins() ? 'true' : 'false'));
				// only write footer/headers that are defined
				if ($oddHeader != '') { $objWriter->writeElement('oddHeader', $oddHeader); }
				if ($oddFooter != '') { $objWriter->writeElement('oddFooter', $oddFooter); }
				if ($evenHeader != '') { $objWriter->writeElement('evenHeader', $evenHeader); }
				if ($evenFooter != '') { $objWriter->writeElement('evenFooter',	$evenFooter); }
				if ($firstHeader != '') { $objWriter->writeElement('firstHeader', $firstHeader); } //check : oddHeader/evenHeader = 'XX' + firstHeader is blank
				if ($firstFooter != '') { $objWriter->writeElement('firstFooter',	$firstFooter); }
			$objWriter->endElement();
		}
	}

	/**
	 * Write Breaks
	 *
	 * @param	PHPExcel_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet				$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeBreaks(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// Get row and column breaks
		$aRowBreaks = array();
		$aColumnBreaks = array();
		foreach ($pSheet->getBreaks() as $cell => $breakType) {
			if ($breakType == PHPExcel_Worksheet::BREAK_ROW) {
				$aRowBreaks[] = $cell;
			} else if ($breakType == PHPExcel_Worksheet::BREAK_COLUMN) {
				$aColumnBreaks[] = $cell;
			}
		}

		// rowBreaks
		if (!empty($aRowBreaks)) {
			$objWriter->startElement('rowBreaks');
			$objWriter->writeAttribute('count',			count($aRowBreaks));
			$objWriter->writeAttribute('manualBreakCount',	count($aRowBreaks));

				foreach ($aRowBreaks as $cell) {
					$coords = PHPExcel_Cell::coordinateFromString($cell);

					$objWriter->startElement('brk');
					$objWriter->writeAttribute('id',	$coords[1]);
					$objWriter->writeAttribute('man',	'1');
					$objWriter->endElement();
				}

			$objWriter->endElement();
		}

		// Second, write column breaks
		if (!empty($aColumnBreaks)) {
			$objWriter->startElement('colBreaks');
			$objWriter->writeAttribute('count',			count($aColumnBreaks));
			$objWriter->writeAttribute('manualBreakCount',	count($aColumnBreaks));

				foreach ($aColumnBreaks as $cell) {
					$coords = PHPExcel_Cell::coordinateFromString($cell);

					$objWriter->startElement('brk');
					$objWriter->writeAttribute('id',	PHPExcel_Cell::columnIndexFromString($coords[0]) - 1);
					$objWriter->writeAttribute('man',	'1');
					$objWriter->endElement();
				}

			$objWriter->endElement();
		}
	}

	/**
	 * Write SheetData
	 *
	 * @param	PHPExcel_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet				$pSheet			Worksheet
	 * @param	string[]						$pStringTable	String table
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeSheetData(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null, $pStringTable = null)
	{
		if (is_array($pStringTable)) {
			// Flipped stringtable, for faster index searching
			$aFlippedStringTable = $this->getParentWriter()->getWriterPart('stringtable')->flipStringTable($pStringTable);

			// sheetData
			$objWriter->startElement('sheetData');

				// Get column count
				$colCount = PHPExcel_Cell::columnIndexFromString($pSheet->getHighestColumn());

				// Highest row number
				$highestRow = $pSheet->getHighestRow();

				// Loop through cells
				$cellsByRow = array();
				foreach ($pSheet->getCellCollection() as $cellID) {
					$cellAddress = PHPExcel_Cell::coordinateFromString($cellID);
					$cellsByRow[$cellAddress[1]][] = $cellID;
				}

				$currentRow = 0;
				while($currentRow++ < $highestRow) {
					// Get row dimension
					$rowDimension = $pSheet->getRowDimension($currentRow);

					// Write current row?
					$writeCurrentRow =	isset($cellsByRow[$currentRow]) ||
										$rowDimension->getRowHeight() >= 0 ||
										$rowDimension->getVisible() == false ||
										$rowDimension->getCollapsed() == true ||
										$rowDimension->getOutlineLevel() > 0 ||
										$rowDimension->getXfIndex() !== null;

					if ($writeCurrentRow) {
						// Start a new row
						$objWriter->startElement('row');
						$objWriter->writeAttribute('r',	$currentRow);
						$objWriter->writeAttribute('spans',	'1:' . $colCount);

						// Row dimensions
						if ($rowDimension->getRowHeight() >= 0) {
							$objWriter->writeAttribute('customHeight',	'1');
							$objWriter->writeAttribute('ht',			PHPExcel_Shared_String::FormatNumber($rowDimension->getRowHeight()));
						}

						// Row visibility
						if ($rowDimension->getVisible() == false) {
							$objWriter->writeAttribute('hidden',		'true');
						}

						// Collapsed
						if ($rowDimension->getCollapsed() == true) {
							$objWriter->writeAttribute('collapsed',		'true');
						}

						// Outline level
						if ($rowDimension->getOutlineLevel() > 0) {
							$objWriter->writeAttribute('outlineLevel',	$rowDimension->getOutlineLevel());
						}

						// Style
						if ($rowDimension->getXfIndex() !== null) {
							$objWriter->writeAttribute('s',	$rowDimension->getXfIndex());
							$objWriter->writeAttribute('customFormat', '1');
						}

						// Write cells
						if (isset($cellsByRow[$currentRow])) {
							foreach($cellsByRow[$currentRow] as $cellAddress) {
								// Write cell
								$this->_writeCell($objWriter, $pSheet, $cellAddress, $pStringTable, $aFlippedStringTable);
							}
						}

						// End row
						$objWriter->endElement();
					}
				}

			$objWriter->endElement();
		} else {
			throw new PHPExcel_Writer_Exception("Invalid parameters passed.");
		}
	}

	/**
	 * Write Cell
	 *
	 * @param	PHPExcel_Shared_XMLWriter	$objWriter				XML Writer
	 * @param	PHPExcel_Worksheet			$pSheet					Worksheet
	 * @param	PHPExcel_Cell				$pCellAddress			Cell Address
	 * @param	string[]					$pStringTable			String table
	 * @param	string[]					$pFlippedStringTable	String table (flipped), for faster index searching
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeCell(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null, $pCellAddress = null, $pStringTable = null, $pFlippedStringTable = null)
	{
		if (is_array($pStringTable) && is_array($pFlippedStringTable)) {
			// Cell
			$pCell = $pSheet->getCell($pCellAddress);
			$objWriter->startElement('c');
			$objWriter->writeAttribute('r', $pCellAddress);

			// Sheet styles
			if ($pCell->getXfIndex() != '') {
				$objWriter->writeAttribute('s', $pCell->getXfIndex());
			}

			// If cell value is supplied, write cell value
			$cellValue = $pCell->getValue();
			if (is_object($cellValue) || $cellValue !== '') {
				// Map type
				$mappedType = $pCell->getDataType();

				// Write data type depending on its type
				switch (strtolower($mappedType)) {
					case 'inlinestr':	// Inline string
					case 's':			// String
					case 'b':			// Boolean
						$objWriter->writeAttribute('t', $mappedType);
						break;
					case 'f':			// Formula
						$calculatedValue = ($this->getParentWriter()->getPreCalculateFormulas()) ?
						    $pCell->getCalculatedValue() :
						    $cellValue;
						if (is_string($calculatedValue)) {
							$objWriter->writeAttribute('t', 'str');
						}
						break;
					case 'e':			// Error
						$objWriter->writeAttribute('t', $mappedType);
				}

				// Write data depending on its type
				switch (strtolower($mappedType)) {
					case 'inlinestr':	// Inline string
						if (! $cellValue instanceof PHPExcel_RichText) {
							$objWriter->writeElement('t', PHPExcel_Shared_String::ControlCharacterPHP2OOXML( htmlspecialchars($cellValue) ) );
						} else if ($cellValue instanceof PHPExcel_RichText) {
							$objWriter->startElement('is');
							$this->getParentWriter()->getWriterPart('stringtable')->writeRichText($objWriter, $cellValue);
							$objWriter->endElement();
						}

						break;
					case 's':			// String
						if (! $cellValue instanceof PHPExcel_RichText) {
							if (isset($pFlippedStringTable[$cellValue])) {
								$objWriter->writeElement('v', $pFlippedStringTable[$cellValue]);
							}
						} else if ($cellValue instanceof PHPExcel_RichText) {
							$objWriter->writeElement('v', $pFlippedStringTable[$cellValue->getHashCode()]);
						}

						break;
					case 'f':			// Formula
						$attributes = $pCell->getFormulaAttributes();
						if($attributes['t'] == 'array') {
							$objWriter->startElement('f');
							$objWriter->writeAttribute('t', 'array');
							$objWriter->writeAttribute('ref', $pCellAddress);
							$objWriter->writeAttribute('aca', '1');
							$objWriter->writeAttribute('ca', '1');
							$objWriter->text(substr($cellValue, 1));
							$objWriter->endElement();
						} else {
							$objWriter->writeElement('f', substr($cellValue, 1));
						}
						if ($this->getParentWriter()->getOffice2003Compatibility() === false) {
							if ($this->getParentWriter()->getPreCalculateFormulas()) {
//								$calculatedValue = $pCell->getCalculatedValue();
								if (!is_array($calculatedValue) && substr($calculatedValue, 0, 1) != '#') {
									$objWriter->writeElement('v', PHPExcel_Shared_String::FormatNumber($calculatedValue));
								} else {
									$objWriter->writeElement('v', '0');
								}
							} else {
								$objWriter->writeElement('v', '0');
							}
						}
						break;
					case 'n':			// Numeric
						// force point as decimal separator in case current locale uses comma
						$objWriter->writeElement('v', str_replace(',', '.', $cellValue));
						break;
					case 'b':			// Boolean
						$objWriter->writeElement('v', ($cellValue ? '1' : '0'));
						break;
					case 'e':			// Error
						if (substr($cellValue, 0, 1) == '=') {
							$objWriter->writeElement('f', substr($cellValue, 1));
							$objWriter->writeElement('v', substr($cellValue, 1));
						} else {
							$objWriter->writeElement('v', $cellValue);
						}

						break;
				}
			}

			$objWriter->endElement();
		} else {
			throw new PHPExcel_Writer_Exception("Invalid parameters passed.");
		}
	}

	/**
	 * Write Drawings
	 *
	 * @param	PHPExcel_Shared_XMLWriter	$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet			$pSheet			Worksheet
	 * @param	boolean						$includeCharts	Flag indicating if we should include drawing details for charts
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeDrawings(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null, $includeCharts = FALSE)
	{
		$chartCount = ($includeCharts) ? $pSheet->getChartCollection()->count() : 0;
		// If sheet contains drawings, add the relationships
		if (($pSheet->getDrawingCollection()->count() > 0) ||
			($chartCount > 0)) {
			$objWriter->startElement('drawing');
			$objWriter->writeAttribute('r:id', 'rId1');
			$objWriter->endElement();
		}
	}

	/**
	 * Write LegacyDrawing
	 *
	 * @param	PHPExcel_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet				$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeLegacyDrawing(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// If sheet contains comments, add the relationships
		if (count($pSheet->getComments()) > 0) {
			$objWriter->startElement('legacyDrawing');
			$objWriter->writeAttribute('r:id', 'rId_comments_vml1');
			$objWriter->endElement();
		}
	}

	/**
	 * Write LegacyDrawingHF
	 *
	 * @param	PHPExcel_Shared_XMLWriter		$objWriter		XML Writer
	 * @param	PHPExcel_Worksheet				$pSheet			Worksheet
	 * @throws	PHPExcel_Writer_Exception
	 */
	private function _writeLegacyDrawingHF(PHPExcel_Shared_XMLWriter $objWriter = null, PHPExcel_Worksheet $pSheet = null)
	{
		// If sheet contains images, add the relationships
		if (count($pSheet->getHeaderFooter()->getImages()) > 0) {
			$objWriter->startElement('legacyDrawingHF');
			$objWriter->writeAttribute('r:id', 'rId_headerfooter_vml1');
			$objWriter->endElement();
		}
	}
	
		/** 
	 * add an entry to the extlst array
	 * @param	string	groupid (e.g. conditionalformattings)
	 * @param	string	id of the data to be added
	 * @param	array	data to be added
	 * @throws PHPExcel_Exception
	*/
	protected function addEntryToExtLstArray($groupid, $cellReference, $id, $data)
	{
		$ref = str_replace(':','_',$cellReference);
		if (!$this->_extlst) { $this->_extlst = array(); }
		if (isset($this->_extlst[$groupid][$ref][$id]))
		{
			throw new PHPExcel_Exception('Unhandled situation : CLASSID '.$id.' is already used in the extlst list');
		}
		
		if (!isset($this->_extlst[$groupid])) { $this->_extlst[$groupid] = array(); }
		if (!isset($this->_extlst[$groupid][$ref])) { $this->_extlst[$groupid][$ref] = array(); }
		
		$this->_extlst[$groupid][$ref][$id] = $data;
	}
	
	/**
	 * return the extLst array
	 *
	 * @returns	array	extLst array
	 */
	protected function getExtLstArray()
	{
		return $this->_extlst;
	}
	
	/**
	 * Clear the extLst
	 *
	 */
	protected function clearExtLst()
	{
		$this->_extLst = null;
	}

}
