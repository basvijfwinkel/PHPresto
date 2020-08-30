<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

date_default_timezone_set('Europe/London');
require_once dirname(__FILE__) . '/../Classes/PHPExcel.php';


$objPHPExcel = new PHPExcel();
$objWorksheet = $objPHPExcel->getActiveSheet();


$objWorksheet->fromArray(
    array(
    array('',    2010,    2011,    2012),
    array('Q1',   12,   15,        21),
    array('Q2',   56,   73,        86),
    array('Q3',   52,   61,        69),
    array('Q4',   30,   32,        1),
    )
);

//    Set the Labels for each data series we want to plot
//        Datatype
//        Cell reference for data
//        Format Code
//        Number of datapoints in series
//        Data values
//        Data Marker
$dataSeriesLabels = array(
    new PHPExcel_Chart_DataSeriesValues('String', $sheetname.'!$B$2:$B$5', NULL, 1),
);
//    Set the X-Axis Labels
$xAxisTickValues = array(
    new PHPExcel_Chart_DataSeriesValues('String', $sheetname.'!$A$2:$A$5', NULL, 4),
);
//    Set the Data values for each data series we want to plot
//        Datatype
//        Cell reference for data
//        Format Code
//        Number of datapoints in series
//        Data values
//        Data Marker
$dataSeriesValues = array(
    new PHPExcel_Chart_DataSeriesValues('Number', $sheetname.'!$B$2:$B$5', NULL, 4),
);


$dataSeriesSizes = array(
    new PHPExcel_Chart_DataSeriesValues('Number', $sheetname.'!$D$2:$D$5', NULL, 1),
);

// Bubble labels
$dataSeriesLabels  = array(new PHPExcel_Chart_DataSeriesValues('String', $sheetname.'!$A$2:$A$5', NULL, $rowcount));

// Bubble scale
$bubbleScale = 100;

//    Build the dataseries
$series = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART,   // plotType
    NULL,                                          // plotGrouping (Scatter charts don't have any grouping)
    range(0, count($dataSeriesValues)-1),          // plotOrder
    $dataSeriesLabels,                             // plotLabel
    $xAxisTickValues,                              // plotCategory
    $dataSeriesValues,                             // plotValues
    NULL,                                          // plotDirection
    NULL,                                          // smooth line
    PHPExcel_Chart_DataSeries::STYLE_LINEMARKER,   // plotStyle
    $dataSeriesSizes                               // BubbleChartSizea
    $dataSeriesLabels,                             // BubbleChartLabels
    $bubbleScale                                   // bubble scale
);


// locate the labels on the right
$layout =  new PHPExcel_Chart_Layout();
$layout->setDataLabelPosition(PHPExcel_Chart_Layout::LABEL_POS_RIGHT);
$layout->setShowVal(1);

//    Set the series in the plot area
$plotArea = new PHPExcel_Chart_PlotArea($layout, array($series));
//    Set the chart legend
$legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_TOPRIGHT, NULL, false);

$title = new PHPExcel_Chart_Title('Test Scatter Chart');
$yAxisLabel = new PHPExcel_Chart_Title('Value ($k)');

    $xAxis = new PHPExcel_Chart_Axis();
    $xAxis->setAxisOptionsProperty('hide_major_gridlines', true);
    $yAxis = new PHPExcel_Chart_Axis();
    $yAxis->setAxisOptionsProperty('hide_major_gridlines', true);

//    Create the chart
$chart = new PHPExcel_Chart(
    'chart1',        // name
    $title,            // title
    $legend,        // legend
    $plotArea,        // plotArea
    true,            // plotVisibleOnly
    0,                // displayBlanksAs
    NULL,            // xAxisLabel
    $yAxisLabel,        // yAxisLabel
    $xAxis,
    $yAxis
);

//    Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition('A7');
$chart->setBottomRightPosition('H20');

//    Add the chart to the worksheet
$objWorksheet->addChart($chart);

// Save Excel 2007 file
echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;


// Echo memory peak usage
echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;

// Echo done
echo date('H:i:s') , " Done writing file" , EOL;
echo 'File has been created in ' , getcwd() , EOL;
