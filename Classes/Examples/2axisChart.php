<?php
include_once './phpexcel/PHPExcel/Classes/PHPExcel.php';

$phpexcel = new PHPExcel();
$phpexcel->setActiveSheetIndex(0);
$sheet = $phpexcel->getActiveSheet();
$sheet->setTitle('Sheet1');
$format = 'm/d/yy';
// add data
$sheet->setCellValueByColumnAndRow(0, 1, 'chart');
$sheet->setCellValueByColumnAndRow(1, 1, 'series1');
$sheet->setCellValueByColumnAndRow(2, 1, 'series2');
$sheet->setCellValueByColumnAndRow(3, 1, 'series3');
$sheet->setCellValueByColumnAndRow(4, 1, 'series4');

$sheet->setCellValueByColumnAndRow(0, 2, PHPExcel_Shared_Date::PHPToExcel(new DateTime('2016-12-01')));
$sheet->setCellValueByColumnAndRow(1, 2, 11);
$sheet->setCellValueByColumnAndRow(2, 2, 22);
$sheet->setCellValueByColumnAndRow(3, 2, 333);
$sheet->setCellValueByColumnAndRow(4, 2, 444);
$sheet->getStyleByColumnAndRow(0, 2)->getNumberFormat()->setFormatCode($format);

$sheet->setCellValueByColumnAndRow(0, 3, PHPExcel_Shared_Date::PHPToExcel(new DateTime('2016-12-02')));
$sheet->setCellValueByColumnAndRow(1, 3, 12);
$sheet->setCellValueByColumnAndRow(2, 3, 23);
$sheet->setCellValueByColumnAndRow(3, 3, 334);
$sheet->setCellValueByColumnAndRow(4, 3, 445);
$sheet->getStyleByColumnAndRow(0, 3)->getNumberFormat()->setFormatCode($format);

$sheet->setCellValueByColumnAndRow(0, 4, PHPExcel_Shared_Date::PHPToExcel(new DateTime('2016-12-03')));
$sheet->setCellValueByColumnAndRow(1, 4, 13);
$sheet->setCellValueByColumnAndRow(2, 4, 24);
$sheet->setCellValueByColumnAndRow(3, 4, 335);
$sheet->setCellValueByColumnAndRow(4, 4, 446);
$sheet->getStyleByColumnAndRow(0, 4)->getNumberFormat()->setFormatCode($format);

$sheet->setCellValueByColumnAndRow(0, 5, PHPExcel_Shared_Date::PHPToExcel(new DateTime('2016-12-04')));
$sheet->setCellValueByColumnAndRow(1, 5, 14);
$sheet->setCellValueByColumnAndRow(2, 5, 25);
$sheet->setCellValueByColumnAndRow(3, 5, 336);
$sheet->setCellValueByColumnAndRow(4, 5, 447);
$sheet->getStyleByColumnAndRow(0, 5)->getNumberFormat()->setFormatCode($format);

//========= ADD CHART =============
$label1 = new PHPExcel_Chart_DataSeriesValues('String', "'Sheet1'!B1", null, 1);
$label2 = new PHPExcel_Chart_DataSeriesValues('String', "'Sheet1'!C1", null, 1);
$label3 = new PHPExcel_Chart_DataSeriesValues('String', "'Sheet1'!D1", null, 1);
$label4 = new PHPExcel_Chart_DataSeriesValues('String', "'Sheet1'!E1", null, 1);

$chrtCols  = "'Sheet1'!A2:A5";
$chrtVals  = "'Sheet1'!B2:B5";
$chrtVals2 = "'Sheet1'!C2:C5";
$chrtVals3 = "'Sheet1'!D2:D5";
$chrtVals4 = "'Sheet1'!E2:E5";

$periods = new PHPExcel_Chart_DataSeriesValues('String', $chrtCols, null, 4);
$values  = new PHPExcel_Chart_DataSeriesValues('Number', $chrtVals, null, 4);
$values2 = new PHPExcel_Chart_DataSeriesValues('Number', $chrtVals2, null,4);
$values3 = new PHPExcel_Chart_DataSeriesValues('Number', $chrtVals3, null,4);
$values4 = new PHPExcel_Chart_DataSeriesValues('Number', $chrtVals4, null,4);

$series1 = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_LINECHART,
    PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
    array(0,1),
    array($label1,$label2),
    array($periods,$periods),
    array($values,$values2)
);

$series2 = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_LINECHART,
    PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
    array(0,1),
    array($label3,$label4),
    array($periods,$periods),
    array($values3,$values4)
);


$series1->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
$layout = new PHPExcel_Chart_Layout();

$plotarea = new PHPExcel_Chart_PlotArea($layout, array($series1, $series2));
$chart = new PHPExcel_Chart('sample', null, null, $plotarea);

$secondaryYAxis = new PHPExcel_Chart_Axis();
$secondaryYAxis->setAxisOptionsProperty('vertical_axis_position', 'r');
$secondaryYAxis->setAxisOptionsProperty('orientation','minMax');
$chart->setSecondaryYAxis($secondaryYAxis);
$secondaryXAxis = new PHPExcel_Chart_Axis();
$secondaryXAxis->setAxisOptionsProperty('horizontal_crosses','max');
$chart->setSecondaryXAxis($secondaryXAxis);

$chart->setTopLeftPosition('A7');
$chart->setBottomRightPosition('F20');
$sheet->addChart($chart);

//========= OUTPUT RESULT =========
$writer = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel2007');
$writer->setIncludeCharts(TRUE);
$writer->save('chart.xlsx');

?>