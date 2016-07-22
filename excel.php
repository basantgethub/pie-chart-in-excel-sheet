<?php
/** Error reporting 
//error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');
date_default_timezone_set('Europe/London');*/

//set_include_path(get_include_path() . PATH_SEPARATOR . 'excelPHP/phpexcel-master/Classes/PHPExcel.php');
/** PHPExcel */
include 'excelPHP/Classes/PHPExcel.php';
$objPHPExcel = new PHPExcel();
$objWorksheet = $objPHPExcel->getActiveSheet();
$objWorksheet->fromArray(
	array(
		array('',	2010,	2011,	2012),
		array('Q1',   12,   15,		21),
		array('Q2',   56,   73,		86),
		array('Q3',   52,   61,		69),
		array('Q4',   30,   32,		0),
	)
);
$dataseriesLabels1 = array(new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$1', NULL, 1), //    2010
			new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$1', NULL, 1), //    2011
			new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$1', NULL, 1), //    2012
		);
$xAxisTickValues1 = array(
			new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$5', NULL, 4), 
		);
$dataSeriesValues1 = array(
			new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', NULL, 4),
			new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', NULL, 4),
			new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$2:$D$5', NULL, 4),
		);
$series1 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_PIECHART,				// plotType
	PHPExcel_Chart_DataSeries::GROUPING_STANDARD,			// plotGrouping
	range(0, count($dataSeriesValues1)-1),					// plotOrder
	$dataseriesLabels1,										// plotLabel
	$xAxisTickValues1,										// plotCategory
	$dataSeriesValues1										// plotValues
);
//	Set up a layout object for the Pie chart
$layout1 = new PHPExcel_Chart_Layout();
$layout1->setShowVal(TRUE);
$layout1->setShowPercent(TRUE);
//	Set the series in the plot area
$plotarea1 = new PHPExcel_Chart_PlotArea($layout1, array($series1));
//	Set the chart legend
$legend1 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, null, false);
$title1 = new PHPExcel_Chart_Title('Test Pie Chart');
//	Create the chart
$chart1 = new PHPExcel_Chart(
	'chart1',		// name
	$title1,		// title
	$legend1,		// legend
	$plotarea1,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	null,			// xAxisLabel
	null			// yAxisLabel		- Pie charts don't have a Y-Axis
);
//	Set the position where the chart should appear in the worksheet
$chart1->setTopLeftPosition('A7');
$chart1->setBottomRightPosition('H20');
//	Add the chart to the worksheet
$objWorksheet->addChart($chart1);
$dataseriesLabels2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$1', null, 1),	//	2011
);
$xAxisTickValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$5', null, 4),	//	Q1 to Q4
);
$dataSeriesValues2 = array(
	new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', null, 4),
);
$series2 = new PHPExcel_Chart_DataSeries(
	PHPExcel_Chart_DataSeries::TYPE_DONUTCHART,		// plotType
	PHPExcel_Chart_DataSeries::GROUPING_STANDARD,	// plotGrouping
	range(0, count($dataSeriesValues2)-1),			// plotOrder
	$dataseriesLabels2,								// plotLabel
	$xAxisTickValues2,								// plotCategory
	$dataSeriesValues2								// plotValues
);
//	Set up a layout object for the Pie chart
$layout2 = new PHPExcel_Chart_Layout();
$layout2->setShowVal(TRUE);
$layout2->setShowCatName(TRUE);
//	Set the series in the plot area
$plotarea2 = new PHPExcel_Chart_PlotArea($layout2, array($series2));
$title2 = new PHPExcel_Chart_Title('Test Donut Chart');
//	Create the chart
$chart2 = new PHPExcel_Chart(
	'chart2',		// name
	$title2,		// title
	NULL,			// legend
	$plotarea2,		// plotArea
	true,			// plotVisibleOnly
	0,				// displayBlanksAs
	null,			// xAxisLabel
	null			// yAxisLabel		- Like Pie charts, Donut charts don't have a Y-Axis
);
//	Set the position where the chart should appear in the worksheet
$chart2->setTopLeftPosition('I7');
$chart2->setBottomRightPosition('P20');
//	Add the chart to the worksheet
$objWorksheet->addChart($chart2);
// Save Excel 2007 file
//echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$filename='test.xlsx'; //save our workbook as this file name
header('Content-Type: application/vnd.ms-excel'); //mime type
header('Content-Disposition: attachment;filename="'.$filename.'"'); //tell browser what's the file name
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->setIncludeCharts(TRUE);
$objWriter->save('php://output');
//$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
//echo date('H:i:s') , " File written to " , str_replace('.php', '.xlsx', pathinfo(__FILE__, PATHINFO_BASENAME)) , EOL;
// Echo memory peak usage
//echo date('H:i:s') , " Peak memory usage: " , (memory_get_peak_usage(true) / 1024 / 1024) , " MB" , EOL;
// Echo done
//echo date('H:i:s') , " Done writing file" , EOL;
//echo 'File has been created in ' , getcwd() , EOL;
