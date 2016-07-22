 <?php
	
	 $data = array();
	$data=
		array(
			array('',	2010,	2011,	2012),
			array('Q1',   12,   15,		21),
			array('Q2',   56,   73,		86),
			array('Q3',   52,   61,		69),
			array('Q4',   30,   32,		0),
		);
	
	
    $filename =   'Excel-Demo';    
                include "excelPHP/Classes/PHPExcel.php";
               
                $objPHPExcel = new PHPExcel();
                
                foreach($data as $key => $value) {
                  print_r($key);
				  
                    $objWorksheet = $objPHPExcel->setActiveSheetIndex($key - 1);
					//$objWorksheet = $objPHPExcel->setActiveSheetIndex($key);
                    $objWorksheet = $objPHPExcel->getActiveSheet()->setCellValue('A1', $value);
                    $objWorksheet = $objPHPExcel->getActiveSheet()->setTitle($value);
                    $sheetName = $objPHPExcel->getActiveSheet()->getTitle();
                    
                    if(!empty($value) && $value != ' '){
                        $sheetData = array();
                        foreach ($data as $row => $val){
                            $sheetData[$row] = array($val['Date'], $val[$value]);
                        }

                        $sheet = array(
                                array('Date',       $value)
                            );
                        
                        
                        $sheet = array_merge($sheet, $sheetData);
                        $countSheet = count($sheet);

                        $objWorksheet ->fromArray($sheet);

                        $dataseriesLabels1 = array(
                            new PHPExcel_Chart_DataSeriesValues('String',"'".$sheetName."'".'!$B$1', NULL, 1),   //  Temperature
                        );

                        $xAxisTickValues = array(
                                new PHPExcel_Chart_DataSeriesValues('String',"'".$sheetName."'".'!$A$2:$A'.$countSheet, NULL, $countSheet),    //  Jan to Dec
                        );

                        $dataSeriesValues1 = array(
                                new PHPExcel_Chart_DataSeriesValues('Number', "'".$sheetName."'".'!$B$2:$B'.$countSheet, NULL, $countSheet),
                        );

                        //  Build the dataseries
                        $series1 = new PHPExcel_Chart_DataSeries(
                                PHPExcel_Chart_DataSeries::TYPE_BARCHART,       // plotType
                                PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED,  // plotGrouping
                                range(0, count($dataSeriesValues1)-1),          // plotOrder
                                $dataseriesLabels1,                             // plotLabel
                                $xAxisTickValues,                               // plotCategory
                                $dataSeriesValues1                              // plotValues
                        );
                        //  Set additional dataseries parameters
                        //      Make it a vertical column rather than a horizontal bar graph
                        $series1->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);


                        //  Set the series in the plot area
                        $plotarea = new PHPExcel_Chart_PlotArea(NULL, array($series1));
                        //  Set the chart legend
                        $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);

                        $title = new PHPExcel_Chart_Title($value);


                        //  Create the chart
                        $chart = new PHPExcel_Chart(
                                '',       // name
                                $title,         // title
                                $legend,        // legend
                                $plotarea,      // plotArea
                                true,           // plotVisibleOnly
                                0,              // displayBlanksAs
                                NULL,           // xAxisLabel
                                NULL            // yAxisLabel
                        );

                        //  Set the position where the chart should appear in the worksheet
                        $chart->setTopLeftPosition('F2');
                        $chart->setBottomRightPosition('O16');
                        //  Add the chart to the worksheet
                        $objWorksheet->addChart($chart);
                    }

                    $objPHPExcel->createSheet();
                }
               
                $filename .= '.xls';
                header('Content-Type: application/vnd.ms-excel');
                header('Content-Disposition: attachment;filename="'.$filename.'"');
                header('Cache-Control: max-age=0');

                $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
                $objWriter->setIncludeCharts(TRUE);
                $objWriter->save('php://output');
    