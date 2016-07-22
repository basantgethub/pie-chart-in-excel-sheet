 <?php
	
	$data = Array ( 
				'0' => Array ( 
						'Date' => 'Jun 29th 2016',
						'Email Invites' => '10' ,
						'Facebook Shares' => 4 ,
						'Tweets' => 2 ,
						'Facebook Messages' => 12,
						'Whatsapp' => 20,
						'SMS' => 20,
						'Manual Leads' => 14,
						'Google+ Converts' => 1,
						'LinkedIn Shares' => 30,
						'Pinterest' => 70 
						),
				'1' => Array ( 
						'Date' => 'Jul 1st 2016',
						'Email Invites' => '8',
						'Facebook Shares' => 40,
						'Tweets' => 50,
						'Facebook Messages' => 20,
						'Whatsapp' => 70,
						'SMS' => 60,
						'Manual Leads' => 10,
						'Google+ Converts' => 0,
						'LinkedIn Shares' => 10,
						'Pinterest' => 30 ),
				'2' => Array ( 
						'Date' => 'Jul 11th 2016',
						'Email Invites' => '20',
						'Facebook Shares' => 70,
						'Tweets' => 40,
						'Facebook Messages' => 30,
						'Whatsapp' => 80,
						'SMS' => 90,
						'Manual Leads' => 50,
						'Google+ Converts' => 20,
						'LinkedIn Shares' => 10,
						'Pinterest' => 8 ),
			);
		return exportChartCsv($data,'test');
		
         function exportChartCsv($data = '', $filename =''){
               
                include("excelPHP\Classes\PHPExcel.php");
               
                $setActive = array_keys($data[0]);
                unset($setActive[0]);
               
                $objPHPExcel = new PHPExcel();
                
                foreach($setActive as $key => $value) {
                  
                    $objWorksheet = $objPHPExcel->setActiveSheetIndex($key - 1);
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
        }
    