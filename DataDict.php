<?php

class DataDict
{
    public function createDataDictExcel($dsn,$username,$password,$dbName)
    {
        $pdo = new PDO($dsn,$username,$password);
        $pdo->exec('set names utf8');
        //库中的表信息
        $sql = "select `TABLE_NAME`,`TABLE_COMMENT` from information_schema.`TABLES` as a where a.TABLE_SCHEMA = '".$dbName."';";
        $total_info = $pdo->query($sql);

        $objPHPExcel = new \PHPExcel();
        $objPHPExcel->setactivesheetindex(0);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(60);//设置列宽度
        $objPHPExcel->getActiveSheet()->setCellValue('A1','表名');
        $objPHPExcel->getActiveSheet()->setCellValue('B1','说明');
        foreach ($total_info as $k => $row){
            $objPHPExcel->getActiveSheet()->setCellValue('A'.($k+2),$row[0]);
            $objPHPExcel->getActiveSheet()->getStyle( 'A'.($k+2))->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);
            //生成表内超链接
            $objPHPExcel-> getActiveSheet()-> getCell('A'.($k+2))-> getHyperlink()-> setUrl("sheet://'Worksheet ".($k+1)."'!A1");
            $objPHPExcel->getActiveSheet()->setCellValue('B'.($k+2),$row[1]);

            $objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(50);//设置列宽度
        }

        //遍历各个表信息并分别生成各自的sheet
        $sql = "select `TABLE_NAME` from information_schema.`TABLES` as a where a.TABLE_SCHEMA = '".$dbName."';";
        $total_info = $pdo->query($sql);
        foreach ($total_info as $key => $tableName){
            $objPHPExcel->createSheet();
            $index = intval($key+1);
            $objPHPExcel->setactivesheetindex($index);
            $objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(30);//设置列宽度
            $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(60);//设置列宽度
            //表头信息
            $objPHPExcel->getActiveSheet()->mergeCells('A1:E1');
            $objPHPExcel->getActiveSheet()->getStyle( 'A1')->getFont()->setBold(true);
            $objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            //返回目录
            $objPHPExcel->getActiveSheet()->setCellValue('F1','返回目录');
            $objPHPExcel->getActiveSheet()->getStyle( 'F1')->getFont()->setBold(true);
            $objPHPExcel->getActiveSheet()->getStyle( 'F1')->getFont()->setUnderline(PHPExcel_Style_Font::UNDERLINE_SINGLE);
            $objPHPExcel-> getActiveSheet()-> getCell('F1')-> getHyperlink()->setUrl("sheet://'Worksheet'!A1");

            $objPHPExcel->getActiveSheet()->setCellValue('A1',$tableName[0]);
            $objPHPExcel->getActiveSheet()->setCellValue('A2','字段名称');
            $objPHPExcel->getActiveSheet()->setCellValue('B2','字段类型');
            $objPHPExcel->getActiveSheet()->setCellValue('C2','是否为空');
            $objPHPExcel->getActiveSheet()->setCellValue('D2','是否自增');
            $objPHPExcel->getActiveSheet()->setCellValue('E2','字段说明');
            //查询表信息
            $sql = "select * from information_schema.`COLUMNS` as a where a.TABLE_NAME = '".$tableName[0]."';";
            $tableInfo = $pdo->query($sql);
            foreach ($tableInfo as $key2 => $row){
                $objPHPExcel->getActiveSheet()->setCellValue('A'.($key2+3),$row['COLUMN_NAME']);
                $objPHPExcel->getActiveSheet()->setCellValue('B'.($key2+3),$row['DATA_TYPE']);

                $objPHPExcel->getActiveSheet()->setCellValue('C'.($key2+3),$row['IS_NULLABLE']);
                if(strstr($row['EXTRA'] ,'auto_increment')){
                    $objPHPExcel->getActiveSheet()->setCellValue('D'.($key2+3),'YES');
                }else{
                    $objPHPExcel->getActiveSheet()->setCellValue('D'.($key2+3),"NO");
                }

                $objPHPExcel->getActiveSheet()->setCellValue('E'.($key2+3),$row['COLUMN_COMMENT']);
            }
        }
        $objPHPExcel->setactivesheetindex(0);

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="数据字典.xlsx"');
        header('Cache-Control: max-age=0');

        $objWriter = PHPExcel_IOFactory:: createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save( 'php://output');
        exit;
    }
}


?>