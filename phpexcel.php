/** 
    *   方法名 :   excelToTable 
    *   作用  :   【私有】将excel数据导入数据表中 
    *   @param1 ：   file 用户上传的文件信息 
    *   @param2 ：   tableid 用来区别是哪张表，1-statistics_rawdata_pct，2-statistics_rawdata_apply，3-statistics_rawdata_auth，4-statistics_rawdata_valid 
    *   @param3 ：   month_number 导入的数据属于哪一期的，比如201510 
    *   @param4 ：   table_head 用来判断excel表格是否有表头，默认有 
    *   @author :    canxiaochen
    */    
    private function excelToTable($file,$tableid,$month_number,$table_head=1){  
        if(!empty($file['name'])){  
              
            $file_types = explode ( ".", $file['name'] );  
            $excel_type = array('xls','csv','xlsx');  
            //判断是不是excel文件  
            if (!in_array(strtolower(end($file_types)),$excel_type)){  
                $this->show_msg("不是Excel文件，重新上传","/search/patentStatistics/uploadRawdata");  
            }  
  
            //设置上传路径  
            $savePath = _WWW_ . 'www/tmp/';  
  
            //以时间来命名上传的文件  
            $str = date ( 'Ymdhis' );  
            $file_name = $str.".".end($file_types);  
  
            //是否上传成功  
            $tmp_file = $file['tmp_name'];  
            if (!copy($tmp_file,$savePath.$file_name)){  
                $this->show_msg("上传失败","/search/patentStatistics/uploadRawdata");  
            }  
              
            if($tableid=="1"){  
                $rawdata_obj = $this->rawdata_pctmodel;    
            }elseif($tableid=="2"){  
                $rawdata_obj = $this->rawdata_applymodel;      
            }elseif($tableid=="3"){  
                $rawdata_obj = $this->rawdata_authmodel;   
            }elseif($tableid=="4"){  
                $rawdata_obj = $this->rawdata_validmodel;  
            }else{  
                $this->show_msg("您要导入的数据表不存在！","/search/patentStatistics/uploadRawdata");  
            }  
              
            if($rawdata_obj)  
                $fields = $rawdata_obj->returnFields();  
            else  
                $this->show_msg("未能指定明确的表！","/search/patentStatistics/uploadRawdata");  
              
            //定义导入失败记录的文档  
            $logfile = $savePath.$str.'.txt';  
              
            //读取excel，存成数组，该数组的key是从1开始  
            $res = $this->excelToArray($savePath.$file_name,end($file_types));  
            //echo 12321321;exit;  
            //如果有表头，则过滤掉第一行  
            if($table_head)  
                unset($res[1]);  
              
            //循环写入，不一次性写入，防止有错误的记录；错误记录会记录下第一个字段到txt文档中去  
            foreach($res as $k =>$v){  
                foreach($fields as $key=>$val){  
                    if($v[$key]===null){  
                        $v[$key] = 'null';  
                    }  
                    $data[$val] = $v[$key];  
                }  
                //该字段比较特殊，必须导入表中都有该字段  
                $data['month_number'] = $month_number;  
                $result = $rawdata_obj->addSave($data);  
                unset($data);  
                if(!$result){  
                    $this ->logFile($logfile,$v[0]);  
                }  
            }  
            if(file_get_contents($logfile))  
                return $logfile;  
            else  
                return true;  
        }  
    }  
      
    /** 
    *   方法名 :   excelToArray 
    *   作用  :   【私有】将excel数据转换成数组 
    *   @param1 ：   filename excel文件名 
    *   @param2 ：   filetype excel格式（xls、xlsx、csv） 
    *   @param3 ：   encode 编码格式，默认utf8 
    *   @return ：   返回2维数组，最小的key为1 
    *   @author :   canxiaochen 
    */    
    private function excelToArray($filename,$filetype,$encode='utf-8'){  
        if(strtolower($filetype)=='xls'){  
            $objReader = PHPExcel_IOFactory::createReader('Excel5');  
        }elseif(strtolower($filetype)=='xlsx'){  
            $objReader = PHPExcel_IOFactory::createReader('Excel2007');  
        }elseif(strtolower($filetype)=='csv'){  
            $objReader = PHPExcel_IOFactory::createReader('CSV');  
        }  
              
        $objReader->setReadDataOnly(true);  
        $objPHPExcel = $objReader->load($filename);  
        $objWorksheet = $objPHPExcel->getActiveSheet();  
        $highestRow = $objWorksheet->getHighestRow();  
        $highestColumn = $objWorksheet->getHighestColumn();  
        $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);  
        $excelData = array();  
        for ($row = 1; $row <= $highestRow; $row++) {  
            for ($col = 0; $col < $highestColumnIndex; $col++) {  
                $excelData[$row][] =(string)$objWorksheet->getCellByColumnAndRow($col, $row)
                ->getValue();  
            }  
        }  
        return $excelData;  
    } 



/** 
    *   方法名 :   exportData 
    *   作用  :   导出数据 
    *   @author canxiaochen
    *   @return excel文件路径 
    */    
    public function exportDataAction(){  
        $title = strip_tags($_POST['title']);//excel第一行标题  
        $max_col_num = $_POST['max_col_num'];//最大列数  
        $th_num_arr = explode(',',trim($_POST['th_col_num_str']));//取th各行的列数  
        array_shift($th_num_arr);//删除首行th  
        $head_line = count($th_num_arr);//列标题的th行数  
        $th_data = explode('@@@',trim($_POST['th_data_str']));  
        array_shift($th_data);//删除首行th(就是第一行标题)  
          
        $th_data2 = array();  
        foreach($th_data as $k=>$v){  
            $th_data2[] = strip_tags($v);     
        }  
        //将一维数组（值）按照另一个数组（个数）拆分成二维数组  
        foreach($th_num_arr as $k=>$v){  
            foreach($th_data2 as $key=>$val){  
                if($key<$v)  
                    $temp[] = $val;   
            }  
            $th_data2 = array_values(array_diff($th_data2,$temp));  
            $head[] = $temp;  
            unset($temp);  
        }  
        //补空  
        foreach($head as $k=>$v){  
            if(count($head[$k])<$max_col_num){  
                for($i=0;$i<$max_col_num-count($head[$k]);$i++){  
                    $temp[] = '';  
                }  
                if($k==0)  
                    $head2[] = array_merge($head[$k],$temp);  
                else  
                    $head2[] = array_merge($temp,$head[$k]);  
            }  
            unset($temp);  
                  
        }  
              
        //获取所有td的值  
        $td_data = explode('@@@',trim($_POST['td_data_str']));  
        $data = array();  
        foreach($td_data as $k=>$v){  
            $data[$k/$max_col_num][$k%$max_col_num] = strip_tags($v);  
        }  
  
        $path = $this -> getExcel($title,$title,$head2,$data);  
        echo json_encode(array('href'=>$path)) ;  
    }  
      
      
    /** 
    *   方法名:    getExcel 
    *   作用  :   将数据转换为Excel格式 
    *   @author canxiaochen 
    *   @param1 文件名 
    *   @param2 sheet名称 
    *   @param3 字段名(必须二维数组) 
    *   @param4 数据 
    *   @return excel文件 
    */    
    private function getExcel($fileName,$fileName2,$headArr,$data){  
        //对数据进行检验  
        if(empty($data) || !is_array($data)){  
            die("数据必须为数组");  
        }  
        //检查文件名  
        if(empty($fileName)){  
            exit;  
        }  
        //组装文件名  
        $date = date("Y_m_d",time());  
        $fileName .= "_{$date}.xls";  
  
        error_reporting(E_ALL);  
        ini_set('display_errors', TRUE);  
        ini_set('display_startup_errors', TRUE);  
        date_default_timezone_set('PRC');  
  
        if (PHP_SAPI == 'cli')  
            die('只能通过浏览器运行');  
          
        //创建PHPExcel对象  
        $objPHPExcel = new PHPExcel();  
        $objProps = $objPHPExcel->getProperties();  
        //设置表名称  
        $objPHPExcel->setActiveSheetIndex(0) ->setCellValue('A1', $fileName2);  
        //设置表头  
          
        for($i=0;$i<count($headArr);$i++){  
            $line_num = 2;  
            $line_num += $i;  
            $key = ord("A");//A--65  
            $key2 = ord("@");//@--64  
            foreach($headArr[$i] as $v){  
                if($key>ord("Z")){  
                    $key2 += 1;  
                    $key = ord("A");  
                    $colum = chr($key2).chr($key);//超过26个字母时才会启用  dingling 20150626  
                }else{  
                    if($key2>=ord("A")){  
                        $colum = chr($key2).chr($key);  
                    }else{  
                        $colum = chr($key);  
                    }  
                }  
                $objPHPExcel->setActiveSheetIndex(0) ->setCellValue($colum.$line_num,$v);  
                $key += 1;  
            }  
        }  
          
          
        $column = count($headArr)+2;  
        $objActSheet = $objPHPExcel->getActiveSheet();  
  
        foreach($data as $v){ //行写入  
            $span = ord("A");  
            $span2 = ord("@");  
            foreach($headArr[0] as $key=>$val){  
                if($span>ord("Z")){  
                    $span2 += 1;  
                    $span = ord("A");  
                    $j = chr($span2).chr($span);//超过26个字母时才会启用  dingling 20150626  
                }else{  
                    if($span2>=ord("A")){  
                        $j = chr($span2).chr($span);  
                    }else{  
                        $j = chr($span);  
                    }  
                }  
                $objActSheet->setCellValue($j.$column, strip_tags($v[$key]));  
                $span++;  
            }  
            $column++;  
        }  
  
        $fileName = iconv("utf-8", "gb2312", $fileName);  
          
        $objPHPExcel->getActiveSheet()->setTitle($fileName2);// 重命名表  
        $objPHPExcel->setActiveSheetIndex(0);// 设置活动单指数到第一个表,所以Excel打开这是第一个表   
          
        ob_end_clean();//清除缓冲区,避免乱码  
        // Redirect output to a client’s web browser (Excel5)  
        header('Content-Type: application/vnd.ms-excel');  
        header("Content-Disposition: attachment;filename=\"$fileName\"");  
        header('Cache-Control: max-age=0');  
          
        // If you're serving to IE 9, then the following may be needed  
        header('Cache-Control: max-age=1');  
        // If you're serving to IE over SSL, then the following may be needed  
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past  
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified  
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1  
        header ('Pragma: public'); // HTTP/1.0  
  
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
        //$objWriter->save('php://output'); //文件通过浏览器下载  
        //指定存放路径  
        $savePath = _WWW_ . 'www/tmp/';  
        $file = time().'.xls';  
        $objWriter->save($savePath.$file); //将文件存放到指定目录  
        return '/tmp/'.$file;  
    }   
