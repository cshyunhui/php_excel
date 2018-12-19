<?php
 
class Exceloper {
    
    
    /*     * 导出excel
     * @param $head 表头中文
     * @param $fields 字段列表
     * @param $data 数据集合
     * @param $name 文件名
     */
    function exportExcel($head, $fields, $data, $name) {
        set_time_limit(0);
        require_once "src/PHPExcel.php";
        $key_array = array();
        for ($i = 0; $i < 26; $i++) {
            $key_array[] = chr($i + 65);
        }
        $cacheMethod = PHPExcel_CachedObjectStorageFactory:: cache_to_phpTemp;
        $cacheSettings = array('memoryCacheSize' => '8MB');
        PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);
        $objPHPExcel = new PHPExcel();
        $objPHPExcel->setActiveSheetIndex(0);
        $objActSheet = $objPHPExcel->getActiveSheet();
        $objActSheet->setTitle($name);
        foreach ($head as $key => $value) {
            $column = $this->num_to_excel_column($key + 1, $key_array);
            $objActSheet->setCellValueExplicit($column . 1, $value, PHPExcel_Cell_DataType::TYPE_STRING);
        }
        foreach ($data as $k => $obj) {
            $num = $k + 2;
            $j = 1;
            foreach ($fields as $field) {
                $column = $this->num_to_excel_column($j, $key_array);
                $objActSheet->setCellValueExplicit($column . $num, $obj[$field], PHPExcel_Cell_DataType::TYPE_STRING);
                $j++;
            }
        }
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        ob_clean();
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $name . '.xls"');
        header('Cache-Control: max-age=0');
        $objWriter->save('php://output');
    }

    /*     * 导入excel 
     * @param $filePath 文件实际路径
     */
    function importExcel($filePath) {
        require_once  "src/PHPExcel.php";


        /*         * 默认用excel2007读取excel，若格式不对，则用之前的版本进行读取 */
        $PHPReader = new PHPExcel_Reader_Excel2007();
        if (!$PHPReader->canRead($filePath)) {
            $PHPReader = new PHPExcel_Reader_Excel5();
            if (!$PHPReader->canRead($filePath)) {
                echo 'no Excel';
                return;
            }
        }

        $objPHPExcel = $PHPReader->load($filePath);

        /*         * 读取excel文件中的第一个工作表 */
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $highestRow = $objWorksheet->getHighestRow();
        $highestColumn = $objWorksheet->getHighestColumn();
        $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
        $excelData = array();
        $titleData = array();
        for ($row = 1; $row <= $highestRow; $row++) {
            if ($row == 1) {
                for ($col = 0; $col < $highestColumnIndex; $col++) {
                    $titleData[] = (string) $objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
                }
            } else {
                for ($col = 0; $col < $highestColumnIndex; $col++) {
                    $excelData[$row - 1][] = (string) $objWorksheet->getCellByColumnAndRow($col, $row)->getCalculatedValue();
                }
            }
        }
        $rs["title"] = $titleData;
        $rs["data"] = $excelData;
        return $rs;
    }

//数字转换为A、B...AZ
    function num_to_excel_column($n, $key_array) {
        $str = "";
        while ($n > 0) {
            $yu = $n % 26;
            $n = intval($n / 26);
            if ($yu == 0) {
                $str = $str . $key_array[25];
                $n--;
            } else {
                $str = $str . $key_array[$yu - 1];
            }
        }
        return strrev($str);
    }

}
