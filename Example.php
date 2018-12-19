<?php

require 'excel.php';
$ExcelOper = new ExcelOper();

//导出数据
$new_data = null; //数据列表array;
$head = array("订单号", "订单金额");
$fields = array("order_no", "price"); //数据列表字段名
$file_name = 'check_order_' . date("ymdhis", time()); //文件名
$ExcelOper->exportExcel($head, $fields, $new_data, $file_name);


//导入数据
$List = $ExcelOper->importExcel($filepath); //文件第一行为标题 第二行起为数据
 