<?php
include_once ( __DIR__ . '/PHPExcel_1.8.0_doc/Classes/PHPExcel.php');
include_once ( __DIR__ . '/PHPExcel_1.8.0_doc/Classes/PHPExcel/IOFactory.php');

$reader = PHPExcel_IOFactory::createReader('Excel2007');
// $excel = $reader->load("template.xlsx");
$excel = $reader->load("mitsumori.xlsx");


// Excel2007形式(xlsx)で出力する
$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');

$sheet = $excel->getActiveSheet();//有効になっているシートを代入
$sheet->setCellValue('A6', isset($_POST['name']) ? $_POST['name'].'さん' : '名無しさん'); // セルに名前を入力

$file_name = "template_output.xlsx"; //ダウンロードさせるファイル名

$writer->save($file_name);

// $file_path = dirname(__FILE__) . $file_name;//ダウンロードさせるファイルのパス
$file_path = $file_name;//ダウンロードさせるファイルのパス

header("Content-Type: application/octet-stream");//ダウンロードの指示
header("Content-Disposition: attachment; filename=$file_name");//ダウンロードするファイル名
header("Content-Length:".filesize($file_path));//ダウンロードするファイルのサイズ
ob_end_clean();//ファイル破損エラー防止
readfile(dirname(__FILE__) . '/./' . $file_path);//ダウンロード