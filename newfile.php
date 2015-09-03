<?php

include_once ( __DIR__ . '/PHPExcel_1.8.0_doc/Classes/PHPExcel.php');
include_once ( __DIR__ . '/PHPExcel_1.8.0_doc/Classes/PHPExcel/IOFactory.php');

//エクセルファイルの新規作成
$excel = new PHPExcel();

// シートの設定
$excel->setActiveSheetIndex(0);//何番目のシートか
$sheet = $excel->getActiveSheet();//有効になっているシートを代入

// セルに値を入力
$sheet->setCellValue('A1', 'こんにちは！');//A1のセルにこんにちは！という値を入力

// Excel2007形式で出力する
$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');

$file_name = 'output.xlsx';
$writer->save($file_name);

// $file_path = dirname(__FILE__) . $file_name;//ダウンロードさせるファイルのパス
$file_path = $file_name;//ダウンロードさせるファイルのパス

header("Content-Type: application/octet-stream");//ダウンロードの指示
header("Content-Disposition: attachment; filename=$file_name");//ダウンロードするファイル名
header("Content-Length:".filesize($file_path));//ダウンロードするファイルのサイズ
ob_end_clean();//ファイル破損エラー防止
readfile(dirname(__FILE__) . '/./' . $file_path);//ダウンロード
