<?php

date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

// Excelファイルの読込                              i
// コマンドラインで読込むExcelファイルを渡す
$obj = PHPExcel_IOFactory::createreader('Excel2007');
$book = $obj->load($argv[1]);

// シートの設定
$book->setActiveSheetIndex(0);
$sheet = $book->getActiveSheet();

// 出力SQLファイル名
const OUTPUT_SQL_FILE_NAME = 'insert_category_redirect.sql';

$sql = "";
$count = 0;

for ($row = 3;; $row++) {
    $old_category_id = $sheet->getCellByColumnAndRow(0, $row)->getValue();

    // 旧カテゴリID列がNULLなら抜けてシートの読込終了
    if(!isset($old_category_id)) {
        break;
    }
    $new_category_id = $sheet->getCellByColumnAndRow(3, $row)->getValue();

    if(isset($new_category_id)){
        $insert_sql = <<< EOS
INSERT INTO dtb_category_redirect (old_category_id, new_category_id, create_date, update_date) VALUES ({$old_category_id}, {$new_category_id}, now(), now());
EOS;
        $sql .= $insert_sql . "\n";
        $count++;
    }
}

$lines = $row - 1;
echo $lines . "行までを読み込みました。" . "\n";
echo $count . "行のSQLを作成しました。 " .  "\n";

$file = OUTPUT_SQL_FILE_NAME;
file_put_contents($file, $sql);
