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

// 新カテゴリIDは1000から振り直す
const ADJUSTMENT_ID_PLUS = 1000;
// 大カテゴリの親カテゴリID
const LARGE_PARENT_CATEGORY_ID = 0;
// 大カテゴリのlevel
const LARGE_CATEGORY_LEVEL = 1;
// 中カテゴリのlevel
const MIDDLE_CATEGORY_LEVEL = 2;
// 小カテゴリのlevel
const SMALL_CATEGORY_LEVEL = 3;
// ランク用の調整数字
// const NUMBER_FOR_RANK = 3000;

const CREATOR_ID = 2;
const SHOP_ID = 0;

$creator_id = CREATOR_ID;
$shop_id = SHOP_ID;
$sql = "";

$count = 0;

for ($row = 3;; $row++) {
    $no_cell = $sheet->getCellByColumnAndRow(1, $row)->getValue();
    // No列がNULLなら抜けてシートの読込終了
    if(!isset($no_cell)) {
        break;
    }

    // 大カテゴリ
    $large_category_name = $sheet->getCellByColumnAndRow(2, $row)->getValue();
    if(isset($large_category_name)) {
        $val = $sheet->getCellByColumnAndRow(1, $row)->getCalculatedValue();
        $large_category_id = $val + ADJUSTMENT_ID_PLUS;
        $parent_category_id = LARGE_PARENT_CATEGORY_ID;
        $level = LARGE_CATEGORY_LEVEL;
        // $rank = NUMBER_FOR_RANK - $large_category_id;
        $rank = $large_category_id;

        $insert_sql = <<< EOS
INSERT INTO dtb_category (category_id, category_name, parent_category_id, level, rank, creator_id, create_date, update_date, shop_id) VALUES ({$large_category_id}, '{$large_category_name}', {$parent_category_id}, {$level}, {$rank}, {$creator_id}, now(), now(), {$shop_id});
EOS;
        $sql .= $insert_sql . "\n";
        $count++;
    } else {
        // 中カテゴリ
        $middle_category_name = $sheet->getCellByColumnAndRow(3, $row)->getValue();
        if(isset($middle_category_name)) {
            $val = $sheet->getCellByColumnAndRow(1, $row)->getCalculatedValue();
            $middle_category_id = $val + ADJUSTMENT_ID_PLUS;
            $parent_category_id = $large_category_id;
            $level = MIDDLE_CATEGORY_LEVEL;
            // $rank = NUMBER_FOR_RANK - $middle_category_id;
            $rank = $middle_category_id;

            $insert_sql = <<< EOS
INSERT INTO dtb_category (category_id, category_name, parent_category_id, level, rank, creator_id, create_date, update_date, shop_id) VALUES ({$middle_category_id}, '{$middle_category_name}', {$parent_category_id}, {$level}, {$rank}, {$creator_id}, now(), now(), {$shop_id});
EOS;
            $sql .= $insert_sql . "\n";
            $count++;
        } else {
            // 小カテゴリ
            $small_category_name = $sheet->getCellByColumnAndRow(4, $row)->getValue();
            if(isset($small_category_name)) {
                $val = $sheet->getCellByColumnAndRow(1, $row)->getCalculatedValue();
                $small_category_id = $val + ADJUSTMENT_ID_PLUS;
                $parent_category_id = $middle_category_id;
                $level = SMALL_CATEGORY_LEVEL;
                // $rank = NUMBER_FOR_RANK - $small_category_id;
                $rank = $small_category_id;

                $insert_sql = <<< EOS
INSERT INTO dtb_category (category_id, category_name, parent_category_id, level, rank, creator_id, create_date, update_date, shop_id) VALUES ({$small_category_id}, '{$small_category_name}', {$parent_category_id}, {$level}, {$rank}, {$creator_id}, now(), now(), {$shop_id});
EOS;
                $sql .= $insert_sql . "\n";
                $count++;
            } else {
                echo $row . "行目が空行です。" . "\n";
            }
        }
    }
}

$lines = $row - 1;

echo $lines . "行までを読み込みました。" . "\n";
echo $count . "行のSQLを作成しました。 " .  "\n";

$file = 'new_category_insert.sql';
file_put_contents($file, $sql);

