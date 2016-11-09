<?php

date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/vendor/autoload.php';

// 新カテゴリIDは1000から振り直す
const ADJUSTMENT_ID_PLUS = 1000;

// Excelファイルの読込
$file = "new_category.xlsx";
$obj = PHPExcel_IOFactory::createreader('Excel2007');
$book = $obj->load($file);

// シートの設定
$book->setActiveSheetIndex(0);
$sheet = $book->getActiveSheet();

$arr_product_categories = array();
$sql = "";

for ($row = 3;; $row++) {
// for ($row = 3; $row <= 5; $row++) {
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
        $category_id = $large_category_id;
        $category_name = $large_category_name;
    } else {
        // 中カテゴリ
        $middle_category_name = $sheet->getCellByColumnAndRow(3, $row)->getValue();
        if(isset($middle_category_name)) {
            $val = $sheet->getCellByColumnAndRow(1, $row)->getCalculatedValue();
            $middle_category_id = $val + ADJUSTMENT_ID_PLUS;
            $category_id = $middle_category_id;
            $category_name = $middle_category_name;
        } else {
            // 小カテゴリ
            $small_category_name = $sheet->getCellByColumnAndRow(4, $row)->getValue();
            if(isset($small_category_name)) {
                $val = $sheet->getCellByColumnAndRow(1, $row)->getCalculatedValue();
                $small_category_id = $val + ADJUSTMENT_ID_PLUS;
                $category_id = $small_category_id;
                $category_name = $small_category_name;
            }
        }
    }

    for ($column = 5;; $column++) {
        $product_name = $sheet->getCellByColumnAndRow($column, $row)->getValue();

        if(isset($product_name)) {
            $insert_sql = <<< EOS
UPDATE dtb_products SET category_id = {$category_id} WHERE name = '{$product_name}';
EOS;
            $sql .= $insert_sql . "\n";
        } else {
            break;
        }
    }
}

$file = 'product_category_relation.sql';
file_put_contents($file, $sql);

// // MySQLへの接続処理
// try {
//     $pdo = new PDO('mysql:host=192.168.42.30:3306;dbname=cuoca_dev;charset=utf8', 'root', 'pass', array(PDO::ATTR_EMULATE_PREPARES => false));
// } catch ( PDOException $e) {
//     exit('データベース接続失敗.' . $e->getMessage());
// }
//
// $stmt = $pdo->query("SELECT * FROM dtb_products WHERE del_flg = 0 AND haiban_flg = 0 AND status = 1 ORDER BY product_id");
//
// while ($row = $stmt->fetch(PDO::FETCH_ASSOC)) {
//     echo $row['product_id'] . ', ' . $row['name'] . ', ' . $row['category_id'];
//     if (in_array($row['name'], $arr_product)) {
//         echo ", " . '-----> Excelシートに存在あり.';
//     }
//     echo "\n";
// }
