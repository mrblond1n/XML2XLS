<?php

//    https://otpravka.pochta.ru/specification#/nogroup-normalization_adress
//    https://otpravka.pochta.ru/specification#/nogroup-rate_calculate

function utf8ucfirst($string)
{
    $first = mb_strtoupper(mb_substr($string, 0, 1, 'UTF-8'), 'UTF-8');
    $other = mb_substr(mb_strtolower($string, 'UTF-8'), 1, 100500, 'UTF-8');
    return $first . $other;
}

require_once __DIR__ . '/config.php';

if (isset($_FILES['file']['error']) && $_FILES['file']['error'] == 0 && substr($_FILES['file']['name'], -4) == '.xml') {
    $xml = json_decode(json_encode(simplexml_load_file($_FILES['file']['tmp_name'])), true);

    $codes = array();
    $f = fopen('codes.txt', 'r');
    while (!feof($f)) {
        $s = trim(fgets($f));if ($s == '') {
            continue;
        }

        list($code, $price, $kolvo, $mailtype) = explode('|', $s);
        $codes[$code] = array('price' => $price, 'kolvo' => $kolvo, 'mailtype' => $mailtype);
    }
    fclose($f);

    $indexes = array();
    $f = fopen('indexes.txt', 'r');
    while (!feof($f)) {
        $s = trim(fgets($f));if ($s == '') {
            continue;
        }

        list($index, $gorod, $region, $codez, $optimize) = explode('|', $s);
        $indexes[$index]['codes'] = explode(',', $codez);
        $indexes[$index]['optimize'] = $optimize;
        $indexes[$index]['gorod'] = $gorod;
        $indexes[$index]['region'] = $region;
    }
    fclose($f);

    require_once __DIR__ . '/phpexcel/Classes/PHPExcel.php';
    $phpexcel = new PHPExcel();
    $page = $phpexcel->setActiveSheetIndex(0);
    $page->setCellValue("A1", "ADDRESSLINE");
    $page->setCellValue("B1", "ADRESAT");
    $page->setCellValue("C1", "MASS");
    $page->setCellValue("D1", "VALUE");
    $page->setCellValue("E1", "PAYMENT");
    $page->setCellValue("F1", "COMMENT");
    $page->setCellValue("G1", "MAILTYPE");
    $page->setCellValue("H1", "ДОСТАВКА");
    $page->setCellValue("I1", "ОШИБКА");
    $page->setCellValue("J1", "МИНИМУМ");

    foreach ($xml['Order'] as $k => $v) {
        $error = '';
        $highlite = '';
        $zakaz = $v['ExtID'];
        $srcaddress = $v['ClientReceiver']['Address'];
        $address = '';
        if ($srcaddress['Region']) {
            $address .= $srcaddress['Region'] . ', ';
        }

        if ($srcaddress['Area']) {
            $address .= $srcaddress['Area'] . ', ';
        }

        if ($srcaddress['City']) {
            $address .= $srcaddress['City'] . ', ';
        }

        if ($srcaddress['Street']) {
            $address .= $srcaddress['Street'] . ', ';
        }

        if ($srcaddress['Home']) {
            $address .= 'д.' . $srcaddress['Home'] . ', ';
        }

        if ($srcaddress['Building']) {
            $address .= $srcaddress['Building'] . ', ';
        }

        if ($srcaddress['Flat']) {
            $address .= 'кв.' . $srcaddress['Flat'] . ', ';
        }

        $address = mb_substr($address, 0, -2, 'UTF-8');

        $cnt = 0;
        foreach ($v['Content']['Item'] as $v1) {
            if (is_array($v1)) {
                $cnt++;
            }

        }
        if ($cnt > 0) {
            $error .= " \r\nНеверно указано количество товара";
            $highlite = 'FF0000';
        } else {
            $code = $v['Content']['Item']['GoodsCode'];
            $count = $v['Content']['Item']['Count'];
        }
        $ordersum = sprintf('%02.2f', 600 * $count);
        $deliverysum = $v['OrderDeliverySum'];

        if (!isset($indexes[$srcaddress['Zipcode']])) {
            $error .= " \r\nHет кодов отправления для данного почтового индекса";
            $highlite = 'FF0000';
        } elseif (!in_array($code, $indexes[$srcaddress['Zipcode']]['codes'])) {
            $error .= " \r\nКод отправления $code не соответствует почтовому индексу, допустимые " . implode(',', $indexes[$srcaddress['Zipcode']]['codes']);
            $highlite = 'FF0000';
        } elseif ($count > $codes[$code]['kolvo']) {
            $error .= " \r\nПревышен максимум товаров (" . $codes[$code]['kolvo'] . ' шт) для кода отправления ' . $code . ', указано ' . $count . ' шт';
            $highlite = 'FF0000';
        } elseif ($deliverysum != $codes[$code]['price']) {
            $error .= " \r\nНеверная сумма доставки для кода отправления $code, должно быть " . $codes[$code]['price'] . ' руб';
            $highlite = 'FF0000';
        }
        $min = $deliverysum;
        $minstr = '';
        foreach ($indexes[$srcaddress['Zipcode']]['codes'] as $tempcode) {
            if ($count <= $codes[$tempcode]['kolvo'] && $codes[$tempcode]['price'] < $min) {
                $min = $codes[$tempcode]['price'];
                $minstr = 'Есть возможность отправки дешевле, код отправления ' . $tempcode . ', стоимость ' . $min;
            }
        }

//          if($deliverysum==$codes[$code]['price']) $deliverysum='0.00';

        $mailtype = '';
        if ($code == 1) {
            $mailtype = '';
        }

        $address = $srcaddress['Zipcode'] . ', ' . utf8ucfirst($indexes[$srcaddress['Zipcode']]['gorod']) . ', ';
        if ($srcaddress['Street']) {
            $address .= $srcaddress['Street'] . ', ';
        }

        if ($srcaddress['Home']) {
            $address .= 'д.' . $srcaddress['Home'] . ', ';
        }

        if ($srcaddress['Building']) {
            $address .= $srcaddress['Building'] . ', ';
        }

        if ($srcaddress['Flat']) {
            $address .= 'кв.' . $srcaddress['Flat'] . ', ';
        }

        $address = mb_substr($address, 0, -2, 'UTF-8');

        $page->setCellValue("A" . ($k + 2), $address);
        $page->setCellValue("B" . ($k + 2), utf8ucfirst($v['ClientReceiver']['LastName']) . ' ' . utf8ucfirst($v['ClientReceiver']['FirstName']) . ' ' . utf8ucfirst($v['ClientReceiver']['MiddleName']));
        $page->setCellValue("C" . ($k + 2), sprintf('%01.2f', 1.75 * $count));
        $page->setCellValue("D" . ($k + 2), $ordersum);
        $page->setCellValue("E" . ($k + 2), 0);
        $page->setCellValue("F" . ($k + 2), $zakaz);
        $page->setCellValue("G" . ($k + 2), $codes[$code]['mailtype']);
        $page->setCellValue("H" . ($k + 2), $deliverysum);
        $page->setCellValue("I" . ($k + 2), ($error == '' ? $error : substr($error, 3)));
        $page->setCellValue("J" . ($k + 2), $minstr);

        if ($highlite) {
            $page->getStyle("A" . ($k + 2))->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB($highlite);
        }

    }

    $page->getColumnDimension("A")->setAutoSize(true);
    $page->getStyle("A1")->getFont()->setBold(true);
    $page->getColumnDimension("B")->setAutoSize(true);
    $page->getStyle("B1")->getFont()->setBold(true);
    $page->getColumnDimension("C")->setAutoSize(true);
    $page->getStyle("C1")->getFont()->setBold(true);
    $page->getColumnDimension("D")->setAutoSize(true);
    $page->getStyle("D1")->getFont()->setBold(true);
    $page->getColumnDimension("E")->setAutoSize(true);
    $page->getStyle("E1")->getFont()->setBold(true);
    $page->getColumnDimension("F")->setAutoSize(true);
    $page->getStyle("F1")->getFont()->setBold(true);
    $page->getColumnDimension("G")->setAutoSize(true);
    $page->getStyle("G1")->getFont()->setBold(true);
    $page->getColumnDimension("H")->setAutoSize(true);
    $page->getStyle("H1")->getFont()->setBold(true);
    $page->getColumnDimension("I")->setAutoSize(true);
    $page->getStyle("I1")->getFont()->setBold(true);
    $page->getColumnDimension("J")->setAutoSize(true);
    $page->getStyle("I1")->getFont()->setBold(true);

    header('Content-type: application/vnd.ms-excel');
    header('Content-Disposition: attachment; filename="file.xls"');
    header('Cache-Control: max-age=0');

    $objWriter = PHPExcel_IOFactory::createWriter($phpexcel, 'Excel5');
    $objWriter->save('php://output');

} else {

    $content = '
<div class="hidden-xs hidden-md col-lg-4"></div>
<div class="col-xs-12 col-md-12 col-lg-4">
<b style="text-transform:uppercase">' . $row['firstname'] . ' ' . $row['lastname'] . '<br><br></b>
<form action="" method="POST" enctype="multipart/form-data">
  <div class="form-group">
    <label for="file">&nbsp;Загрузка XML файла</label>
    <input type="file" class="form-control" name="file" id="file" value="" placeholder="Загрузка XML файла">
  </div>
  <button type="submit" class="btn btn-success btn-primary btn-lg" style="width:100%">Загрузка XML файла</button>
</form>
</div>
<div class="hidden-xs hidden-md col-lg-4"></div>';
}

if ($error) {
    $content = '
<div class="hidden-xs hidden-md col-lg-4"></div>
<form class="col-xs-12 col-md-12 col-lg-4" action="" method="POST">
  <div class="form-group">
    <label for="email">&nbsp;Введите свой e-mail</label>
    <input type="email" class="form-control" name="email" id="email" value="" placeholder="Введите свой e-mail">
  </div>
  <div class="form-group">
    <label for="password">&nbsp;Введите свой пароль</label>
    <input type="password" class="form-control" name="password" id="password" value="" placeholder="Введите свой пароль">
  </div>
  <button type="submit" class="btn btn-success btn-primary btn-lg" style="width:100%">Войти</button>
</form>
<div class="hidden-xs hidden-md col-lg-4"></div>';
}

?>
<?php header('Content-type: text/html; charset=utf-8');?>
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link href="css/bootstrap.css" rel="stylesheet">
<script src="js/jquery.modern.js"></script>
</head>
<body>
<div class="container" style="padding:30px">
<div class="row">
<?=$content?>
</div>
</div>
</body>
</html>