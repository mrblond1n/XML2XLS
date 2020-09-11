<?php

//    https://otpravka.pochta.ru/specification#/nogroup-normalization_adress
//    https://otpravka.pochta.ru/specification#/nogroup-rate_calculate

function utf8ucfirst($string)
{
    $first = mb_strtoupper(mb_substr($string, 0, 1, 'UTF-8'), 'UTF-8');
    $other = mb_substr(mb_strtolower($string, 'UTF-8'), 1, 100500, 'UTF-8');
    return $first . $other;
}

function getItemValues($content) {
    $message = '';
    $count = 0;
    $orderSum = 0;
    foreach ($content as $item) {
        $message .= $item['GoodsCode'] . ' - ' . $item['Count'] . '; ';
        $count += $item['Count'];
        $orderSum += $item['PriceWithDiscount'];
    }
    return array('Message' => $message, 'Count' => $count, 'OrderSum' => $orderSum);
}

function getFullAddress($src) {
    $string = '';
    $string = gettype($src['Zipcode']) === 'array' ? '' : $src['Zipcode'] . ', ';
    $src['Street'] && $string .= $src['Street'] . ', ';
    $src['Home'] && $string .= 'д.' . $src['Home'] . ', ';
    $src['Building'] && $string .= $src['Building'] . ', ';
    $src['Flat'] && $string .= 'кв.' . $src['Flat'] . ', ';
    return $string;
}

function getFullName($client) {
    return utf8ucfirst($client['LastName']) . ' ' . utf8ucfirst($client['FirstName']) . ' ' . utf8ucfirst($client['MiddleName']);
}

require_once __DIR__ . '/config.php';

if (isset($_FILES['file']['error']) && $_FILES['file']['error'] == 0 && substr($_FILES['file']['name'], -4) == '.xml') {
    $xml = json_decode(json_encode(simplexml_load_file($_FILES['file']['tmp_name'])), true);
    $codes = array();

    require_once __DIR__ . '/phpexcel/Classes/PHPExcel.php';
    $phpexcel = new PHPExcel();
    $page = $phpexcel->setActiveSheetIndex(0);
    $page->setCellValue("A1", "ADDRESSLINE");
    $page->setCellValue("B1", "ADRESAT");
    $page->setCellValue("C1", "MASS"); // ВЕС брать из таблицы (товары для акции с почтой)
    $page->setCellValue("D1", "VALUE");
    $page->setCellValue("E1", "PAYMENT");
    $page->setCellValue("F1", "COMMENT");
    $page->setCellValue("G1", "MAILTYPE");
    $page->setCellValue("H1", "COUNT");
    $page->setCellValue("I1", "ORDERSTATUS");
    $page->setCellValue("J1", "ДОСТАВКА");
    $page->setCellValue("K1", "ИНФОРМАЦИЯ");

    foreach ($xml['Order'] as $k => $v) {
        $error = '';
        $highlite = '';
        $orderId = $v['ExtID'];
        $srcAddress = $v['ClientReceiver']['Address'];
        $clientReceiver = $v['ClientReceiver'];
        $address = getFullAddress($srcAddress);
        $addressant = getFullName($clientReceiver);
        $deliverySum = $v['OrderDeliverySum'];
        $orderStatus = $v['OrderStatus'];
        $content = $v['Content'];
        $countItems = 0;
        $mailType = 23;
        $payment = 0;
        $mass = '–';
        foreach ($content['Item'] as $v1) {
            if (is_array($v1)) {
                $countItems++;
            }
        }
        $itemValues = $countItems ? getItemValues($content['Item']) : getItemValues($content);  

        $page->setCellValue("A" . ($k + 2), $address);
        $page->setCellValue("B" . ($k + 2), $addressant);
        $page->setCellValue("C" . ($k + 2), $mass);
        $page->setCellValue("D" . ($k + 2), $itemValues['OrderSum']);
        $page->setCellValue("E" . ($k + 2), $payment);
        $page->setCellValue("F" . ($k + 2), $orderId);
        $page->setCellValue("G" . ($k + 2), $mailType);
        $page->setCellValue("H" . ($k + 2), $itemValues['Count']);
        $page->setCellValue("I" . ($k + 2), $orderStatus);
        $page->setCellValue("J" . ($k + 2), $deliverySum);
        $page->setCellValue("K" . ($k + 2), $itemValues['Message']);

    }
    $arr = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K');
    foreach ($arr as &$value) {
      $page->getColumnDimension($value)->setAutoSize(true);
      $page->getStyle($value . 1)->getFont()->setBold(true);
    };

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
        <div class="hidden-xs hidden-md col-lg-4"></div>
    ';
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
        <div class="hidden-xs hidden-md col-lg-4"></div>
    ';
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