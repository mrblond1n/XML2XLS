<!-- if (!isset($indexes[$srcaddress['Zipcode']])) {
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

    if($deliverysum==$codes[$code]['price']) $deliverysum='0.00'; 
    

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

         if($deliverysum==$codes[$code]['price']) $deliverysum='0.00';

         $mailtype = '';
        if ($code == 1) {
            $mailtype = '';
        }

         // if ($highlite) {
        //     $page->getStyle("A" . ($k + 2))->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setRGB($highlite);
        // }
    
    -->