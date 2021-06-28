<?php

require __DIR__ . '/vendor/autoload.php';

use think\facade\Db;

$config = include 'config.php';
Db::setConfig($config['db']);

function loadAliPay2()
{
    global $config;

    // 支付宝
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
    $reader->setInputEncoding('GBK');
    $reader->setDelimiter('@');
    $spreadsheet = $reader->load($config['alipay']['filename']);
    $sheet = $spreadsheet->getActiveSheet();
    for ($i=3; $i<=$sheet->getHighestRow() - 21; $i++) {
        $arr = explode(',', $sheet->getCell('A'.$i)->getValue());
        $type_str = trim($arr[0]);
        $type = $type_str == '收入' ? 1 : ($type_str == '支出' ? 2 : 0);

        $trans_time_str = trim($arr[7]);
        $trans_time = strtotime($trans_time_str);

        $data = [
            'account' => $config['alipay']['account'],
            'account_type_name' => '支付宝', // 账户类型
            'trans_type_name' => trim($arr[3]), // 交易方式
            'trans_person' => trim($arr[1]), // 交易对方
            'type' => $type, // 交易类型 1收入 2支出
            'description' => trim($arr[2]),
            'amount' => trim($arr[4]) * ($type == 2 ? -1 : 1),
            'trans_no' => trim($arr[5], " \t\n\r\0\x0B\""),
            'merchant_order_no' => trim($arr[6], " \t\n\r\0\x0B\""),
            'trans_time' => $trans_time
        ];
        Db::name('personbill')->insert($data);
    }
}

function loadAlipay()
{
    global $config;

    // 支付宝
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
    $reader->setInputEncoding('GBK');
    $reader->setDelimiter('`');
    $spreadsheet = $reader->load($config['alipay']['filename']);
    $sheet = $spreadsheet->getActiveSheet();
    for ($i=3; $i<=$sheet->getHighestRow() - 21; $i++) {
        $arr = explode(',', $sheet->getCell('A'.$i)->getValue());
        $type_str = trim($arr[0]);
        $type = $type_str == '收入' ? 1 : ($type_str == '支出' ? 2 : 0);

        $trans_time_str = trim($arr[10]);
        $trans_time = strtotime($trans_time_str);

        $data = [
            'account' => $config['alipay']['account'],
            'account_type_name' => '支付宝', // 账户类型
            'trans_type_name' => trim($arr[4]), // 交易方式
            'trans_person' => trim($arr[1]), // 交易对方
            'type' => $type, // 交易类型 1收入 2支出
            'description' => trim($arr[3]),
            'amount' => trim($arr[5]) * ($type == 2 ? -1 : 1),
            'trans_no' => trim($arr[8], " \t\n\r\0\x0B\""),
            'merchant_order_no' => trim($arr[9], " \t\n\r\0\x0B\""),
            'trans_time' => $trans_time
        ];
        try {
            Db::name('personbill')->insert($data);
        } catch (\think\db\exception\PDOException $e) {
            if ($e->getCode() == 10501) {
                continue;
            }

            throw $e;
        }
    }
}

// 微信
function loadWeChat()
{
    global $config;

    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
    $reader->setDelimiter('@');
    $spreadsheet = $reader->load($config['wechat']['filename']);
    $sheet = $spreadsheet->getActiveSheet();
    for ($i=18; $i<=$sheet->getHighestRow(); $i++) {
        $arr = explode(',', $sheet->getCell('A'.$i)->getValue());
        $type_str = trim($arr[4]);
        $type = $type_str == '收入' ? 1 : ($type_str == '支出' ? 2 : 0);

        $trans_time_str = trim($arr[0]);
        $trans_time = strtotime($trans_time_str);

        $data = [
            'account' => $config['wechat']['account'],
            'account_type_name' => '微信',
            'trans_type_name' => trim($arr[6]),
            'trans_person' => trim($arr[2]),
            'type' => $type,
            'description' => trim($arr[3], " \t\n\r\0\x0B\""),
            'amount' => trim($arr[5], " \t\n\r\0\x0B¥") * ($type == 2 ? -1 : 1),
            'trans_no' => trim($arr[8]),
            'merchant_order_no' => trim($arr[9]),
            'trans_time' => $trans_time
        ];
        try {
            Db::name('personbill')->insert($data);
        } catch (\think\db\exception\PDOException $e) {
            if ($e->getCode() == 10501) {
                continue;
            }

            throw $e;
        }
    }
}

// 建设银行
function loadCCB()
{
    global $config;

    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($config['ccb']['filename']);
    $sheet = $spreadsheet->getActiveSheet();
    for ($i=7; $i<=$sheet->getHighestRow() - 1; $i++) {
        $type = floatval(trim($sheet->getCell('D'.$i)->getValue())) > 0 ? 2 : 1;

        $trans_time_str = trim($sheet->getCell('B'.$i)->getValue()) . ' ' . trim($sheet->getCell('C'.$i)->getValue());
        $trans_time = strtotime($trans_time_str);

        Db::name('personbill')->insert([
            'account' => $config['ccb']['account'],
            'account_type_name' => '建设银行',
            'trans_type_name' => '银行卡',
            'trans_person' => trim($sheet->getCell('J'.$i)->getValue()),
            'type' => $type,
            'description' => trim($sheet->getCell('H'.$i)->getValue()),
            'amount' => $type == 2 ? -1 * $sheet->getCell('D'.$i)->getValue() : $sheet->getCell('E'.$i)->getValue(),
            'trans_time' => $trans_time
        ]);
    }
}

function main()
{
    global $config;

//    if (file_exists($config['ccb']['filename'])) {
//        loadCCB();
//    }
//    if (file_exists($config['alipay']['filename'])) {
//        loadAliPay();
//    }
    if (file_exists($config['wechat']['filename'])) {
        loadWeChat();
    }
}

main();
