# PhpSpreadsheetWrapper

[![Build Status](https://travis-ci.org/satthi/PhpSpreadsheetWrapper.svg?branch=master)](https://travis-ci.org/satthi/PhpSpreadsheetWrapper)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/satthi/PhpSpreadsheetWrapper/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/satthi/PhpSpreadsheetWrapper/?branch=master)

このプロジェクトは[PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet)を自分が使いやすいように対応したものになります。

## インストール
composer.json
```
{
	"require": {
		"satthi/phpspreadsheetwrapper": "*"
	}
}
```

`composer install`

## 使い方(基本)

```php
<?php
require('./vendor/autoload.php');
use PhpSpreadsheetWrapper\PhpSpreadsheetWrapper;

class hoge{

    public function fuga(){
        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper();
        //テンプレート使用の場合
        //$PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper('./template.xlsx');
        $PhpSpreadsheetWrapper->setVal('設定したい値', 3, 1, 0);
        $PhpSpreadsheetWrapper->write('export.xlsx');
    }
}

$hoge = new hoge();
$hoge->fuga();
```

## 準備している関数

```php
/**
* setVal
* 値のセット
* @param text $value 値
* @param integer $col 行 一番左は1
* @param integer $row 列 一番上は1
* @param integer $sheetNo シート番号 default 0
* @param integer $refCol 参照セル行 default null
* @param integer $refRow 参照セル列 default null
* @param integer $refSheet 参照シート default null
* @author hagiwara
*/
$PhpSpreadsheetWrapper->setVal('設定したい値', 3, 1, 0);

/**
* geVal
* 値の取得
* @param integer $col 行 一番左は1
* @param integer $row 列 一番上は1
* @param integer $sheetNo シート番号 default 0
* @author hagiwara
*/
$PhpSpreadsheetWrapper->getVal(3, 1, 0);

/**
* setImage
* 画像のセット
* @param text $img 画像のファイルパス
* @param integer $col 行
* @param integer $row 列
* @param integer $sheetNo シート番号 default 0
* @param integer $height 画像の縦幅 default null
* @param integer $width 画像の横幅 default null
* @param boolean $proportial 縦横比を維持するか default false
* @param integer $offsetx セルから何ピクセルずらすか（X軸) default null
* @param integer $offsety セルから何ピクセルずらすか（Y軸) default null
* @author hagiwara
*/
$PhpSpreadsheetWrapper->setImage('img/hoge.gif', 1, 1, 0);

/**
* cellMerge
* セルのマージ
* @param integer $col1 行
* @param integer $row1 列
* @param integer $col2 行
* @param integer $row2 列
* @param integer $sheetNo シート番号
* @author hagiwara
*/
$PhpSpreadsheetWrapper->cellMerge(1, 1, 1, 3, 0);

/**
* styleCopy
* セルの書式コピー
* @param integer $col 行
* @param integer $row 列
* @param integer $sheetNo シート番号
* @param integer $refCol 参照セル行
* @param integer $refRow 参照セル列
* @param integer $refSheet 参照シート
* @author hagiwara
*/
$PhpSpreadsheetWrapper->cellMerge(1, 1, 0, 1, 1, 1);

/**
* setStyle
* 書式のセット(まとめて)
* @param integer $col 行
* @param integer $row 列
* @param integer $sheetNo シート番号
* @param array $style スタイル情報
* @author hagiwara
*/
$style = [
    //フォント名
    'font' => 'HGP行書体',
    /*
    //underline パラメータリスト
    'double' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLE
    'doubleaccounting' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLEACCOUNTING
    'none' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE
    'single' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLE
    'singleaccounting' => \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLEACCOUNTING
    */
    'underline' => 'single',
    'bold' => true,
    'italic' => true,
    // 現状打消し線がうまく動作していない
    // 'strikethrough' => true,
    //ARGB
    'color' => 'FFFF0000',
    'size' => 40,
    /*
    //alignh パラメータリスト
    'general' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_GENERAL,
    'center' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    'left' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
    'right' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
    'justify' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_JUSTIFY,
    'countinuous' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER_CONTINUOUS,
    */
    'alignh' => 'justify',
    /*
    //alignv パラメータリスト
    'bottom' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM,
    'center' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
    'justify' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_JUSTIFY,
    'top' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP,
    */
    'alignv' => 'bottom',
    /*
    //罫線の位置
    'left' => null,
    'right' => null,
    'top' => null,
    'bottom' => null,
    'diagonal' => null,
    'all_borders' => null,
    'outline' => null,
    'inside' => null,
    'vertical' => null,
    'horizontal' => null,

    //罫線の種類
    'none' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE,
    'thin' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
    'medium' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
    'dashed' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHED,
    'dotted' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOTTED,
    'thick' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
    'double' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE,
    'hair' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR,
    'mediumdashed' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHED,
    'dashdot' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT,
    'mediumdashdot' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOT,
    'dashdotdot' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT,
    'mediumdashdotdot' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUMDASHDOTDOT,
    'slantdashdot' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_SLANTDASHDOT,
    */
    'border' => [
        'top' => [
            'type' => 'mediumdashed',
            'color' => 'FF664422'
        ],
        'bottom' => [
            'type' => 'double',
            'color' => 'FF224466'
        ],
    ],
    'bgcolor' => 'FF0000FF',
    /*
    //bgpattern パラメータリスト
    'linear' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
    'path' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_PATH,
    'none' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_NONE,
    'darkdown' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKDOWN,
    'darkgray' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKGRAY,
    'darkgrid' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKGRID,
    'darkhorizontal' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKHORIZONTAL,
    'darktrellis' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKTRELLIS,
    'darkup' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKUP,
    'darkvertical' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_DARKVERTICAL,
    'gray0625' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_GRAY0625,
    'gray125' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_GRAY125,
    'lightdown' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTDOWN,
    'lightgray' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTGRAY,
    'lightgrid' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTGRID,
    'lighthorizontal' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTHORIZONTAL,
    'lighttrellis' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTTRELLIS,
    'lightup' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTUP,
    'lightvertical' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_LIGHTVERTICAL,
    'mediumgray' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_PATTERN_MEDIUMGRAY,
    'solid' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
    */
    'bgpattern' => 'lighthorizontal',
];
$PhpSpreadsheetWrapper->setStyle(3, 1, 0, $style);

/**
* createSheet
* シートの作成
* @param text $name
* @author hagiwara
*/
$PhpSpreadsheetWrapper->createSheet('hoge');

/**
* deleteSheet
* シートの削除
* @param integer $sheetNo
* @author hagiwara
*/
$PhpSpreadsheetWrapper->deleteSheet(4);

/**
* copySheet
* シートのコピー
* @param integer $sheetNo
* @param integer $position nullの場合は一番後ろ
* @param text $name
* @author hagiwara
*/
$PhpSpreadsheetWrapper->copySheet(0, null, 'copy sheet');

/**
* renameSheet
* シート名の変更
* @param integer $sheetNo
* @param text $name
* @author hagiwara
*/
$PhpSpreadsheetWrapper->renameSheet(0, 'rename');

/**
* write
* xlsxファイルの書き込み
* @param text $file 書き込み先のファイルパス
* @author hagiwara
*/
$PhpSpreadsheetWrapper->write('php://output');
```


```

## License ##

The MIT Lisence

Copyright (c) 2018 Fusic Co., Ltd. (http://fusic.co.jp)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Author ##

Satoru Hagiwara
