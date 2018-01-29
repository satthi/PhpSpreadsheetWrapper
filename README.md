# PhpExcelWrapper

[![Build Status](https://travis-ci.org/satthi/PhpExcelWrapper.svg?branch=master)](https://travis-ci.org/satthi/PhpExcelWrapper)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/satthi/PhpExcelWrapper/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/satthi/PhpExcelWrapper/?branch=master)

このプロジェクトは[PHPExcel](https://github.com/PHPOffice/PHPExcel)を自分が使いやすいように対応したものになります。

## インストール
composer.json
```
{
	"require": {
		"satthi/phpexcelwrapper": "*"
	}
}
```

`composer install`

## 使い方(基本)

```php
<?php
require('./vendor/autoload.php');
use PhpExcelWrapper\PhpExcelWrapper;

class hoge{

    public function fuga(){
        $PhpExcelWrapper = new PhpExcelWrapper();
        //テンプレート使用の場合
        //$PhpExcelWrapper = new PhpExcelWrapper('./template.xlsx');
        $PhpExcelWrapper->setVal('設定したい値', 3, 1, 0);
        $PhpExcelWrapper->write('export.xlsx');
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
* @param integer $col 行 一番左は0
* @param integer $row 列 一番上は1
* @param integer $sheetNo シート番号 default 0
* @param integer $refCol 参照セル行 default null
* @param integer $refRow 参照セル列 default null
* @param integer $refSheet 参照シート default null
* @author hagiwara
*/
$PhpExcelWrapper->setVal('設定したい値', 3, 1, 0);

/**
* geVal
* 値の取得
* @param integer $col 行 一番左は0
* @param integer $row 列 一番上は1
* @param integer $sheetNo シート番号 default 0
* @author hagiwara
*/
$PhpExcelWrapper->getVal(3, 1, 0);

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
$PhpExcelWrapper->setImage('img/hoge.gif', 1, 1, 0);

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
$PhpExcelWrapper->cellMerge(0, 1, 0, 3, 0);

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
$PhpExcelWrapper->cellMerge(0, 1, 0, 0, 1, 1);

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
    'double' => PHPExcel_Style_Font::UNDERLINE_DOUBLE
    'doubleaccounting' => PHPExcel_Style_Font::UNDERLINE_DOUBLEACCOUNTING
    'none' => PHPExcel_Style_Font::UNDERLINE_NONE
    'single' => PHPExcel_Style_Font::UNDERLINE_SINGLE
    'singleaccounting' => PHPExcel_Style_Font::UNDERLINE_SINGLEACCOUNTING
    */
    'underline' => 'single',
    'bold' => true,
    'italic' => true,
    'strikethrough' => true,
    //ARGB
    'color' => 'FFFF0000',
    'size' => 40,
    /*
    //alignh パラメータリスト
    'general' => PHPExcel_Style_Alignment::HORIZONTAL_GENERAL,
    'center' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
    'left' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
    'right' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
    'justify' => PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY,
    'countinuous' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER_CONTINUOUS,
    */
    'alignh' => 'justify',
    /*
    //alignv パラメータリスト
    'bottom' => PHPExcel_Style_Alignment::VERTICAL_BOTTOM,
    'center' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
    'justify' => PHPExcel_Style_Alignment::VERTICAL_JUSTIFY,
    'top' => PHPExcel_Style_Alignment::VERTICAL_TOP,
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
    'none' => PHPExcel_Style_Border::BORDER_NONE,
    'thin' => PHPExcel_Style_Border::BORDER_THIN,
    'medium' => PHPExcel_Style_Border::BORDER_MEDIUM,
    'dashed' => PHPExcel_Style_Border::BORDER_DASHED,
    'dotted' => PHPExcel_Style_Border::BORDER_DOTTED,
    'thick' => PHPExcel_Style_Border::BORDER_THICK,
    'double' => PHPExcel_Style_Border::BORDER_DOUBLE,
    'hair' => PHPExcel_Style_Border::BORDER_HAIR,
    'mediumdashed' => PHPExcel_Style_Border::BORDER_MEDIUMDASHED,
    'dashdot' => PHPExcel_Style_Border::BORDER_DASHDOT,
    'mediumdashdot' => PHPExcel_Style_Border::BORDER_MEDIUMDASHDOT,
    'dashdotdot' => PHPExcel_Style_Border::BORDER_DASHDOTDOT,
    'mediumdashdotdot' => PHPExcel_Style_Border::BORDER_MEDIUMDASHDOTDOT,
    'slantdashdot' => PHPExcel_Style_Border::BORDER_SLANTDASHDOT,
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
    'linear' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
    'path' => PHPExcel_Style_Fill::FILL_GRADIENT_PATH,
    'none' => PHPExcel_Style_Fill::FILL_NONE,
    'darkdown' => PHPExcel_Style_Fill::FILL_PATTERN_DARKDOWN,
    'darkgray' => PHPExcel_Style_Fill::FILL_PATTERN_DARKGRAY,
    'darkgrid' => PHPExcel_Style_Fill::FILL_PATTERN_DARKGRID,
    'darkhorizontal' => PHPExcel_Style_Fill::FILL_PATTERN_DARKHORIZONTAL,
    'darktrellis' => PHPExcel_Style_Fill::FILL_PATTERN_DARKTRELLIS,
    'darkup' => PHPExcel_Style_Fill::FILL_PATTERN_DARKUP,
    'darkvertical' => PHPExcel_Style_Fill::FILL_PATTERN_DARKVERTICAL,
    'gray0625' => PHPExcel_Style_Fill::FILL_PATTERN_GRAY0625,
    'gray125' => PHPExcel_Style_Fill::FILL_PATTERN_GRAY125,
    'lightdown' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTDOWN,
    'lightgray' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTGRAY,
    'lightgrid' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTGRID,
    'lighthorizontal' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTHORIZONTAL,
    'lighttrellis' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTTRELLIS,
    'lightup' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTUP,
    'lightvertical' => PHPExcel_Style_Fill::FILL_PATTERN_LIGHTVERTICAL,
    'mediumgray' => PHPExcel_Style_Fill::FILL_PATTERN_MEDIUMGRAY,
    'solid' => PHPExcel_Style_Fill::FILL_SOLID,
    */
    'bgpattern' => 'lighthorizontal',
];
$PhpExcelWrapper->setStyle(3, 1, 0, $style);

/**
* createSheet
* シートの作成
* @param text $name
* @author hagiwara
*/
$PhpExcelWrapper->createSheet('hoge');

/**
* deleteSheet
* シートの削除
* @param integer $sheetNo
* @author hagiwara
*/
$PhpExcelWrapper->deleteSheet(4);

/**
* copySheet
* シートのコピー
* @param integer $sheetNo
* @param integer $position nullの場合は一番後ろ
* @param text $name
* @author hagiwara
*/
$PhpExcelWrapper->copySheet(0, null, 'copy sheet');

/**
* renameSheet
* シート名の変更
* @param integer $sheetNo
* @param text $name
* @author hagiwara
*/
$PhpExcelWrapper->renameSheet(0, 'rename');

/**
* write
* xlsxファイルの書き込み
* @param text $file 書き込み先のファイルパス
* @author hagiwara
*/
$PhpExcelWrapper->write('php://output');
```


```

## License ##

The MIT Lisence

Copyright (c) 2016 Fusic Co., Ltd. (http://fusic.co.jp)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Author ##

Satoru Hagiwara
