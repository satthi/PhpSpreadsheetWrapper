<?php

namespace PhpSpreadsheetWrapper;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

/**
* PhpSpreadsheetWrapper
* Spreadsheetを記載しやすくするためのラッパー
*/
class PhpSpreadsheetWrapper
{
    private $__phpexcel;
    private $__sheet = [];
    private $__deleteSheetList = [];
    private static $__underlineType = [
        'double' => Font::UNDERLINE_DOUBLE,
        'doubleaccounting' => Font::UNDERLINE_DOUBLEACCOUNTING,
        'none' => Font::UNDERLINE_NONE,
        'single' => Font::UNDERLINE_SINGLE,
        'singleaccounting' => Font::UNDERLINE_SINGLEACCOUNTING,
    ];

    private static $__borderType = [
        'none' => Border::BORDER_NONE,
        'thin' => Border::BORDER_THIN,
        'medium' => Border::BORDER_MEDIUM,
        'dashed' => Border::BORDER_DASHED,
        'dotted' => Border::BORDER_DOTTED,
        'thick' => Border::BORDER_THICK,
        'double' => Border::BORDER_DOUBLE,
        'hair' => Border::BORDER_HAIR,
        'mediumdashed' => Border::BORDER_MEDIUMDASHED,
        'dashdot' => Border::BORDER_DASHDOT,
        'mediumdashdot' => Border::BORDER_MEDIUMDASHDOT,
        'dashdotdot' => Border::BORDER_DASHDOTDOT,
        'mediumdashdotdot' => Border::BORDER_MEDIUMDASHDOTDOT,
        'slantdashdot' => Border::BORDER_SLANTDASHDOT,
    ];

    private static $__alignHolizonalType = [
        'general' => Alignment::HORIZONTAL_GENERAL,
        'center' => Alignment::HORIZONTAL_CENTER,
        'left' => Alignment::HORIZONTAL_LEFT,
        'right' => Alignment::HORIZONTAL_RIGHT,
        'justify' => Alignment::HORIZONTAL_JUSTIFY,
        'countinuous' => Alignment::HORIZONTAL_CENTER_CONTINUOUS,
    ];

    private static $__alignVerticalType = [
        'bottom' => Alignment::VERTICAL_BOTTOM,
        'center' => Alignment::VERTICAL_CENTER,
        'justify' => Alignment::VERTICAL_JUSTIFY,
        'top' => Alignment::VERTICAL_TOP,
    ];

    private static $__fillType = [
        'linear' => Fill::FILL_GRADIENT_LINEAR,
        'path' => Fill::FILL_GRADIENT_PATH,
        'none' => Fill::FILL_NONE,
        'darkdown' => Fill::FILL_PATTERN_DARKDOWN,
        'darkgray' => Fill::FILL_PATTERN_DARKGRAY,
        'darkgrid' => Fill::FILL_PATTERN_DARKGRID,
        'darkhorizontal' => Fill::FILL_PATTERN_DARKHORIZONTAL,
        'darktrellis' => Fill::FILL_PATTERN_DARKTRELLIS,
        'darkup' => Fill::FILL_PATTERN_DARKUP,
        'darkvertical' => Fill::FILL_PATTERN_DARKVERTICAL,
        'gray0625' => Fill::FILL_PATTERN_GRAY0625,
        'gray125' => Fill::FILL_PATTERN_GRAY125,
        'lightdown' => Fill::FILL_PATTERN_LIGHTDOWN,
        'lightgray' => Fill::FILL_PATTERN_LIGHTGRAY,
        'lightgrid' => Fill::FILL_PATTERN_LIGHTGRID,
        'lighthorizontal' => Fill::FILL_PATTERN_LIGHTHORIZONTAL,
        'lighttrellis' => Fill::FILL_PATTERN_LIGHTTRELLIS,
        'lightup' => Fill::FILL_PATTERN_LIGHTUP,
        'lightvertical' => Fill::FILL_PATTERN_LIGHTVERTICAL,
        'mediumgray' => Fill::FILL_PATTERN_MEDIUMGRAY,
        'solid' => Fill::FILL_SOLID,
    ];

    /**
    * __construct
    *
    * @param string $template テンプレートファイルのパス
    * @author hagiwara
    */
    public function __construct($template = null, $type = 'Xlsx')
    {
        if ($template === null) {
            //テンプレート無し
            $this->__phpexcel = new Spreadsheet();
        } else {
            //テンプレートの読み込み
            $reader = IOFactory::createReader($type);
            $this->__phpexcel = $reader->load($template);
        }
    }

    /**
    * setVal
    * 値のセット
    * @param string $value 値
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param integer $refCol 参照セル行
    * @param integer $refRow 参照セル列
    * @param integer $refSheet 参照シート
    * @param string $dataType 書き込みするデータの方を強制的に指定
    * @author hagiwara
    */
    public function setVal($value, $col, $row, $sheetNo = 0, $refCol = null, $refRow = null, $refSheet = 0, $dataType = null)
    {
        $cellInfo = $this->cellInfo($col, $row);
        //値のセット
        if (is_null($dataType)) {
            $this->getSheet($sheetNo)->setCellValue($cellInfo, $value);
        } else {
            $this->getSheet($sheetNo)->getCell($cellInfo)->setValueExplicit($value, $dataType);
        }

        //参照セルの指定がある場合には書式をコピーする
        if (!is_null($refCol) && !is_null($refRow)) {
            $this->styleCopy($col, $row, $sheetNo, $refCol, $refRow, $refSheet);
        }
    }

    /**
    * getVal
    * 値の取得
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @author hagiwara
    */
    public function getVal($col, $row, $sheetNo = 0)
    {
        $cellInfo = $this->cellInfo($col, $row);
        //値の取得
        return $this->getSheet($sheetNo)->getCell($cellInfo)->getValue();
    }

    /**
    * getPlainVal
    * 値の取得
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @author hagiwara
    */
    public function getPlainVal($col, $row, $sheetNo = 0)
    {
        $val = $this->getVal($col, $row, $sheetNo);

        // RichTextが来たときにplainのtextを返すため
        if (is_object($val) && is_a($val, '\PhpOffice\PhpSpreadsheet\RichText\RichText')) {
            $val = $val-> getPlainText();
        }
        return $val;
    }

    /**
    * setImage
    * 画像のセット
    * @param string $img 画像のファイルパス
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param integer $height 画像の縦幅
    * @param integer $width 画像の横幅
    * @param boolean $proportial 縦横比を維持するか
    * @param integer $offsetx セルから何ピクセルずらすか（X軸)
    * @param integer $offsety セルから何ピクセルずらすか（Y軸)
    * @author hagiwara
    */
    public function setImage($img, $col, $row, $sheetNo = 0, $height = null, $width = null, $proportial = false, $offsetx = null, $offsety = null)
    {
        $cellInfo = $this->cellInfo($col, $row);

        $objDrawing = new Drawing();//画像用のオプジェクト作成
        $objDrawing->setPath($img);//貼り付ける画像のパスを指定
        $objDrawing->setCoordinates($cellInfo);//位置
        if (!is_null($proportial)) {
            $objDrawing->setResizeProportional($proportial);//縦横比の変更なし
        }
        if (!is_null($height)) {
            $objDrawing->setHeight($height);//画像の高さを指定
        }
        if (!is_null($width)) {
            $objDrawing->setWidth($width);//画像の高さを指定
        }
        if (!is_null($offsetx)) {
            $objDrawing->setOffsetX($offsetx);//指定した位置からどれだけ横方向にずらすか。
        }
        if (!is_null($offsety)) {
            $objDrawing->setOffsetY($offsety);//指定した位置からどれだけ縦方向にずらすか。
        }
        $objDrawing->setWorksheet($this->getSheet($sheetNo));
    }

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
    public function cellMerge($col1, $row1, $col2, $row2, $sheetNo)
    {
        $cell1Info = $this->cellInfo($col1, $row1);
        $cell2Info = $this->cellInfo($col2, $row2);

        $this->getSheet($sheetNo)->mergeCells($cell1Info . ':' . $cell2Info);
    }


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
    public function styleCopy($col, $row, $sheetNo, $refCol, $refRow, $refSheet)
    {
        $cellInfo = $this->cellInfo($col, $row);
        $refCellInfo = $this->cellInfo($refCol, $refRow);
        $style = $this->getSheet($refSheet)->getStyle($refCellInfo);

        $this->getSheet($sheetNo)->duplicateStyle($style, $cellInfo);
    }

    /**
    * setStyle
    * 書式のセット(まとめて)
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param array $style スタイル情報
    * @author hagiwara
    */
    public function setStyle($col, $row, $sheetNo, $style)
    {
        $default_style = [
            'font' => null,
            'underline' => null,
            'bold' => null,
            'italic' => null,
            'strikethrough' => null,
            'color' => null,
            'size' => null,
            'alignh' => null,
            'alignv' => null,
            'bgcolor' => null,
            'bgpattern' => null,
            'border' => null,
        ];
        $style = array_merge($default_style, $style);
        $this->setFontName($col, $row, $sheetNo, $style['font']);
        $this->setUnderline($col, $row, $sheetNo, $style['underline']);
        $this->setFontBold($col, $row, $sheetNo, $style['bold']);
        $this->setItalic($col, $row, $sheetNo, $style['italic']);
        $this->setStrikethrough($col, $row, $sheetNo, $style['strikethrough']);
        $this->setColor($col, $row, $sheetNo, $style['color']);
        $this->setSize($col, $row, $sheetNo, $style['size']);
        $this->setAlignHolizonal($col, $row, $sheetNo, $style['alignh']);
        $this->setAlignVertical($col, $row, $sheetNo, $style['alignv']);
        $this->setBackgroundColor($col, $row, $sheetNo, $style['bgcolor'], $style['bgpattern']);
        $this->setBorder($col, $row, $sheetNo, $style['border']);
    }

    /**
    * setFontName
    * フォントのセット
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param string|null $fontName フォント名
    * @author hagiwara
    */
    public function setFontName($col, $row, $sheetNo, $fontName)
    {
        if (is_null($fontName)) {
            return;
        }
        $this->getFont($col, $row, $sheetNo)->setName($fontName);
    }

    /**
    * setUnderline
    * 下線のセット
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param string|null $underline 下線の種類
    * @author hagiwara
    */
    public function setUnderline($col, $row, $sheetNo, $underline)
    {
        if (is_null($underline)) {
            return;
        }
        $this->getFont($col, $row, $sheetNo)->setUnderline($underline);
    }


    /**
    * getUnderlineType
    * 下線の種類の設定
    * @param string $type
    * @author hagiwara
    */
    private function getUnderlineType($type)
    {
        $type_list = self::$__underlineType;
        if (array_key_exists($type, $type_list)) {
            return $type_list[$type];
        }
        return Border::UNDERLINE_NONE;
    }

    /**
    * setFontBold
    * 太字のセット
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param boolean|null $bold 太字を引くか
    * @author hagiwara
    */
    public function setFontBold($col, $row, $sheetNo, $bold)
    {
        if (is_null($bold)) {
            return;
        }
        $this->getFont($col, $row, $sheetNo)->setBold($bold);
    }

    /**
    * setItalic
    * イタリックのセット
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param boolean|null $italic イタリックにするか
    * @author hagiwara
    */
    public function setItalic($col, $row, $sheetNo, $italic)
    {
        if (is_null($italic)) {
            return;
        }
        $this->getFont($col, $row, $sheetNo)->setItalic($italic);
    }

    /**
    * setStrikethrough
    * 打ち消し線のセット
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param boolean|null $strikethrough 打ち消し線をつけるか
    * @author hagiwara
    */
    public function setStrikethrough($col, $row, $sheetNo, $strikethrough)
    {
        if (is_null($strikethrough)) {
            return;
        }
        var_dump($col);
        var_dump($row);
        var_dump($sheetNo);
        var_dump($strikethrough);
        $this->getFont($col, $row, $sheetNo)->setStrikethrough($strikethrough);
    }

    /**
    * setColor
    * 文字の色
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param string|null $color 色(ARGB)
    * @author hagiwara
    */
    public function setColor($col, $row, $sheetNo, $color)
    {
        if (is_null($color)) {
            return;
        }
        $this->getFont($col, $row, $sheetNo)->getColor()->setARGB($color);
    }

    /**
    * setSize
    * 文字サイズ
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param integer|null $size
    * @author hagiwara
    */
    public function setSize($col, $row, $sheetNo, $size)
    {
        if (is_null($size)) {
            return;
        }
        $this->getFont($col, $row, $sheetNo)->setSize($size);
    }

    /**
    * getFont
    * fontデータ取得
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @author hagiwara
    */
    private function getFont($col, $row, $sheetNo)
    {
        $cellInfo = $this->cellInfo($col, $row);
        return $this->getSheet($sheetNo)->getStyle($cellInfo)->getFont();
    }

    /**
    * setAlignHolizonal
    * 水平方向のalign
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param string|null $type
    * typeはgetAlignHolizonalType参照
    * @author hagiwara
    */
    public function setAlignHolizonal($col, $row, $sheetNo, $type)
    {
        if (is_null($type)) {
            return;
        }
        $this->getAlignment($col, $row, $sheetNo)->setHorizontal($this->getAlignHolizonalType($type));
    }

    /**
    * setAlignVertical
    * 垂直方法のalign
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param string|null $type
    * typeはgetAlignVerticalType参照
    * @author hagiwara
    */
    public function setAlignVertical($col, $row, $sheetNo, $type)
    {
        if (is_null($type)) {
            return;
        }
        $this->getAlignment($col, $row, $sheetNo)->setVertical($this->getAlignVerticalType($type));
    }

    /**
    * getAlignment
    * alignmentデータ取得
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @author hagiwara
    */
    private function getAlignment($col, $row, $sheetNo)
    {
        $cellInfo = $this->cellInfo($col, $row);
        return $this->getSheet($sheetNo)->getStyle($cellInfo)->getAlignment();
    }

    /**
    * setBorder
    * 罫線の設定
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param array|null $border
    * borderの内部はgetBorderType参照
    * @author hagiwara
    */
    public function setBorder($col, $row, $sheetNo, $border)
    {
        if (is_null($border)) {
            return;
        }
        $cellInfo = $this->cellInfo($col, $row);
        $default_border = [
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
        ];
        $border = array_merge($default_border, $border);
        foreach ($border as $border_position => $border_setting) {
            if (!is_null($border_setting)) {
                $borderInfo =  $this->getSheet($sheetNo)->getStyle($cellInfo)->getBorders()->{'get' . $this->camelize($border_position)}();
                if (array_key_exists('type', $border_setting)) {
                    $borderInfo->setBorderStyle($this->getBorderType($border_setting['type']));
                }
                if (array_key_exists('color', $border_setting)) {
                    $borderInfo->getColor()->setARGB($border_setting['color']);
                }
            }
        }
    }

    /**
    * setBackgroundColor
    * 背景色の設定
    * @param integer $col 行
    * @param integer $row 列
    * @param integer $sheetNo シート番号
    * @param string $color 色
    * @param string $fillType 塗りつぶし方(デフォルトsolid)
    * fillTypeの内部はgetFillType参照
    * @author hagiwara
    */
    public function setBackgroundColor($col, $row, $sheetNo, $color, $fillType = 'solid')
    {
        $cellInfo = $this->cellInfo($col, $row);

        $this->getSheet($sheetNo)->getStyle($cellInfo)->getFill()->setFillType($this->getFillType($fillType))->getStartColor()->setARGB($color);
    }

    /**
    * getBorderType
    * 罫線の種類の設定
    * @param string $type
    * @author hagiwara
    */
    private function getBorderType($type)
    {
        $type_list = self::$__borderType;
        if (array_key_exists($type, $type_list)) {
            return $type_list[$type];
        }
        return Border::BORDER_NONE;
    }

    /**
    * getAlignHolizonalType
    * 水平方向のAlignの設定
    * @param string $type
    * @author hagiwara
    */
    private function getAlignHolizonalType($type)
    {
        $type_list = self::$__alignHolizonalType;
        if (array_key_exists($type, $type_list)) {
            return $type_list[$type];
        }
        return Alignment::HORIZONTAL_GENERAL;
    }

    /**
    * getAlignVerticalType
    * 垂直方向のAlignの設定
    * @param string $type
    * @author hagiwara
    */
    private function getAlignVerticalType($type)
    {
        $type_list = self::$__alignVerticalType;
        if (array_key_exists($type, $type_list)) {
            return $type_list[$type];
        }
        return null;
    }

    /**
    * getFillType
    * 塗りつぶしの設定
    * @param string $type
    * @author hagiwara
    */
    private function getFillType($type)
    {
        $type_list = self::$__fillType;
        if (array_key_exists($type, $type_list)) {
            return $type_list[$type];
        }
        return Fill::FILL_SOLID;
    }

    /**
    * createSheet
    * シートの作成
    * @param string $name
    * @author hagiwara
    */
    public function createSheet($name = null)
    {
        //シートの新規作成
        $newSheet = $this->__phpexcel->createSheet();
        $sheetNo = $this->__phpexcel->getIndex($newSheet);
        $this->__sheet[$sheetNo] = $newSheet;
        if (!is_null($name)) {
            $this->renameSheet($sheetNo, $name);
        }
    }

    /**
    * deleteSheet
    * シートの削除
    * @param integer $sheetNo
    * @author hagiwara
    */
    public function deleteSheet($sheetNo)
    {
        //シートの削除は一番最後に行う
        $this->__deleteSheetList[] = $sheetNo;
    }

    /**
    * copySheet
    * シートのコピー
    * @param integer $sheetNo
    * @param integer $position
    * @param string $name
    * @author hagiwara
    */
    public function copySheet($sheetNo, $position = null, $name = null)
    {
        $base = $this->getSheet($sheetNo)->copy();
        if ($name === null) {
            $name = uniqid();
        }
        $base->setTitle($name);

        // $positionが null(省略時含む)の場合は最後尾に追加される
        $this->__phpexcel->addSheet($base, $position);
    }

    /**
    * renameSheet
    * シート名の変更
    * @param integer $sheetNo
    * @param string $name
    * @author hagiwara
    */
    public function renameSheet($sheetNo, $name)
    {
        $this->getSheet($sheetNo)->setTitle($name);
    }

    /**
    * write
    * xlsxファイルの書き込み
    * @param string $file 書き込み先のファイルパス
    * @author hagiwara
    */
    public function write($file, $type = 'Xlsx')
    {
        //書き込み前に削除シートを削除する
        foreach ($this->__deleteSheetList as $deleteSheet) {
            $this->__phpexcel->removeSheetByIndex($deleteSheet);
        }
        $writer = IOFactory::createWriter($this->__phpexcel, $type);
        $writer->save($file);
    }

    /**
    * getReader
    * readerを返す(※直接Spreadsheetの関数を実行できるように)
    * @author hagiwara
    */
    public function getReader()
    {
        return $this->__phpexcel;
    }

    /**
    * getSheet
    * シート情報の読み込み
    * @param integer $sheetNo シート番号
    * @author hagiwara
    * @return null|\Spreadsheet_Worksheet
    */
    private function getSheet($sheetNo)
    {
        if (!array_key_exists($sheetNo, $this->__sheet)) {
            $this->__sheet[$sheetNo] = $this->__phpexcel->setActiveSheetIndex($sheetNo);
        }
        return $this->__sheet[$sheetNo];
    }

    /**
    * cellInfo
    * R1C1参照をA1参照に変換して返す
    * @param integer $col 行
    * @param integer $row 列
    * @author hagiwara
    */
    private function cellInfo($col, $row)
    {
        $stringCol = Coordinate::stringFromColumnIndex($col);
        return $stringCol . $row;
    }

    /**
    * cellInfo
    * http://qiita.com/Hiraku/items/036080976884fad1e450
    * @param string $str
    */
    private function camelize($str)
    {
        $str = ucwords($str, '_');
        return str_replace('_', '', $str);
    }
}
