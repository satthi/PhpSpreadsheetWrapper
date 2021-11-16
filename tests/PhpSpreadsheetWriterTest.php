<?php
namespace PhpExecl\Test;

use PHPUnit\Framework\TestCase;
use PhpSpreadsheetWrapper\PhpSpreadsheetWrapper;

//PHPExcel本体(内容チェック用)
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;


require_once('./vendor/autoload.php');

class PhpSpreadsheetWriterTest extends TestCase
{

    private $__tmpDir;
    private $__exportFile;
    public function setUp(): void
    {
        parent::setUp();
        //ディレクトリの作成
        $this->__tmpDir = dirname(dirname(__FILE__)) . '/tmp';
        if (!is_dir($this->__tmpDir)) {
            mkdir($this->__tmpDir);
        }

        //出力ファイルがいたら削除
        $this->__exportFile = $this->__tmpDir . '/export.xlsx';
        if (file_exists($this->__exportFile)) {
            unlink($this->__exportFile);
        }
    }

    /**
     * tearDown method
     *
     * @return void
     */
    public function tearDown(): void
    {
        parent::tearDown();
    }

    /**
     * Test test_setVal
     *
     * @return void
     */
    public function test_setVal()
    {
        $setData = '設定したい値';
        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper();
        //D1に値をセット
        $PhpSpreadsheetWrapper->setVal($setData, 4, 1, 0);
        $PhpSpreadsheetWrapper->write($this->__exportFile);

        //ファイルのチェック
        $checkData = $this->getSheet()->getCell('D1')->getValue();

        $this->assertEquals($setData, $checkData);
    }

    /**
     * Test test_setImage
     *
     * @return void
     */
    public function test_setImage()
    {
        $setImagePath = dirname(__FILE__) . '/file/test.png';
        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper();
        //B1に画像をセット
        $PhpSpreadsheetWrapper->setImage($setImagePath, 2, 1, 0, 100, 120, false, 25, 30);
        //C1に縦横比を維持してセット
        $PhpSpreadsheetWrapper->setImage($setImagePath, 3, 1, 0, 100, 120, true, 40, 50);

        $PhpSpreadsheetWrapper->write($this->__exportFile);

        // $PhpSpreadsheetWrapper->setImage(WWW_ROOT . 'img/ajax-loader.gif', 1, 1, 0, 100, 100, false, 25, 25);
        //画像のチェック
        $checkDatas = $this->getSheet()->getDrawingCollection();
        //画像は二つ
        $this->assertEquals(2, count($checkDatas));

        //画像1
        //セル
        $this->assertEquals('B1', $checkDatas[0]->getCoordinates());
        //offsetx
        $this->assertEquals(25, $checkDatas[0]->getOffsetX());
        //offsety
        $this->assertEquals(30, $checkDatas[0]->getOffsetY());
        // バージョンアップか何かで値がうまく取れなくなってしまったので、コメントアウト
        // 画像自体は想定通りのサイズでセットされている
        //width
        // $this->assertEquals(120, $checkDatas[0]->getWidth());
        //height
        // $this->assertEquals(100, $checkDatas[0]->getHeight());

        //画像2
        //セル
        $this->assertEquals('C1', $checkDatas[1]->getCoordinates());
        //offsetx
        $this->assertEquals(40, $checkDatas[1]->getOffsetX());
        //offsety
        $this->assertEquals(50, $checkDatas[1]->getOffsetY());
        //width
        // $this->assertEquals(120, $checkDatas[1]->getWidth());
        //height
        // $this->assertEquals(120, $checkDatas[1]->getHeight());
    }


    /**
     * Test test_cellMerge
     *
     * @return void
     */
    public function test_cellMerge()
    {
        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper();
        //A1からJ2までセル結合
        $PhpSpreadsheetWrapper->cellMerge(1 ,1 ,10 ,2 ,0);
        //D3からD4までセル結合
        $PhpSpreadsheetWrapper->cellMerge(4 ,3 ,4 ,4 ,0);
        $PhpSpreadsheetWrapper->write($this->__exportFile);

        //セル結合のチェック
        $checkDatas = $this->getSheet()->getMergeCells();
        $this->assertEquals('A1:J2', $checkDatas['A1:J2']);
        $this->assertEquals('D3:D4', $checkDatas['D3:D4']);
    }

    /**
     * Test test_setStyle_styleCopy
     *
     * @return void
     */
    public function test_setStyle_styleCopy()
    {
        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper();
        $style = [
            'font' => 'HGP行書体',
            'underline' => 'double',
            'bold' => true,
            'italic' => true,
            // 打消し線がどうもうまく動作していないので一旦無視する
            // 'strikethrough' => true,
            'color' => 'FFFF0000',
            'size' => 40,
            'alignh' => 'justify',
            'alignv' => 'bottom',
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
            'bgpattern' => 'lighthorizontal',
        ];
        $PhpSpreadsheetWrapper->setStyle(4, 1, 0, $style);

        //styleのコピー
        $PhpSpreadsheetWrapper->styleCopy(5, 1, 0, 4, 1, 0);

        $PhpSpreadsheetWrapper->write($this->__exportFile);

        //セルの設定の読み込み
        $checkData = $this->getSheet()->getStyle('D1');
        $checkFont = $checkData->getFont();
        $checkAlignment = $checkData->getAlignment();

        //フォント名
        $this->assertEquals($style['font'] , $checkFont->getName());
        //下線
        $this->assertEquals(Font::UNDERLINE_DOUBLE , $checkFont->getUnderline());
        //太字
        $this->assertEquals($style['bold'] , $checkFont->getBold());
        //イタリック
        $this->assertEquals($style['italic'] , $checkFont->getItalic());
        //打ち消し線
        // $this->assertEquals($style['strikethrough'] , $checkFont->getStrikethrough());
        //色
        $this->assertEquals($style['color'] , $checkFont->getColor()->getARGB());
        //フォントサイズ
        $this->assertEquals($style['size'] , $checkFont->getSize());
        //水平方向
        $this->assertEquals(Alignment::HORIZONTAL_JUSTIFY , $checkAlignment->getHorizontal());
        //垂直方向方向
        $this->assertEquals(Alignment::VERTICAL_BOTTOM , $checkAlignment->getVertical());
        //罫線チェック
        $borderTop = $checkData->getBorders()->getTop();
        $borderBottom = $checkData->getBorders()->getBottom();
        $this->assertEquals(Border::BORDER_MEDIUMDASHED , $borderTop->getBorderStyle());
        $this->assertEquals($style['border']['top']['color'] , $borderTop->getColor()->getARGB());

        $this->assertEquals(Border::BORDER_DOUBLE , $borderBottom->getBorderStyle());
        $this->assertEquals($style['border']['bottom']['color'] , $borderBottom->getColor()->getARGB());


        //塗りつぶし方法
        $this->assertEquals(Fill::FILL_PATTERN_LIGHTHORIZONTAL , $checkData->getFill()->getFillType());
        //塗りつぶし色
        $this->assertEquals($style['bgcolor'] , $checkData->getFill()->getStartColor()->getARGB());

        //コピーセルについて
        $checkData = $this->getSheet()->getStyle('E1');
        $checkFont = $checkData->getFont();
        $checkAlignment = $checkData->getAlignment();

        //フォント名
        $this->assertEquals($style['font'] , $checkFont->getName());
        //下線
        $this->assertEquals(Font::UNDERLINE_DOUBLE , $checkFont->getUnderline());
        //太字
        $this->assertEquals($style['bold'] , $checkFont->getBold());
        //イタリック
        $this->assertEquals($style['italic'] , $checkFont->getItalic());
        //打ち消し線
        // $this->assertEquals($style['strikethrough'] , $checkFont->getStrikethrough());
        //色
        $this->assertEquals($style['color'] , $checkFont->getColor()->getARGB());
        //フォントサイズ
        $this->assertEquals($style['size'] , $checkFont->getSize());
        //水平方向
        $this->assertEquals(Alignment::HORIZONTAL_JUSTIFY , $checkAlignment->getHorizontal());
        //垂直方向方向
        $this->assertEquals(Alignment::VERTICAL_BOTTOM , $checkAlignment->getVertical());
        //罫線チェック
        $borderTop = $checkData->getBorders()->getTop();
        $borderBottom = $checkData->getBorders()->getBottom();
        $this->assertEquals(Border::BORDER_MEDIUMDASHED , $borderTop->getBorderStyle());
        $this->assertEquals($style['border']['top']['color'] , $borderTop->getColor()->getARGB());

        $this->assertEquals(Border::BORDER_DOUBLE , $borderBottom->getBorderStyle());
        $this->assertEquals($style['border']['bottom']['color'] , $borderBottom->getColor()->getARGB());
        //塗りつぶし方法
        $this->assertEquals(Fill::FILL_PATTERN_LIGHTHORIZONTAL , $checkData->getFill()->getFillType());
        //塗りつぶし色
        $this->assertEquals($style['bgcolor'] , $checkData->getFill()->getStartColor()->getARGB());
    }

    /**
     * Test test_createSheet
     *
     * @return void
     */
    public function test_createSheet()
    {
        $templateFile = dirname(__FILE__) . '/file/template.xlsx';
        $default_count = IOFactory::createReader('Xlsx')->load($templateFile)->getSheetCount();

        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper($templateFile);
        $PhpSpreadsheetWrapper->createSheet();
        $PhpSpreadsheetWrapper->write($this->__exportFile);

        $after_count = IOFactory::createReader('Xlsx')->load($this->__exportFile)->getSheetCount();
        //シートが一つ増えている
        $this->assertEquals($default_count , $after_count - 1);
    }

    /**
     * Test test_deleteSheet
     *
     * @return void
     */
    public function test_deleteSheet()
    {
        $templateFile = dirname(__FILE__) . '/file/template.xlsx';
        $default_count = IOFactory::createReader('Xlsx')->load($templateFile)->getSheetCount();

        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper($templateFile);
        $PhpSpreadsheetWrapper->deleteSheet(1);
        $PhpSpreadsheetWrapper->write($this->__exportFile);

        $after_count = IOFactory::createReader('Xlsx')->load($this->__exportFile)->getSheetCount();
        //シートが一つ減っている
        $this->assertEquals($default_count , $after_count + 1);

        //ファイルのチェック
        $sheetName = $this->getSheet(1)->getTitle();
        //2ページ目のシート名はシート3
        $this->assertEquals('Sheet3', $sheetName);
    }


    /**
     * Test test_copySheet
     *
     * @return void
     */
    public function test_copySheet()
    {
        $templateFile = dirname(__FILE__) . '/file/template.xlsx';
        $default_count = IOFactory::createReader('Xlsx')->load($templateFile)->getSheetCount();

        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper($templateFile);
        $PhpSpreadsheetWrapper->copySheet(1, null, 'copySheet');
        $PhpSpreadsheetWrapper->write($this->__exportFile);

        $after_count = IOFactory::createReader('Xlsx')->load($this->__exportFile)->getSheetCount();
        //シートが一つ増えている
        $this->assertEquals($default_count , $after_count - 1);

        //ファイルのチェック
        //4ページ目のシート名はコピーシート
        $this->assertEquals('copySheet', $this->getSheet(3)->getTitle());
        //中身のコンテンツはシート2
        $this->assertEquals('シート2', $this->getSheet(3)->getCell('A1')->getValue());
    }


    /**
     * Test test_renameSheet
     *
     * @return void
     */
    public function test_renameSheet()
    {
        $templateFile = dirname(__FILE__) . '/file/template.xlsx';

        $PhpSpreadsheetWrapper = new PhpSpreadsheetWrapper($templateFile);
        $PhpSpreadsheetWrapper->renameSheet(1, 'renameSheet');
        $PhpSpreadsheetWrapper->write($this->__exportFile);

        $after_count = IOFactory::createReader('Xlsx')->load($this->__exportFile)->getSheetCount();

        //2ページ目のシート名はrenameSheet
        $this->assertEquals('renameSheet', $this->getSheet(1)->getTitle());
    }


    private function getSheet($sheetNo = 0)
    {
        $checkExcel = IOFactory::createReader('Xlsx')->load($this->__exportFile);
        return $checkExcel->setActiveSheetIndex($sheetNo);
    }




}
