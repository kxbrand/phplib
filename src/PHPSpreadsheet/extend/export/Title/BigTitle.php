<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Title;

use kxbrand\phplib\PHPSpreadsheet\core\Common;
use kxbrand\phplib\PHPSpreadsheet\core\Export;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Style\StyleConfig;
use PhpOffice\PhpSpreadsheet\Exception;

class BigTitle
{
    private $export;
    private $mergeCount;
    private $styleObj;
    private $defaultStyle;

    public function __construct(int $mergeCount = 1)
    {
        $this->export = Export::getInstance();
        $this->mergeCount = $mergeCount;

        //设置默认样式
        $this->defaultStyle = [
            'bold' => true,
            'font-size' => ''
        ];
    }

    /**
     * @title 人工设置样式对象
     * @param StyleConfig $styleObj
     * @author millionmile
     * @time 2020/07/06 12:09
     */
    public function setStyleObj(StyleConfig $styleObj)
    {
        $this->styleObj = $styleObj;
    }


    /**
     * @title 获取样式对象，如果不存在，则自己创建
     * @author millionmile
     * @time 2020/07/06 12:10
     */
    private function getStyleObj()
    {

    }


    public function setBigTitle(string $titleStr)
    {
        $this->getStyleObj();
        $sheet =& $this->export->getActiveSheet();
        //插入一行
        try {
            $this->export->insertRows(1, 1);
            //在首行加入大标题
            $sheet->setCellValue('A1', $titleStr);
            $endCol = Common::getColName($this->mergeCount);
            $sheet->mergeCells('A1:' . $endCol . '1');
        } catch (Exception $e) {
            return $e->getMessage();
        }
        return true;
    }
}