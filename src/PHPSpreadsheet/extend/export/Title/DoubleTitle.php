<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Title;


use kxbrand\phplib\PHPSpreadsheet\core\Common;
use kxbrand\phplib\PHPSpreadsheet\core\Export;
use PhpOffice\PhpSpreadsheet\Exception;

class DoubleTitle
{
    private $export;
    private $activeSheet;

    public function __construct()
    {
        $this->export = Export::getInstance();
        $this->activeSheet =& $this->export->getActiveSheet();
    }

    /**
     * @title 设置二级表头
     * @param array $fieldArr
     * @param bool $replace
     * @return bool
     * @author millionmile
     * @time 2020/09/09 10:08
     */
    function setTitle(array $fieldArr, bool $replace): bool
    {
        //如果覆盖，那么只插入一行
        //不覆盖，则插入两行
        try {
            $this->export->insertRows(1, $replace ? 1 : 2);
            $colCount = 1;
            foreach ($fieldArr as $oneField => $fieldItem) {
                //要操作的开始列
                $startCol = Common::getColName($colCount);
                //如果是多维数组，那么代表里面是二维表头数据
                if (is_array($fieldItem)) {
                    $this->activeSheet->setCellValue($startCol . '1', $oneField);
                    foreach ($fieldItem as $twoField) {
                        $this->activeSheet->setCellValue(Common::getColName($colCount++) . '2', $twoField);
                    }
                    //要合并的结束列
                    $endCol = Common::getColName($colCount - 1);
                    $this->activeSheet->mergeCells($startCol . '1:' . $endCol . '1');
                } else {
                    $this->activeSheet->mergeCells($startCol . '1:' . $startCol . '2');
                    $this->activeSheet->setCellValue($startCol . '1', $fieldItem);
                    $colCount++;
                }
            }
        } catch (Exception $e) {
            return false;
        }
        return true;
    }

}