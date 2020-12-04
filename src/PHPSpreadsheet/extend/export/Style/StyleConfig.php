<?php

namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Style;

use kxbrand\phplib\PHPSpreadsheet\core\Export;

class StyleConfig
{
    private $export;
    private $sheet;

    public function __construct()
    {
        $this->export = Export::getInstance();
        $this->sheet =& $this->export->getActiveSheet();
    }


    /**
     * @title 设置行高
     * @param int $rowNum
     * @param int $rowHeight
     * @author millionmile
     * @time 2020/07/06 16:21
     */
    public function setRowHeight(int $rowNum, int $rowHeight)
    {
        $this->sheet->getRowDimension($rowNum)->setRowHeight($rowHeight);
        return $this;
    }

    /**
     * @title 设置列宽
     * @param string $colName
     * @param int $width
     * @author millionmile
     * @time 2020/07/06 16:25
     */
    public function setColumnWidth(string $colName, int $width = 0)
    {
        if ($width === 0) {
            //如果希望PhpSpreadsheet执行自动宽度计算
            $this->sheet->getColumnDimension($colName)->setAutoSize(true);
        } else {
            $this->sheet->getColumnDimension($colName)->setWidth($width);
        }
    }

    /**
     * @title 设置单元格样式
     * @param string $cellRange
     * @param array $styleArr
     * @return $this
     * @author millionmile
     * @time 2020/07/06 16:23
     */
    public function setCellStyle(string $cellRange, array $styleArr)
    {
        $this->sheet->getStyle($cellRange)->applyFromArray($styleArr);
        return $this;
    }

    /**
     * @title 设置列样式
     * @author millionmile
     * @time 2020/07/07 10:31
     */
    public function setColStyle(string $colName, array $styleArr)
    {
        //设置列宽
        if (isset($styleArr['width'])) {
            if ($styleArr['width'] === 'auto') {
                $this->sheet->getColumnDimension($colName)->setAutoSize(true);
            } elseif (is_numeric($styleArr['width'])) {
                $this->sheet->getColumnDimension($colName)->setWidth($styleArr['width']);
            }
        }
    }


    /**
     * @title 设置列样式
     * @author millionmile
     * @time 2020/07/07 10:31
     */
    public function setRowStyle($rowNum, array $styleArr)
    {
        //设置列宽
        if (isset($styleArr['height'])) {
            if (strpos($rowNum, ':') !== false) {
                list($startNum, $endNum) = explode(':', $rowNum);
                for ($i = $startNum; $i <= $endNum; $i++) {
                    $this->setRowHeight($i, $styleArr['height']);
                }
            } else {
                $this->setRowHeight($rowNum, $styleArr['height']);
            }
        }
    }
}