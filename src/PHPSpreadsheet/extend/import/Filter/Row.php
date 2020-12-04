<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Filter;


use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Font;

class Row
{
    /**
     * @title 检查行范围是否满足
     * @param int $currentRow
     * @param array $filterVArr
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:31
     */
    public static function checkRowRange(int $currentRow, array $filterVArr)
    {
        $filterItemFlag = true;    //默认为与保留
        foreach ($filterVArr as $range) {
            //如果是数组，代表是范围
            if (is_array($range) && $range[0] <= $currentRow && $currentRow <= $range[1]) {
                $filterItemFlag = false;
                break;
            }
        }
        return $filterItemFlag;
    }

}