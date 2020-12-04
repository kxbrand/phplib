<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Header;


use kxbrand\phplib\PHPSpreadsheet\core\Common;
use kxbrand\phplib\PHPSpreadsheet\core\Export;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Style\StyleBuilder;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Style\StyleConfig;

class HeaderStyle
{
    private $styleConfig;
    private $styleBuilder;
    private $export;

    public function __construct()
    {
        $this->styleConfig = new StyleConfig();
        $this->styleBuilder = new StyleBuilder();
        $this->export = Export::getInstance();
    }

    private function analysisStyleRange(int $colNum = 0, string $range = '')
    {
        if ($colNum === 0) {
            $colName = '';
        } else {
            $colName = Common::getColName($colNum);
        }

        if (empty($range)) {
            $currentRow = $this->export->getTheRow();
            return [$colName . '1:' . $colName . $currentRow];
        } else {
            //最终应用样式的列范围数组
            $finalRangeArr = [];
            $rangeArr = explode(',', $range);
            foreach ($rangeArr as $rangeItem) {
                //如果不为空
                if (strpos($rangeItem, '-') !== false) {
                    list($startRow, $endRow) = explode('-', $rangeItem);
                    $finalRangeArr[] = $colName . $startRow . ':' . $colName . $endRow;
                } else {
                    $finalRangeArr[] = $colName . $rangeItem;
                }
            }
            return $finalRangeArr;
        }
    }

    /**
     * @title 分析单元格样式，最后生成可用的样式数组
     * @param array $styleArr
     * @return array
     * @author millionmile
     * @time 2020/07/07 11:08
     */
    private function analysisCellStyle(array $styleArr): array
    {
        $this->styleBuilder->init();
        foreach ($styleArr as $style => $value) {
            switch ($style) {
                case 'font_bold':   //加粗
                    $this->styleBuilder->setFontBold($value);
                    break;
                case 'font_italic': //倾斜
                    $this->styleBuilder->setFontItalic($value);
                    break;
                case 'font_underline':  //下划线
                    $this->styleBuilder->setFontUnderline($value);
                    break;
                case 'font':    //字体
                    $this->styleBuilder->setFont($value);
                    break;
                case 'font_size':   //字体大小
                    $this->styleBuilder->setFontSize($value);
                    break;
                case 'font_color':  //字体颜色
                    $this->styleBuilder->setFontColor($value);
                    break;
                case 'horizontal':  //水平排列
                    $this->styleBuilder->setAlignmentHorizontal($value);
                    break;
                case 'vertical':    //垂直排列
                    $this->styleBuilder->setAlignmentVertical($value);
                    break;
                case 'wrap_text':   //自动换行
                    $this->styleBuilder->setWrapText($value);
                    break;
                case 'shrink_to_fit':   //缩小以适应
                    $this->styleBuilder->setShrinkToFit($value);
                    break;
                case 'text_indent': //首行缩进
                    $this->styleBuilder->setTextIndent($value);
                    break;
                case 'border':  //边框
                    if (is_array($value)) {
                        $this->styleBuilder->setBorder($value[0] ?? '', $value[1] ?? '', $value[2] ?? '');
                    } else {
                        $this->styleBuilder->setBorder($value);
                    }
                    break;
                case 'diagonal':    //对角线
                    if (is_array($value)) {
                        $this->styleBuilder->setDiagonal($value[0] ?? '', $value[1] ?? '', $value[2] ?? '');
                    } else {
                        $this->styleBuilder->setDiagonal($value);
                    }
                    break;
                case 'background':  //设置背景色
                    $this->styleBuilder->setBackgroundColor($value);
                    break;
            }
        }
        return $this->styleBuilder->build()['cellStyle'];
    }


    /**
     * @title 分析列样式，最后生成可用的样式数组
     * @param array $styleArr
     * @return array
     * @author millionmile
     * @time 2020/07/07 11:08
     */
    private function analysisColStyle(array $styleArr): array
    {
        $this->styleBuilder->init();
        foreach ($styleArr as $style => &$value) {
            switch ($style) {
                case 'width':  //设置列宽
                    $this->styleBuilder->setWidth($value);
                    break;
            }
        }
        unset($value);
        return $this->styleBuilder->build()['colStyle'];
    }


    /**
     * @title 分析行样式，最后生成可用的样式数组
     * @param array $styleArr
     * @return array
     * @author millionmile
     * @time 2020/07/07 11:08
     */
    private function analysisRowStyle(array $styleArr): array
    {
        $this->styleBuilder->init();
        foreach ($styleArr as $style => $value) {
            switch ($style) {
                case 'height':  //设置背景色
                    $this->styleBuilder->setHeight($value);
                    break;
            }
        }
        return $this->styleBuilder->build()['rowStyle'];
    }


    /**
     * @title 设置样式（总）
     * @param array $styleArr
     * @author millionmile
     * @time 2020/07/07 11:10
     */
    public function setStyle(array $styleArr)
    {
        foreach ($styleArr as $colNum => &$originStyleItem) {
            $styleItemArr = $this->decodeHeaderStyleData($originStyleItem);
            if (empty($styleItemArr)) {
                continue;
            }

            //设置单元格样式
            $this->setCellStyle($colNum, $styleItemArr);

            //设置列样式
            $this->setColStyle($colNum, $styleItemArr);

            //设置行样式
            $this->setRowStyle($styleItemArr);
        }
    }

    /**
     * @title 设置单元格格式
     * @param int $colNum
     * @param array $styleItemArr
     * @author millionmile
     * @time 2020/07/07 11:14
     */
    private function setCellStyle(int $colNum, array $styleItemArr)
    {
        $finalStyleArr = [];
        foreach ($styleItemArr as &$styleItem) {
            $analysisedStyleArr = $this->analysisCellStyle($styleItem['data']);
            //如果没有样式，跳到下一个
            if (empty($analysisedStyleArr)) {
                continue;
            }
            //分析作用范围
            $rangeArr = $this->analysisStyleRange($colNum, $styleItem['range'] ?? '');
            foreach ($rangeArr as $range) {
                if (!isset($finalStyleArr[$range])) {
                    $finalStyleArr[$range] = $analysisedStyleArr;
                } else {
                    $finalStyleArr[$range] = array_merge_recursive($finalStyleArr[$range], $analysisedStyleArr);
                }
            }
        }
        unset($styleItem);

        //分析单元格样式
        foreach ($finalStyleArr as $range => $finalStyleItem) {
            $this->styleConfig->setCellStyle($range, $finalStyleItem);
        }
    }


    /**
     * @title 设置列格式
     * @param int $colNum
     * @param array $styleItemArr
     * @author millionmile
     * @time 2020/07/07 11:14
     */
    private function setColStyle(int $colNum, array $styleItemArr)
    {
        foreach ($styleItemArr as $i => &$styleItem) {
            //分析作用范围
            //分析单元格样式
            $analysisedStyleArr = $this->analysisColStyle($styleItem['data']);
            if (empty($analysisedStyleArr)) {
                continue;
            }
            //设置单元格样式
            $this->styleConfig->setColStyle(Common::getColName($colNum), $analysisedStyleArr);
        }
        unset($styleItem);
    }


    /**
     * @title 设置行格式
     * @param array $styleArr
     * @author millionmile
     * @time 2020/07/07 11:15
     */
    private function setRowStyle(array $styleItemArr)
    {
        foreach ($styleItemArr as $styleItem) {
            //分析单元格样式
            $analysisedStyleArr = $this->analysisRowStyle($styleItem['data']);

            //分析作用范围
            $rangeArr = $this->analysisStyleRange(0, $styleItem['range'] ?? '');

            //如果没有样式，跳到下一个
            if (empty($analysisedStyleArr)) {
                continue;
            }
            //设置单元格样式
            foreach ($rangeArr as $range) {
                $this->styleConfig->setRowStyle($range, $analysisedStyleArr);
            }
        }
    }


    private function decodeHeaderStyleData(array $styleItem): array
    {
        //如果不是range和data直接分开来的，那么要进行字符串转换处理
        $finalStyleItemTemp = [];
        foreach ($styleItem as $fieldName => $value) {
            $fieldNameArr = explode(':', $fieldName);
            $finalStyleItemTemp[] = [
                'range' => $fieldNameArr[1] ?? '',
                'data' => [
                    $fieldNameArr[0] => $value
                ]
            ];
        }
        return $finalStyleItemTemp;
    }
}