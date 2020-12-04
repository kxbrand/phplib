<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Filter;


use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
use kxbrand\phplib\PHPSpreadsheet\extend\import\Color\Color;

class Style
{

    /**
     * @title 检查单元格是否存在
     * @param string $field
     * @param array $data
     * @param Cell $cell
     * @author millionmile
     * @time 2020/09/21 23:58
     */
    public static function checkCellStyle(string $field, array $data, Cell $cell): bool
    {
        //将switch部分转移到Style类中
        switch ($field) {
            case 'font':
                return self::checkCellFontStyle($data, $cell);
            case 'font_size':
                return self::checkCellFontSize($data, $cell);
            case 'horizontal':
                return self::checkCellHorizontalStyle($data, $cell);
            case 'vertical':
                return self::checkCellVerticalStyle($data, $cell);
            case 'alignment':
                return self::checkCellAlignmentStyle($data, $cell);
            case 'border':
                return self::checkCellBorderStyle($data, $cell);
            case 'color':
                return self::checkCellColorStyle($data, $cell);
            case 'background':
                return self::checkCellBackgroundStyle($data, $cell);
            case 'font_name':
                return self::checkCellFontName($data, $cell);
            default:
                //如果不存在，跳过
                return true;
        }
    }

    /**
     * @title 检查单元格样式的公共处理函数
     * @param array $data
     * @param callable $cb
     * @return bool
     * @author millionmile
     * @time 2020/09/22 22:54
     */
    public static function checkCellStyleCommon(array $data, callable $cb): bool
    {
        $filterFinalItemFlag = false;    //默认为或保留
        foreach ($data as $filterV) {
            $filterItemFlag = true;    //默认为与保留
            foreach ($filterV as $filterItem) {
                $filterItemFlag = $cb($filterItem);
                if (!$filterItemFlag) {
                    break;
                }
            }
            //全部与通过，那么$filterItemFlag将会是true
            if ($filterItemFlag) {
                $filterFinalItemFlag = true;
                break;
            }
        }
        return $filterFinalItemFlag;
    }

    /**
     * @title 验证单元格字体样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:06
     */
    public static function checkCellFontStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            switch ($filterItem) {
                case 'underline':
                    //如果存在下划线，那么为true
                    return $cell->getStyle()->getFont()->getUnderline() !== Font::UNDERLINE_NONE ? true : false;
                case 'bold':
                    //如果没有，那么标为$filterItemFlag=false
                    return $cell->getStyle()->getFont()->getBold();
                case 'italic':
                    //如果没有，那么标为false
                    return $cell->getStyle()->getFont()->getItalic();
                default:
                    //未知的默认为false
                    return false;
            }
        });
    }


    /**
     * @title 验证单元格字体是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:06
     */
    public static function checkCellFontName(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            return $cell->getStyle()->getFont()->getName() === $filterItem;
        });
    }


    /**
     * @title 验证单元格字体大小是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:06
     */
    public static function checkCellFontSize(array $data, Cell $cell): bool
    {
        $filterItemFlag = false;    //默认为或保留
        $cellFontSize = $cell->getStyle()->getFont()->getSize();
        foreach ($data as $range) {
            //如果是数组，代表是范围
            if ((is_array($range) && $range[0] <= $cellFontSize && $cellFontSize <= $range[1])
                || (is_numeric($range) && $cellFontSize == floatval($range))
            ) {
                $filterItemFlag = true;
                break;
            }
        }
        return $filterItemFlag;
    }


    /**
     * @title 验证单元格水平对齐方式样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:06
     */
    public static function checkCellHorizontalStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            switch ($filterItem) {
                case 'center':
                    //如果水平对齐方式是center，那么为true
                    return $cell->getStyle()->getAlignment()->getHorizontal() === Alignment::HORIZONTAL_CENTER;
                case 'left':
                    //如果水平对齐方式是center，那么为true
                    return $cell->getStyle()->getAlignment()->getHorizontal() === Alignment::HORIZONTAL_LEFT;
                case 'right':
                    //如果没有，那么标为false
                    return $cell->getStyle()->getAlignment()->getHorizontal() === Alignment::HORIZONTAL_RIGHT;
                case 'none':
                case 'general':
                    return $cell->getStyle()->getAlignment()->getHorizontal() === Alignment::HORIZONTAL_GENERAL;
                default:
                    //未知的默认为false
                    return false;
            }
        });
    }


    /**
     * @title 验证单元格垂直对齐方式样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:17
     */
    public static function checkCellVerticalStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            switch ($filterItem) {
                case 'center':
                    //如果水平对齐方式是center，那么为true
                    return $cell->getStyle()->getAlignment()->getVertical() === Alignment::HORIZONTAL_CENTER;
                case 'top':
                    //如果水平对齐方式是center，那么为true
                    return $cell->getStyle()->getAlignment()->getVertical() === Alignment::VERTICAL_TOP;
                case 'bottom':
                    //如果没有，那么标为false
                    return $cell->getStyle()->getAlignment()->getVertical() === Alignment::VERTICAL_BOTTOM;
                case 'none':
                case 'justify':
                    return $cell->getStyle()->getAlignment()->getVertical() === Alignment::VERTICAL_JUSTIFY;
                default:
                    //未知的默认为false
                    return false;
            }
        });
    }


    /**
     * @title 验证单元格对齐方式（自动换行，文本自适应、文本缩进）样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:17
     */
    public static function checkCellAlignmentStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            switch ($filterItem) {
                case 'wrap_text':
                    //如果自动换行，那么为true
                    return $cell->getStyle()->getAlignment()->getWrapText();
                case 'shrink_to_fit':
                    //如果文本自动适应（缩小字体填充）
                    return $cell->getStyle()->getAlignment()->getShrinkToFit();
                case 'indent':
                    //如果存在文本缩进
                    return $cell->getStyle()->getAlignment()->getIndent() ? true : false;
                default:
                    //未知的默认为false
                    return false;
            }
        });
    }


    /**
     * @title 验证单元格边框样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:17
     */
    public static function checkCellBorderStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            switch ($filterItem) {
                case 'all':
                    return
                        $cell->getStyle()->getBorders()->getTop()->getBorderStyle() !== Border::BORDER_NONE &&
                        $cell->getStyle()->getBorders()->getBottom()->getBorderStyle() !== Border::BORDER_NONE &&
                        $cell->getStyle()->getBorders()->getLeft()->getBorderStyle() !== Border::BORDER_NONE &&
                        $cell->getStyle()->getBorders()->getRight()->getBorderStyle() !== Border::BORDER_NONE;
                case 'top':
                    return $cell->getStyle()->getBorders()->getTop()->getBorderStyle() !== Border::BORDER_NONE;
                case 'left':
                    return $cell->getStyle()->getBorders()->getLeft()->getBorderStyle() !== Border::BORDER_NONE;
                case 'bottom':
                    return $cell->getStyle()->getBorders()->getBottom()->getBorderStyle() !== Border::BORDER_NONE;
                case 'right':
                    return $cell->getStyle()->getBorders()->getRight()->getBorderStyle() !== Border::BORDER_NONE;
                default:
                    //未知的默认为false
                    return false;
            }
        });
    }


    /**
     * @title 验证单元格字体颜色样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:06
     */
    public static function checkCellColorStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            $colorRes = Color::formatColorRGB($filterItem);
            //如果没有明确的type，那么按照false处理
            switch ($colorRes['type']) {
                case 'rgb':
                    return $colorRes['color'] === $cell->getStyle()->getFont()->getColor()->getRGB();
                case 'argb':
                    return $colorRes['color'] === $cell->getStyle()->getFont()->getColor()->getARGB();
                default:
                    return false;
            }
        });
    }


    /**
     * @title 验证单元格背景颜色样式是否满足
     * @param array $data
     * @param Cell $cell
     * @return bool
     * @author millionmile
     * @time 2020/09/22 00:06
     */
    public static function checkCellBackgroundStyle(array $data, Cell $cell): bool
    {
        return self::checkCellStyleCommon($data, function ($filterItem) use ($cell) {
            $colorRes = Color::formatColorRGB($filterItem);
            //如果没有明确的type，那么按照false处理
            switch ($colorRes['type']) {
                case 'rgb':
                    return $colorRes['color'] === $cell->getStyle()->getFill()->getStartColor()->getRGB();
                case 'argb':
                    return $colorRes['color'] === $cell->getStyle()->getFill()->getStartColor()->getARGB();
                default:
                    return false;
            }
        });
    }
}