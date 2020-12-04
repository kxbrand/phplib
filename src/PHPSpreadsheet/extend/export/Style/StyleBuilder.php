<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Style;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class StyleBuilder
{
    private $cellStyleArray = [];                       //最终导出的样式数组
    private $colStyleArray = [];
    private $rowStyleArray = [];

    /**
     * @title 初始化构造器内容，方便重新使用
     * @author millionmile
     * @time 2020/07/06 17:02
     */
    public function init()
    {
        $this->cellStyleArray = [];
        $this->colStyleArray = [];
        $this->rowStyleArray = [];
        return $this;
    }

    /**
     * @title 检查颜色字符串是否是十六进制
     * @param string $color
     * @param bool $return
     * @return array|bool|string
     * @author millionmile
     * @time 2020/07/06 15:46
     */
    private function checkIsColor(string $color, bool $return = false)
    {
        preg_match('/^[0-9a-fA-F]{6}$|^[0-9a-fA-F]{8}$/', $color, $match);
        if (!empty($match)) {
            //如果为true，则返回
            if ($return) {
                return ['argb' => $color];
            }
            return false;
        } else {
            if ($return) {
                return $color;
            }
            return true;
        }
    }


    public function setFontBold(bool $value)
    {
        $this->cellStyleArray['font']['bold'] = $value ? true : false;
        return $this;
    }

    public function setFontItalic(bool $value)
    {
        $this->cellStyleArray['font']['italic'] = $value ? true : false;
        return $this;
    }

    public function setFontUnderline(bool $value)
    {
        $this->cellStyleArray['font']['underline'] = $value ? true : false;
        return $this;
    }

    public function setFont(string $font)
    {
        $this->cellStyleArray['font']['name'] = $font;
        return $this;
    }

    public function setFontSize(int $size)
    {
        $this->cellStyleArray['font']['size'] = $size;
        return $this;
    }

    public function setFontColor(string $color)
    {
        $this->cellStyleArray['font']['color'] = $this->checkIsColor($color, true);
        return $this;
    }

    public function setAlignmentHorizontal(string $horizontal = 'center')
    {
        switch ($horizontal) {
            case 'left':
                $horizontal = Alignment::HORIZONTAL_LEFT;
                break;
            case 'right':
                $horizontal = Alignment::HORIZONTAL_RIGHT;
                break;
            case 'center':
                $horizontal = Alignment::HORIZONTAL_CENTER;
                break;
        }
        $this->cellStyleArray['alignment']['horizontal'] = $horizontal;
        return $this;
    }

    public function setAlignmentVertical(string $vertical = 'center')
    {
        switch ($vertical) {
            case 'top':
                $vertical = Alignment::VERTICAL_TOP;
                break;
            case 'bottom':
                $vertical = Alignment::VERTICAL_BOTTOM;
                break;
            case 'center':
                $vertical = Alignment::VERTICAL_CENTER;
                break;
        }
        $this->cellStyleArray['alignment']['vertical'] = $vertical;
        return $this;
    }

    /**
     * @title 自动换行
     * @author millionmile
     * @time 2020/07/06 16:00
     */
    public function setWrapText(bool $value)
    {
        $this->cellStyleArray['alignment']['wrapText'] = $value ? true : false;
        return $this;
    }

    /**
     * @title 设置文本自适应大小
     * @author millionmile
     * @time 2020/07/06 15:59
     */
    public function setShrinkToFit(bool $value)
    {
        $this->cellStyleArray['alignment']['shrinkToFit'] = $value ? true : false;
        return $this;
    }

    /**
     * @title 设置文本缩进
     * @author millionmile
     * @time 2020/07/06 15:59
     */
    public function setTextIndent(bool $value)
    {
        $this->cellStyleArray['alignment']['indent'] = $value ? true : false;
        return $this;
    }

    /**
     * 设置对角线
     * @title setDiagonal
     * @param string $direction
     * @author millionmile
     * @time 2020/07/06 16:05
     */
    public function setDiagonal(string $direction = 'down', string $borderStyle = '', string $borderColor = '')
    {
        if (empty($borderStyle)) {
            $borderStyle = Border::BORDER_THIN;
        }
        switch (strtolower($direction)) {
            case 'down':
                //从左上到右下
                $direction = Borders::DIAGONAL_DOWN;
                break;
            case 'up':
                //从左下往右上
                $direction = Borders::DIAGONAL_UP;
                break;
            case 'cross':
                //两条线“X”
                $direction = Borders::DIAGONAL_BOTH;
                break;
        }
        $this->cellStyleArray['borders']['diagonalDirection'] = $direction;
        $this->cellStyleArray['borders']['diagonal']['borderStyle'] = $borderStyle;
        if (!empty($borderColor)) {
            $this->cellStyleArray['borders']['diagonal']['color'] = $this->checkIsColor($borderColor, true);
        }
        return $this;
    }


    /**
     * @title 设置边框
     * @param string $borders
     * @param string $borderStyle
     * @param string $borderColor
     * @return $this
     * @author millionmile
     * @time 2020/07/06 18:25
     */
    public function setBorder(string $borders, string $borderStyle = '', string $borderColor = '')
    {
        if (empty($borderStyle)) {
            $borderStyle = Border::BORDER_THIN;
        }
        switch ($borders) {
            case 'allBorders':
            case 'inside':
            case 'outline':
            case 'horizontal':
            case 'top':
            case 'left':
            case 'vertical':
            case 'bottom':
            case 'right':
                $this->cellStyleArray['borders'][$borders]['borderStyle'] = $borderStyle;
                if (!empty($borderColor)) {
                    $this->cellStyleArray['borders'][$borders]['color'] = $this->checkIsColor($borderColor, true);
                }
                break;
        }
        return $this;
    }

    /**
     * @title 设置单元格背景色
     * @param string $color
     * @author millionmile
     * @time 2020/07/06 16:08
     */
    public function setBackgroundColor(string $color)
    {
        $this->cellStyleArray['fill']['fillType'] = Fill::FILL_SOLID;
        if ($this->checkIsColor($color)) {
            echo '仅支持agrb十六进制格式';
            die;
        }
        $this->cellStyleArray['fill']['color'] = $this->checkIsColor($color, true);
        return $this;
    }

    public function setWidth($width = 'auto')
    {
        $this->colStyleArray['width'] = $width;
        return $this;
    }

    public function setHeight(int $height = 0)
    {
        $this->rowStyleArray['height'] = $height;
        return $this;
    }

    //设置样式，最后build出来
    public function build()
    {
        return [
            'cellStyle' => $this->cellStyleArray,
            'colStyle' => $this->colStyleArray,
            'rowStyle' => $this->rowStyleArray,
        ];
    }
}