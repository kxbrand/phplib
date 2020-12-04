<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Color;

use PhpOffice\PhpSpreadsheet\Style\Color as PhpSpreadsheetColor;

class Color
{
    const color_map = [
        'black' => PhpSpreadsheetColor::COLOR_BLACK,
        'white' => PhpSpreadsheetColor::COLOR_WHITE,
        'red' => PhpSpreadsheetColor::COLOR_RED,
        'dark_red' => PhpSpreadsheetColor::COLOR_DARKRED,
        'blue' => PhpSpreadsheetColor::COLOR_BLUE,
        'dark_blue' => PhpSpreadsheetColor::COLOR_DARKBLUE,
        'green' => PhpSpreadsheetColor::COLOR_GREEN,
        'darkgreen' => PhpSpreadsheetColor::COLOR_DARKGREEN,
        'yellow' => PhpSpreadsheetColor::COLOR_YELLOW,
        'dark_yellow' => PhpSpreadsheetColor::COLOR_DARKYELLOW,
    ];

    /**
     * @title 将颜色格式化成RGB/RGBA形式
     * @author millionmile
     * @time 2020/09/22 23:35
     */
    public static function formatColorRGB(string $color): array
    {
        $color = str_replace(' ', '', $color);
        //判断是内置的颜色字符串
        if (isset(self::color_map[$color])) {
            return [
                'type' => 'argb',
                'color' => self::color_map[$color]
            ];
        }
        preg_match('/^#?([0-9a-fA-F]{6})$|^#?([0-9a-fA-F]{8})$/', $color, $match);
        if (!empty($match)) {
            //如果是rgb形式
            if (!isset($match[2]) || empty($match[2])) {
                return [
                    'type' => 'rgb',
                    'color' => strtoupper($match[1])
                ];
            } else {
                return [
                    'type' => 'argb',
                    'color' => strtoupper($match[2])
                ];
            }
        }

        return self::formatColorNumToRGB($color);
    }


    /**
     * @title 将rgb(255,255,0)/rgba(255,255,0,0)格式转成十六进制
     * @param string $color
     * @return array
     * @author millionmile
     * @time 2020/09/23 01:05
     */
    public static function formatColorNumToRGB(string $color): array
    {
        $resColorArr = [
            'type' => false,
            'color' => false,
        ];

        //如果是RGB
        preg_match('/^rgb\(([0-9]{1,3}),([0-9]{1,3}),([0-9]{1,3})\)$/i', $color, $match);
        if (!empty($match)) {
            if (!isset($match[1]) || !isset($match[2]) || !isset($match[3]) || $match[1] > 255 || $match[2] > 255 || $match[3] > 255) {
                return $resColorArr;
            }
            return [
                'type' => 'rgb',
                'color' => strtoupper(
                    sprintf("%02X", $match[1]) .
                    sprintf("%02X", $match[2]) .
                    sprintf("%02X", $match[3])
                )
            ];
        }

        preg_match('/^argb\(([0-9]{1,3}),([0-9]{1,3}),([0-9]{1,3}),([0-9]{1,3})\)$/i', $color, $match);
        if (!empty($match)) {
            if (!isset($match[1]) || !isset($match[2]) || !isset($match[3]) || !isset($match[4]) || $match[1] > 255 || $match[2] > 255 || $match[3] > 255 || $match[4] > 255) {
                return $resColorArr;
            }
            return [
                'type' => 'argb',
                'color' => strtoupper(
                    sprintf("%02X", $match[1]) .
                    sprintf("%02X", $match[2]) .
                    sprintf("%02X", $match[3]) .
                    sprintf("%02X", $match[4])
                )
            ];
        }


        preg_match('/^rgba\(([0-9]{1,3}),([0-9]{1,3}),([0-9]{1,3}),([0-9]{1,3})\)$/i', $color, $match);
        if (!empty($match)) {
            if (!isset($match[1]) || !isset($match[2]) || !isset($match[3]) || !isset($match[4]) || $match[1] > 255 || $match[2] > 255 || $match[3] > 255 || $match[4] > 255) {
                return $resColorArr;
            }
            return [
                'type' => 'argb',
                'color' => strtoupper(
                    sprintf("%02X", $match[4]) .
                    sprintf("%02X", $match[1]) .
                    sprintf("%02X", $match[2]) .
                    sprintf("%02X", $match[3])
                )
            ];
        }
        return $resColorArr;
    }
}