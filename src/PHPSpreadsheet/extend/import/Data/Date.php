<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Data;

use Exception;

class Date
{
    /**
     * @title 将单元格的值转为日期（Y-m-d）
     * @param $value
     * @return false|string
     * @author millionmile
     * @time 2020/09/20 11:05
     */
    public static function date($value)
    {
        $value = self::formatDateStrCommon($value);
        //如果是时间戳，那么直接转为时间返回
        if (self::isTimestamp($value)) {
            return date('Y-m-d', $value);
        }
        //如果不是时间戳，也不是日期，那么直接返回
        if (!self::isDate($value)) {
            return false;
        }
        return date('Y-m-d', strtotime($value));
    }

    /**
     * @title 将单元格的值转为日期时间（Y-m-d H:i:s）
     * @param $value
     * @return false|string
     * @author millionmile
     * @time 2020/09/20 11:05
     */
    public static function datetime($value)
    {
        $value = self::formatDateStrCommon($value);
        //如果是时间戳，那么直接转为时间返回
        if (self::isTimestamp($value)) {
            return date('Y-m-d H:i:s', $value);
        }
        //如果不是时间戳，也不是日期，那么直接返回
        if (!self::isDate($value)) {
            return false;
        }
        return date('Y-m-d H:i:s', strtotime($value));
    }


    /**
     * @title 转为时间戳
     * @param $value
     * @return false|int
     * @author millionmile
     * @time 2020/09/20 10:48
     */
    public static function timestamp($value)
    {
        //如果是时间戳，那么直接返回
        $value = self::formatDateStrCommon($value);
        if (self::isTimestamp($value)) {
            return intval($value);
        }
        //如果不是时间戳，也不是日期，那么直接返回
        if (!self::isDate($value)) {
            return false;
        }
        return strtotime($value);
    }

    /**
     * @title 将错误的日期格式化
     * @param $dateStr
     * @return string
     * @author millionmile
     * @time 2020/09/20 11:00
     */
    private static function formatDateStrCommon($dateStr): string
    {
        $dateStr = Text::trim($dateStr);
        $dateStr = Text::noWrap($dateStr);
        return preg_replace('/[._]/', '-', $dateStr);
    }

    /**
     * @title 判断是否是日期
     * @param $dateString
     * @return bool
     * @author millionmile
     * @time 2020/09/20 10:44
     */
    static function isDate($dateString): bool
    {
        if (!is_string($dateString)) {
            return false;
        }
        return strtotime($dateString) ? true : false;
    }


    /**
     * @title 判断是否是时间戳
     * @param $timestamp
     * @return bool
     * @author millionmile
     * @time 2020/09/20 10:28
     */
    static function isTimestamp(&$timestamp): bool
    {
        if (!is_numeric($timestamp)) {
            return false;
        }
        if (strlen($timestamp) < 10) {
            try {
                $timestamp = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToTimestamp($timestamp);
            } catch (Exception $e) {
                return false;
            }
        }
        return strtotime(date('Y-m-d H:i:s', $timestamp)) == $timestamp;
    }

}