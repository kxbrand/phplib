<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Data;

class Text
{
    /**
     * @title 去除左右空格
     * @param $value
     * @return string
     * @author millionmile
     * @time 2020/09/20 10:54
     */
    public static function trim($value): string
    {
        return trim($value);
    }

    /**
     * @title json字符串转为数组，不是json则直接返回
     * @param $value
     * @return mixed
     * @author millionmile
     * @time 2020/09/23 17:08
     */
    public static function json($value)
    {
        $data = json_decode($value, true);
        if (($data && is_object($data)) || (is_array($data) && !empty($data))) {
            return $data;
        }
        return $value;
    }


    /**
     * @title 取消自动换行-将换行符去除
     * @param $value
     * @return string
     * @author millionmile
     * @time 2020/09/20 10:53
     */
    public static function noWrap($value): string
    {
        return preg_replace('/[\n]/is', '', $value);
    }
}