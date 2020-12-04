<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Data;

class Formatter
{

    /**
     * @title formatterData
     * @param string|object|null $formatter
     * @param null $data
     * @param array $rowData
     * @return null
     * @author millionmile
     * @time 2020/09/20 00:51
     */
    public static function formatterData($formatter = null, $data = null, array $rowData = [])
    {
        //如果不存在格式化函数，那么直接返回
        if (!isset($formatter) || empty($formatter)) {
            return $data;
        }

        switch (gettype($formatter)) {
            case 'string':
                //使用内置格式化函数
                return self::formatterInner($data, $formatter);
            case 'object':
                //使用用户自定义格式化函数
                return $formatter($data, $rowData);
            default:
                return $data;
        }
    }


    /**
     * @title 格式化函数使用内置处理
     * @param mixed $value 单元格数据
     * @param string $formatterStr 内置格式化函数集合
     * @return false|int|string|array
     * @author millionmile
     * @time 2020/09/20 10:22
     */
    private static function formatterInner($value, string $formatterStr = '')
    {
        if(empty($value)){
            return $value;
        }
        $formatterArr = explode('|', $formatterStr);
        while ($formatter = array_shift($formatterArr)) {
            switch ($formatter) {
                case 'date':
                    $value = Date::date($value);
                    break;
                case 'datetime':
                    $value = Date::datetime($value);
                    break;
                case 'timestamp':
                    $value = Date::timestamp($value);
                    break;
                case 'no_warp':
                    $value = Text::noWrap($value);
                    break;
                case 'trim':
                    $value = Text::trim($value);
                    break;
                case 'json':
                    $value = Text::json($value);
                    break;
            }
        }
        return $value;
    }

}