<?php

namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Filter;

use PhpOffice\PhpSpreadsheet\Cell\Cell;

class Filter
{
    const FilterSupportArr = [
        'row_range',
        'font',
        'font_name',
        'font_size',
        'color',
        'background',
        'border',
        'alignment',
        'horizontal',
        'vertical',
    ];

    /**
     * @title 是否要保留
     * @param string $col 列名"A"
     * @param int $currentRow 行数
     * @param Cell $cell 当前单元格对象
     * @param array $filterData 当前活动表header数据
     * @return bool 是否要过滤掉  true为过滤掉
     * @author millionmile
     * @time 2020/09/20 23:32
     */
    public static function isRetained(string $col, int $currentRow, Cell $cell, array $filterData = []): bool
    {
        //如果不存在过滤配置，那么无需过滤掉
        if (empty($filterData) || !isset($filterData[$col]) || empty($filterData[$col])) {
            return true;
        }
        $filterFlag = false;    //如果存在了true，那么直接返回要保留的
        foreach ($filterData[$col] as $filterItem) {
            //一维数组中，只要里面元素全部满足，那么即为true，保留
            foreach ($filterItem as $field => $filterVArr) {
                if ($field === 'row_range') {
                    $filterItemFlag = Row::checkRowRange($currentRow, $filterVArr);
                } else {
                    $filterItemFlag = Style::checkCellStyle($field, $filterVArr, $cell);
                }
                //将switch部分转移到Style类中。一假即假
                if (!$filterItemFlag) {
                    break;
                }
            }
            //如果已经确定结果为true，那么不需要看下一项
            if ($filterItemFlag) {
                $filterFlag = true;
                break;
            }
        }
        return $filterFlag;
    }


    /**
     * @title 将表头的filter字段格式化
     * @param array $headerData
     * @return array
     * @author millionmile
     * @time 2020/09/20 16:02
     */
    public static function formatHeaderFilterData(array $headerData): array
    {
        if (empty($headerData)) {
            return [];
        }

        //最终格式化后的filter数据
        $filterData = [];
        foreach ($headerData as $headerItem) {
            if (!isset($headerItem['filter']) || empty($headerItem['filter'])) {
                continue;
            }

            $filter = self::formatFilterData($headerItem['filter']);
            $filterData[$headerItem['col']] = $filter;
        }

        return array_filter($filterData);
    }

    /**
     * @title 格式化现有的filter数据
     * @param array $filterData
     * @return array
     * @author millionmile
     * @time 2020/09/20 16:25
     */
    private static function formatFilterData(array $filterData): array
    {
//        [     //一维中的是or的情况
//            [
//                  'row_range' => '1:5,30',           //行范围，同原来的style设置范围一样
//                  'color' => '#FF0000|#00FF00',       //应用范围，默认是all，其他的写行号
//                  'font' => 'bold|',
//                  'alignment' => 'center|',
//            ],
//            [                           //二维数组中的是and的情况
//                  'color' => '#FF0000|#00FF00',
//                  'font' => 'bold',
//            ]
//        ]
        if (empty($filterData)) {
            return [];
        }

        $formatFilterField = function (string $filterName, string $filterItem) {
            $filterItem = str_replace(' and ', '&', $filterItem);
            $filterItem = str_replace(' or ', '|', $filterItem);
            switch ($filterName) {
                case 'row_range':
                case 'font_size':
                    $filterItem = str_replace(',', '|', $filterItem);
                    $filterItem = str_replace('-', ':', $filterItem);
                    $filterItemArr = explode('|', $filterItem);
                    foreach ($filterItemArr as &$filterV) {
                        if (strpos($filterV, ':') !== false) {
                            $filterV = self::trim_array(explode(':', $filterV));
                        }
                    }
                    unset($filterV);
                    return $filterItemArr;
//                case 'color':
//                case 'background':
//                case 'border':
//                case 'font':
//                case 'alignment':
//                    return explode('|', $filterItem);
                default:
                    //外层是or
                    $filterItemArr = self::trim_array(explode('|', $filterItem));
                    //内层是and
                    foreach ($filterItemArr as &$filterVItem) {
                        $filterVItem = self::trim_array(explode('&', $filterVItem));
                    }
                    unset($filterVItem);
                    return $filterItemArr;
            }
        };

        $newOrArr = [];  //用于保存起了键名的一维数组内容，统一放为一个add的数组中
        $finalFilterData = [];
        foreach ($filterData as $filterName => &$filterItem) {
            if (empty($filterItem)) {
                continue;
            }

            //如果设置了键名，是or的情况
            if (is_string($filterName)) {
                //如果是支持的，那么转换格式保存到newOrArr中
                if (in_array($filterName, self::FilterSupportArr)) {
                    $newOrArr[$filterName] = array_merge($newOrArr[$filterName] ?? [],
                        $formatFilterField($filterName, $filterItem));
                }
                continue;
            }

            if (is_array($filterItem)) {
                foreach ($filterItem as $filterVName => &$filterV) {
                    $filterV = $formatFilterField($filterVName, $filterV);
                }
                unset($filterV);
                $finalFilterData[] = array_filter($filterItem);
            }

        }
        unset($filterItem);

        $finalFilterData[] = array_filter($newOrArr);

        return array_filter($finalFilterData);
    }

    /**
     * @title 对数组执行trim操作
     * @param $input
     * @return array|string
     * @author millionmile
     * @time 2020/09/22 21:56
     */
    public static function trim_array($input)
    {
        if (!is_array($input)) {
            return trim($input);
        }
        return array_map('self::trim_array', $input);
    }
}