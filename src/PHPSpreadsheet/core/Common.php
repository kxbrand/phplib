<?php


namespace kxbrand\phplib\PHPSpreadsheet\core;

use Exception;

class Common
{
    /**
     * @title 获取所需要的列列名-从1开始
     * @author millionmile
     * @time 2020/07/06 10:07
     */
    public static function getColName($colCount)
    {
        static $colArr = [];
        //如果不存在，则重新获取
        if (!isset($colArr[$colCount - 1])) {
            $colArr = self::getColArr($colCount);
        }
        return $colArr[$colCount - 1] ?? null;
    }


    /**
     * @title 获取列名对应的下标。将A列转为下标1
     * @param string $colName 列名，如"A"
     * @return float|int|mixed|null
     * @throws Exception
     * @author millionmile
     * @time 2020/09/19 18:45
     */
    public static function getColNameToIndex(string $colName)
    {
        static $colNameArr = [];

        $colName = strtoupper($colName);
        if (isset($colNameArr[$colName])) {
            return $colNameArr[$colName];
        }
        //如果不存在，则重新获取
        if (!preg_match('/^[A-Z]+$/', $colName)) {
            throw new Exception('输入的列名错误');
        }
        $index = 0;
        for ($i = 0; $i < strlen($colName); $i++) {
            $index += (ord($colName[$i]) - ord('A') + 1) * (int)pow(26, strlen($colName) - $i - 1);
        }
        $colNameArr[$colName] = $index;
        return $colNameArr[$colName];
    }


    //设置列一维数组(最多701列)
    public static function getColArr($count = 26)
    {
        $indData = 'A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z';
        $indData = explode(',', $indData);
        $curCount = 26;
        for ($i = 0; $i < 26; $i++) {
            for ($j = 0; $j < 26; $j++) {
                if ($curCount >= $count) {
                    return $indData;
                }
                $indData[] = $indData[$i] . $indData[$j];
                $curCount++;
            }
        }
        return $indData;
    }
}