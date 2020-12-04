<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Data;

use kxbrand\phplib\PHPSpreadsheet\core\Import;
use kxbrand\phplib\PHPSpreadsheet\extend\import\Filter\Filter;

class Data
{
    private $import;
    private $filterObj;

    /**
     * 构造函数
     * @param array $headerData 头部数组
     * @param string|array $fileObj
     * @param string|object $cache 是否使用缓存
     * @throws \Exception
     */
    public function __construct()
    {
        $this->import = Import::getInstance();
        $this->filterObj = new Filter();
    }

    /**
     * @title 将原有的excel数组二维数据转成需要的格式
     * @param array $data
     * @return array
     * @author millionmile
     * @time 2020/09/19 23:55
     */
    public function formatData(array $data)
    {
        $headerObj = $this->import->getCurrentHeader();
        $headerData = $headerObj->getHeader();

        if (empty($headerData)) {
            return $data;
        }

        $finalData = [];
        while ($rowData = array_shift($data)) {
            $finalRowData = [];
            foreach ($headerData as $dataKey => $headerItem) {
                //格式化
                $finalRowData[$dataKey] = Formatter::formatterData(
                    $headerItem['formatter'] ?? null,
                    $rowData[$headerItem['col']] ?? null,
                    $rowData);
            }
            $finalData[] = $finalRowData;
        }
        return $finalData;
    }

}