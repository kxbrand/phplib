<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\FlushExport;


use kxbrand\phplib\PHPSpreadsheet\core\Export;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

class Csv
{
    private $export;
    private $headerData;

    public function __construct()
    {
        $this->export = Export::getInstance();
        $this->headerData = $this->export->getHeaderData();
    }

    public function export(string $fileName, callable $writeData)
    {
        $writer = $this->export->getWriter('csv');
        header('Content-Description: File Transfer');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename=' . $fileName . '.csv');
        header('Cache-Control: max-age=0');
        try {
            $writer->save('php://output');
        } catch (Exception $e) {
            echo $e;
            exit();
        }
        $fp = fopen('php://output', 'a');//打开output流

        $appendData = function ($data) use ($fp) {
            $finalRowData = $this->getFinalRowData($data);
            foreach ($finalRowData as $item) {
                mb_convert_variables('GBK', 'UTF-8', $item);
                fputcsv($fp, $item);
            }
            //刷新输出缓冲到浏览器
            ob_flush();
            flush();//必须同时使用 ob_flush() 和flush() 函数来刷新输出缓冲。
        };

        $writeData($appendData);

        fclose($fp);
        exit();
    }

    /**
     * @title 根据标题行格式化数据
     * @param $data
     * @return array
     * @author millionmile
     * @time 2020/07/10 18:30
     */
    private function getFinalRowData($data)
    {
        $finalData = [];
        if (empty($this->headerData)) {
            foreach ($data as &$rowData) {
                $finalData[] = $rowData;
            }
            unset($rowData);
        } else {
            foreach ($data as &$rowData) {
                if (is_integer(key($rowData))) { //如果没有键名
                    $finalRowData = array_slice($rowData, 0, count($this->headerData));
                } else {
                    $finalRowData = [];
                    foreach ($this->headerData as $dataKey) {
                        $finalRowData[] = $rowData[$dataKey] ?? '';   //如果没有该列，那么传空
                    }
                }
                $finalData[] = $finalRowData;
            }
            unset($rowData);
        }
        return $finalData;
    }
}
