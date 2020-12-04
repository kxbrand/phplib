<?php


namespace kxbrand\phplib\PHPSpreadsheet;

use kxbrand\phplib\PHPSpreadsheet\core\Export;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Title\BigTitle;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Cache\CaChe;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Title\DoubleTitle;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ExportService
{
    private $export;


    /**
     * 构造函数
     * @param array $headerData 头部数组
     * @param string|object $cache 是否使用缓存
     */
    public function __construct(array $headerData, $cache = null)
    {
        CaChe::useCaChe($cache);
        $this->export = Export::getInstance();
        $this->export->setHeader($headerData);
    }


    /**
     * @title 写入主体数据
     * @param $data
     * @author millionmile
     * @time 2020/06/29 18:32
     */
    public function appendData($data): void
    {
        $headerData = $this->export->getHeaderData();
        if (empty($headerData)) {
            foreach ($data as &$rowData) {
                $this->export->writeRowData($rowData);
            }
        } else {
            foreach ($data as &$rowData) {
                if (is_integer(key($rowData))) { //如果没有键名
                    $finalRowData = array_slice($rowData, 0, count($headerData));
                } else {
                    $finalRowData = [];
                    foreach ($headerData as $headerItem) {
                        $dataKey = $headerItem['field'];
                        if (!isset($headerItem['formatter']) || !is_callable($headerItem['formatter'])) {
                            $finalRowData[] = $rowData[$dataKey] ?? '';   //如果没有该列，那么传空
                        } else {
                            $finalRowData[] = $headerItem['formatter']($rowData[$dataKey] ?? '', $rowData);
                        }
                    }
                }
                $this->export->writeRowData($finalRowData);
            }
            unset($rowData);
        }
    }


    /**
     * @title 最终所要导入的数据，可复写
     * @author millionmile
     * @time 2020/07/06 10:56
     */
    private function exec(): void
    {
        //写入表头title大单元格
        $this->export->writeStyle();
        //写入统计数据
    }


    /**
     * @title 下载导出数据的文件
     * @param string $fileName 自定义文件名
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @author millionmile
     * @time 2020/06/29 17:08
     */
    public function download(string $fileName): void
    {
        $this->exec();

        //在输出Excel前，缓冲区中处理BOM头
        ob_end_clean();
        ob_start();

        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $fileName . '.xlsx"');
        header('Cache-Control: max-age=0');
        $writer = $this->export->getWriter();
        $writer->save('php://output');

        //删除清空：
        $this->export->delSheet();
        exit;
    }


    /**
     * @title 生成文件
     * @param string $fileName
     * @param string $fileDir
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @author millionmile
     * @time 2020/06/29 17:50
     */
    public function generateFile(string $fileDir, string $fileName): string
    {
        $this->exec();

        $writer = $this->export->getWriter();
        $writer->save($fileDir . '/' . $fileName . '.xlsx');

        //删除清空：
        $this->export->delSheet();
        return $fileDir . '/' . $fileName . '.xlsx';
    }

    /**
     * @title 写入首行大标题
     * @param string $titleStr
     * @author millionmile
     * @time 2020/07/07 15:42
     */
    public function setBigTitle(string $titleStr)
    {
        $headerData = $this->export->getHeaderData();
        $bigTtielObj = new BigTitle(count($headerData));
        $bigTtielObj->setBigTitle($titleStr);
    }


    /**
     * @title 获取当前操作的数据表格
     * @return Worksheet
     * @author millionmile
     * @time 2020/07/06 14:29
     */
    public function &getActiveSheet(): Worksheet
    {
        return $this->export->getActiveSheet();
    }


    /**
     * @title 设置活动sheet的名称
     * @param string $sheetName
     * @author millionmile
     * @time 2020/07/06 16:30
     */
    public function setSheetName(string $sheetName): void
    {
        $this->export->setSheetName($sheetName);
    }

    /**
     * @title 使用流方式下载文件
     * @param string $fileName
     * @param callable $cb
     * @author millionmile
     * @time 2020/07/10 17:40
     */
    public function flushDownload(string $fileName, callable $cb): void
    {
        //csv没有样式，不执行共用添加样式方法
//        $this->exec();

        $flushExport = new extend\export\FlushExport\Csv();
        $flushExport->export($fileName, $cb);
    }

    /**
     * @title 在某行位置插入新行
     * @param int $insertPos
     * @param int $count
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author millionmile
     * @time 2020/08/19 15:00
     */
    public function insertRows(int $insertPos, int $count = 1)
    {
        $this->export->insertRows($insertPos, $count);
    }


    /**
     * @title 切换当前活动表，可使用活动表名称或下标，不存在的话即新增
     * @param $index
     * @param array $headerData
     * @return Worksheet|string
     * @author millionmile
     * @time 2020/08/19 16:39
     */
    public function checkoutSheet($index, array $headerData = [])
    {
        $sheet = $this->export->checkoutSheet($index);

        if (!empty($headerData)) {
            $this->export->setHeader($headerData);
        }
        return $sheet;
    }


    /**
     * @title 创建二级表头
     * @param array $fieldArr
     * @param bool $replace
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author millionmile
     * @time 2020/09/08 17:11
     */
    public function setDoubleTitle(array $fieldArr, bool $replace = false)
    {
        $bigTtielObj = new DoubleTitle();
        $bigTtielObj->setTitle($fieldArr, $replace);
    }
}