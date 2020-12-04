<?php


namespace kxbrand\phplib;

use kxbrand\phplib\PHPSpreadsheet\core\Import;
use kxbrand\phplib\PHPSpreadsheet\extend\export\Cache\CaChe;
use kxbrand\phplib\PHPSpreadsheet\extend\import\page\Page;
use kxbrand\phplib\PHPSpreadsheet\extend\import\Data\Data;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ImportService
{
    private $import;
    private $pageObj;
    private $dataObj;

    /**
     * 构造函数
     * @param array $headerData 头部数组
     * @param string|array $fileObj
     * @param string|object $cache 是否使用缓存
     * @throws \Exception
     */
    public function __construct(array $headerData, $fileObj, $cache = null)
    {
        CaChe::useCaChe($cache);
        $this->import = Import::getInstance($fileObj);
        $this->import->setHeader($headerData);
        $this->pageObj = new Page();
        $this->dataObj = new Data();
    }

    /**
     * @title 分页获取excel数据
     * @param int $page
     * @param int $pageSize
     * @return array|bool
     * @throws \Exception
     * @author millionmile
     * @time 2020/09/19 16:13
     */
    public function getPageData(int $page = 1, int $pageSize = 100)
    {
        $this->pageObj->setPage($page);
        $this->pageObj->setPageSize($pageSize);

        //分页：循环获取，如果有数据的话，数据记录数就+1，否则不加；
        //core核心函数中使用row当前获取数据行做记录，每次获取都+1；（注意每个sheet不一样）

        //获取当前页的页数
        $pageRow = $this->pageObj->getPageRow();

        //获取指定范围内的数据
        $data = $this->import->getPageDataCommon($pageRow['startRow'], $pageRow['endRow']);
        if (!empty($data)) {
            $data = $this->dataObj->formatData($data);
        }
        return $data;
    }


    /**
     * @title 跳过前几行数据不获取-分页常用
     * @param int $skip
     * @param bool $allSheet 是否作用到所有sheet表上
     * @author millionmile
     * @time 2020/09/27 10:13
     */
    public function skipFirstRows(int $skip, bool $allSheet = false): void
    {
        $this->pageObj->skipFirstRows($skip, $allSheet);
    }

    /**
     * @title 分页获取excel数据并处理
     * @param callable $cb
     * @param int $pageSize
     * @author millionmile
     * @time 2020/06/29 18:32
     */
    public function getPageDataCb(callable $cb, int $pageSize = 100): void
    {
        //分页：循环获取，如果有数据的话，数据记录数就+1，否则不加；
        //core核心函数中使用row当前获取数据行做记录，每次获取都+1；（注意每个sheet不一样）
        $this->pageObj->setPageSize($pageSize);

        //获取当前头部信息，进行数据获取
//        $headerData = $this->import->getHeaderData();
        while (true) {
            //获取当前页的页数
            $pageRow = $this->pageObj->getPageRow();
            $data = $this->import->getPageDataCommon($pageRow['startRow'], $pageRow['endRow']);
            if ($data === false) {
                break;
            }
            if (!empty($data)) {
                $data = $this->dataObj->formatData($data);
            }
            $this->pageObj->nextPage();
            $cb($data);
        }
    }


    /**
     * @title 获取当前活动表的所有数据
     * @return array
     * @author millionmile
     * @time 2020/06/29 18:32
     */
    public function getTheSheetAllData(): array
    {
        //获取当前头部信息，进行数据获取
        $pageRow = $this->pageObj->getPageRow();
        $data = $this->import->getPageDataCommon($pageRow['startRow']);
        if (!empty($data)) {
            $data = $this->dataObj->formatData($data);
        }
        return $data ?: [];
    }


    /**
     * @title 清空sheet对象
     * @author millionmile
     * @time 2020/06/29 17:08
     */
    public function delSheet(): void
    {
        //删除清空
        if (!empty($this->import)) {
            $this->import->delSheet();
        }
    }

    /**
     * @title 析构函数 销毁sheet
     */
    public function __destruct()
    {
        //删除清空
        if (!empty($this->import)) {
            $this->import->delSheet();
        }
    }

    /**
     * @title 获取当前操作的数据表格
     * @return Worksheet
     * @author millionmile
     * @time 2020/07/06 14:29
     */
    public function &getActiveSheet(): Worksheet
    {
        return $this->import->getActiveSheet();
    }


    /**
     * @title 获取所有工作表的名称集合
     * @return string[]
     * @author millionmile
     * @time 2020/09/27 11:21
     */
    public function getSheetNames(): array
    {
        return $this->import->getSheetNames();
    }


    /**
     * @title 获取所有sheet数据
     * @return array
     * @author millionmile
     * @time 2020/07/06 16:30
     */
    public function getAllSheetData(): array
    {
        $sheetNames = $this->import->getSheetNames();
        $preActiveSheetIndex = $this->import->getActiveSheetIndex();    //记录当前激活的sheet表
        $allSheetData = [];
        foreach ($sheetNames as $sheetName) {
            //切换到某个位置中，然后获取数据
            $this->import->checkoutSheet($sheetName);
            $allSheetData[$sheetName] = $this->getTheSheetAllData();
        }
        //切换回原来激活的sheet表
        $this->import->checkoutSheet($preActiveSheetIndex);
        return $allSheetData;
    }

    /**
     * @title 切换当前活动表，可使用活动表名称或下标，不存在的话即新增
     * @param mixed $index
     * @param array $headerData
     * @return Worksheet|string
     * @author millionmile
     * @time 2020/08/19 16:39
     */
    public function checkoutSheet($index, array $headerData = [])
    {
        $sheet = $this->import->checkoutSheet($index);
        if (!empty($headerData)) {
            $this->import->setHeader($headerData);
        }
        return $sheet;
    }

}