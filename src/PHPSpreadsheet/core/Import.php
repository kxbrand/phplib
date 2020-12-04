<?php


namespace kxbrand\phplib\PHPSpreadsheet\core;


use Exception;
use kxbrand\phplib\PHPSpreadsheet\extend\common\Code;
use kxbrand\phplib\PHPSpreadsheet\extend\import\Filter\Filter;
use kxbrand\phplib\PHPSpreadsheet\extend\import\File\File;
use kxbrand\phplib\PHPSpreadsheet\extend\import\Header\Header;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Import
{
    private static $instance;
    private $spreadsheet;   //当前创建的excel对象
    private $activeSheet;   //当前使用的sheet表
    private $headerObj = [];
    private $heightRow = []; //各sheet表的最大行数


    /**
     * Import constructor.
     * @param string|array $fileObj
     * @throws \Exception
     */
    private function __construct($fileObj)
    {
        //默认存在表对象
        $this->spreadsheet = new Spreadsheet();
        $this->activeSheet = new Worksheet();
        //设置当前阅读的文件
        $this->setFile($fileObj);
    }


    /**
     * @title getInstance
     * @param string|array|null $fileObj
     * @return Import
     * @throws \Exception
     * @author millionmile
     * @time 2020/09/19 10:32
     */
    public static function getInstance($fileObj = null): Import
    {
        if (!self::$instance) {
            if ($fileObj === null) {
                throw new Exception('第一次使用必须传入文件');
            }
            self::$instance = new self($fileObj);
        }
        return self::$instance;
    }

    //克隆函数设为私有
    private function __clone()
    {
        // 禁止克隆
    }

    //反序列化函数设为私有
    private function __wakeup()
    {
        // 禁止反序列化  Disable unserialize
    }


    /**
     * @title 设置当前阅读的文件
     * @throws \Exception
     * @author millionmile
     * @time 2020/09/19 10:26
     */
    public function setFile($fileObj)
    {
        //验证上传的文件是否有误
        $fileRes = File::getFileData($fileObj);
        if ($fileRes['code'] === Code::ERROR) {
            throw new Exception($fileRes['msg']);
        }
        $fileValidateRes = File::validateFile($fileRes['data'], ['xls', 'xlsx', 'csv']);
        if ($fileValidateRes['code'] === Code::ERROR) {
            throw new Exception($fileRes['msg']);
        }

        $this->spreadsheet = IOFactory::load($fileRes['data']['path']);

        //获取默认工作表作为当前活动表
        $this->activeSheet = $this->spreadsheet->getActiveSheet();
    }


    /**
     * @title 获取当前工作表的表头
     * @return Header
     * @author millionmile
     * @time 2020/09/19 11:40
     */
    public function &getCurrentHeader(): Header
    {
        $activeIndex = $this->getActiveSheetIndex();
        if (!isset($this->headerObj[$activeIndex])) {
            $this->headerObj[$activeIndex] = new Header();
        }
        return $this->headerObj[$activeIndex];
    }

    /**
     * @title 获取当前工作表的最高行数
     * @return int
     * @author millionmile
     * @time 2020/09/19 16:21
     */
    private function getCurrentHeightRow(): int
    {
        $activeIndex = $this->getActiveSheetIndex();
        if (!isset($this->heightRow[$activeIndex])) {
            $this->heightRow[$activeIndex] = $this->activeSheet->getHighestDataRow();
        }
        return $this->heightRow[$activeIndex];
    }


    /**
     * @title 设置头文件信息
     * @param $headerData
     * @author millionmile
     * @time 2020/07/07 15:38
     */
    public function setHeader($headerData)
    {
        $headerObj = &$this->getCurrentHeader();
        $headerObj->setHeader($headerData);
    }

    /**
     * @title 获取头文件信息
     * @return array
     * @author millionmile
     * @time 2020/07/07 15:38
     */
    public function getHeaderData()
    {
        return $this->getCurrentHeader()->getHeader() ?? [];
    }


    /**
     * @title 获取当前引用sheet对象
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @author millionmile
     * @time 2020/06/29 17:03
     */
    public function &getActiveSheet(): Worksheet
    {
        return $this->activeSheet;
    }


    /**
     * @title 删除清空sheet表单内容
     * @author millionmile
     * @time 2020/07/06 11:20
     */
    public function delSheet(): void
    {
        $this->spreadsheet->disconnectWorksheets();
        $this->spreadsheet = null;
        self::$instance = null;
    }

    /**
     * @title 切换当前活动表，可使用活动表名称或下标，不存在的话即新增
     * @param $index
     * @return Worksheet|string
     * @author millionmile
     * @time 2020/08/19 16:39
     */
    public function checkoutSheet($index)
    {
        try {
            switch (gettype($index)) {
                case 'integer':
                    //如果是数字，代表使用下标进行操作
                    //判断是否存在，没有的话就新增-----开始
                    try {
                        $this->spreadsheet->getSheet($index);
                    } catch (Exception $e) {
                        //如果不存在，那么返回不存在
                        throw new Exception('目标工作表不存在');
                    }
                    //判断是否存在，没有的话就新增-----结束

                    //切换sheet-----开始
                    $this->activeSheet = $this->spreadsheet->setActiveSheetIndex($index);
                    //切换sheet-----结束
                    return $this->getActiveSheet();
                case 'string':
                    //使用文件名称进行处理
                    //如果sheet不存在，那么直接返回不存在
                    if (!$this->spreadsheet->sheetNameExists($index)) {
                        throw new Exception('目标工作表不存在');
                    }
                    //切换sheet-----开始
                    $this->activeSheet = $this->spreadsheet->setActiveSheetIndexByName($index);
                    //切换sheet-----结束
                    return $this->getActiveSheet();
                default :
                    return '切换工作表格式错误';
            }
        } catch (Exception $e) {
            return $e->getMessage();
        }
    }

    /**
     * @title 获取当前活动表下标
     * @return int
     * @author millionmile
     * @time 2020/09/19 11:39
     */
    public function getActiveSheetIndex()
    {
        return $this->spreadsheet->getActiveSheetIndex();
    }

    /**
     * @title 获取所有工作表的名称集合-用于循环获取所有工作表的数据
     * @return string[]
     * @author millionmile
     * @time 2020/09/19 11:48
     */
    public function getSheetNames(): array
    {
        return $this->spreadsheet->getSheetNames();
    }


    /**
     * @title 获取分页数据的公共方法
     * @param int $startRow 开始行数
     * @param int $endRow 结束行数，null代表获取整表数据
     * @return array|bool
     * @author millionmile
     * @time 2020/09/19 16:15
     */
    public function getPageDataCommon(int $startRow = 1, int $endRow = null)
    {
        try {
            if ($startRow <= 0) {
                throw new Exception('开始行数必为正整数');
            }

            $heightRow = $this->getCurrentHeightRow();
            if ($heightRow < $startRow) {
                //如果最高行数已经不存在，那么直接返回false
                throw new Exception('行数超过范围');
            }

            //如果最高行数小于最终获取数据的行数，那么调整最终获取数据行数为heightRow
            if ($endRow !== null && $heightRow < $endRow) {
                $endRow = $heightRow;
            }
            //使用迭代器遍历单元格
            $rowIterator = $this->activeSheet->getRowIterator($startRow, $endRow);
            $headerFilterData = Filter::formatHeaderFilterData($this->getHeaderData());

            $data = [];
            $currentRow = $startRow;    //记录当前行数，用于筛选数据使用
            foreach ($rowIterator as $row) {
                //获取单元格迭代器
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false); //遍历所有单元格-即使没有值
                $rowData = [];
                foreach ($cellIterator as $col => $cell) {  //$col是列名"A"
                    //如果过滤掉
                    if (!Filter::isRetained($col, $currentRow, $cell, $headerFilterData)) {
                        continue;
                    }
                    $rowData[$col] = $cell->getValue();
                }
                //下一行
                $currentRow++;
                //如果所有值都为空，直接跳过
                if (empty(array_filter($rowData))) {
                    continue;
                }
                $data[] = $rowData;
            }
        } catch (Exception $e) {
            return false;
        }
        return $data;
    }

}