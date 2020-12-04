<?php


namespace kxbrand\phplib\PHPSpreadsheet\core;

use kxbrand\phplib\PHPSpreadsheet\extend\export\Header\Header;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Export
{
    private static $instance;
    private $spreadsheet;   //当前创建的excel对象
    private $activeSheet;   //当前使用的sheet表
    private $currentRow = [];    //当前操作的行
    private $colArr;
    private $headerObj = [];


    //构造函数
    private function __construct()
    {
        $this->spreadsheet = new Spreadsheet();
        $this->activeSheet = $this->spreadsheet->getActiveSheet();
        $activeIndex = $this->getActiveSheetIndex();
        $this->currentRow[$activeIndex] = 1;
    }

    public static function getInstance(): Export
    {
        if (!self::$instance) {
            self::$instance = new self();
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
     * @title 设置活动sheet的名称
     * @param string $sheetName
     * @author millionmile
     * @time 2020/07/06 16:30
     */
    public function setSheetName(string $sheetName)
    {
        $this->activeSheet->setTitle($sheetName);
    }

    public function &getCurrentHeader(): Header
    {
        $activeIndex = $this->getActiveSheetIndex();
        if (!isset($this->headerObj[$activeIndex])) {
            $this->headerObj[$activeIndex] = new Header();
        }
        return $this->headerObj[$activeIndex];
    }

    public function &getCurrentRow(): int
    {
        $activeIndex = $this->getActiveSheetIndex();
        if (!isset($this->currentRow[$activeIndex])) {
            $this->currentRow[$activeIndex] = 1;
        }
        return $this->currentRow[$activeIndex];
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
        $headerData = $headerObj->setHeader($headerData);
        $this->colArr = Common::getColArr(count($headerData));

        //写入头部内容
        $this->writeHeaderData();
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
     * @title 一维数组写入数据
     * @param array $data
     * @author millionmile
     * @time 2020/06/29 17:32
     */
    public function writeRowData(array $data): void
    {
        //当前操作列
        $currentCol = 0;
        $currentRow = &$this->getCurrentRow();
        foreach ($data as $value) {
            $this->activeSheet->setCellValue($this->colArr[$currentCol++] . $currentRow, $value);
        }
        $this->nextRow();
    }


    /**
     * @title 写入头部数据
     * @author millionmile
     * @time 2020/06/29 18:29
     */
    private function writeHeaderData()
    {
        $headerObj = &$this->getCurrentHeader();
        $headerData = $headerObj->getHeader();
        if (!empty($headerData)) {
            $this->writeRowData(array_keys($headerData));
        }
    }


    /**
     * @title 获取Xlsx对象
     * @author millionmile
     * @time 2020/07/06 11:19
     */
    public function getWriter($writeType = 'xlsx')
    {
        $writer = null;
        switch ($writeType) {
            case 'xlsx':
                $writer = new Xlsx($this->spreadsheet);
                break;
            case 'csv':
                $writer = new Csv($this->spreadsheet);
                break;
        }
        return $writer;
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
     * @title 操作下一行
     * @author millionmile
     * @time 2020/07/06 12:29
     */
    public function nextRow()
    {
        $currentRow =& $this->getCurrentRow();
        $currentRow++;
    }


    /**
     * @title 插入行
     * @param int $insertPos 要插入新行的位置
     * @param int $count 插入的行数
     * @throws Exception
     * @author millionmile
     * @time 2020/07/07 15:54
     */
    public function insertRows(int $insertPos, int $count = 1)
    {
        $this->activeSheet->insertNewRowBefore($insertPos, $count);
        $currentRow =& $this->getCurrentRow();
        $currentRow += $count;
    }


    /**
     * @title 获取当前的行数
     * @author millionmile
     * @time 2020/07/06 16:45
     */
    public function getTheRow()
    {
        return $this->getCurrentRow() - 1;
    }

    /**
     * @title 写完数据后，填充样式
     * @author millionmile
     * @time 2020/07/07 15:38
     */
    public function writeStyle()
    {
        foreach ($this->headerObj as $index => &$headerObj) {
            $this->checkoutSheet($index);
            $headerObj->setFinalStyle();
        }
        unset($headerObj);
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
                        //如果不存在，那么新建一个
                        $this->spreadsheet->createSheet($index);
                    }
                    //判断是否存在，没有的话就新增-----结束

                    //切换sheet-----开始
                    $this->activeSheet = $this->spreadsheet->setActiveSheetIndex($index);
                    //切换sheet-----结束
                    return $this->getActiveSheet();
                case 'string':
                    //使用文件名称进行处理
                    //如果sheet不存在，那么新建
                    if (!$this->spreadsheet->sheetNameExists($index)) {
                        $newSheet = $this->spreadsheet->createSheet();
                        $newSheet->setTitle($index);    //更改sheet名称
                    }
                    //切换sheet-----开始
                    $this->activeSheet = $this->spreadsheet->setActiveSheetIndexByName($index);
                    //切换sheet-----结束
                    return $this->getActiveSheet();
                default :
                    return 'nnn';
            }
        } catch (Exception $e) {
            return $e->getMessage();
        }
    }

    public function getActiveSheetIndex()
    {
        return $this->spreadsheet->getActiveSheetIndex();
    }
}