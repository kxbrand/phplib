<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\page;

use kxbrand\phplib\PHPSpreadsheet\core\Import;

class Page
{
    private $pageArr = [];          //默认页码为1
    private $pageSizeArr = [];    //默认每页100行
    private $skipArr = [];
    private $import;
    private $allSkip = 0; //所有sheet都跳过这些行数-如果原表已经有跳过行数，以原表的为准
    const DefaultPageSize = 100;
    const DefaultPage = 1;

    /**
     * Page constructor.
     * @throws \Exception
     */
    public function __construct()
    {
        $this->import = Import::getInstance();
    }


    /**
     * @title 设置当前页码
     * @param int $page
     * @author millionmile
     * @time 2020/09/19 15:40
     */
    public function setPage(int $page): void
    {
        $sheetIndex = $this->import->getActiveSheetIndex();
        $this->pageArr[$sheetIndex] = $page;
    }

    /**
     * @title 设置每页条数
     * @param int $pageSize
     * @author millionmile
     * @time 2020/09/19 15:41
     */
    public function setPageSize(int $pageSize): void
    {
        $sheetIndex = $this->import->getActiveSheetIndex();
        $this->pageSizeArr[$sheetIndex] = $pageSize;
    }


    /**
     * @title skipRows
     * @param int $skip 跳过前几行数据不获取-分页常用
     * @param bool $allSheet 是否作用到所有sheet表上
     * @author millionmile
     * @time 2020/09/27 10:14
     */
    public function skipFirstRows(int $skip, bool $allSheet = false)
    {
        $sheetIndex = $this->import->getActiveSheetIndex();
        $this->skipArr[$sheetIndex] = $skip;
        if ($allSheet) {
            $this->allSkip = $skip;
        }
    }


    /**
     * @title 获取分页时的起始结束行数
     * @return array
     * @author millionmile
     * @time 2020/09/27 10:22
     */
    public function getPageRow(): array
    {
        $sheetIndex = $this->import->getActiveSheetIndex();
        $pageSize = $this->pageSizeArr[$sheetIndex] ?? self::DefaultPageSize;
        $page = $this->pageArr[$sheetIndex] ?? self::DefaultPage;
        $skip = $this->skipArr[$sheetIndex] ?? $this->allSkip;
        return [
            'startRow' => $pageSize * ($page - 1) + 1 + $skip,
            'endRow' => $pageSize * $page + $skip,
        ];
    }

    /**
     * @title 当前活动表跳到下一页
     * @author millionmile
     * @time 2020/09/27 10:53
     */
    public function nextPage()
    {
        $sheetIndex = $this->import->getActiveSheetIndex();
        $this->pageArr[$sheetIndex]++;
    }
}