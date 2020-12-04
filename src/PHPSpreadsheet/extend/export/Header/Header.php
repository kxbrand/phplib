<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Header;

use kxbrand\phplib\PHPSpreadsheet\core\Export;

class Header
{
    private $headerData = [];
    private $headerStyle;
    private $headerStyleArr = [];
    private $export;

    public function __construct()
    {
        $this->headerStyle = new HeaderStyle();
        $this->export = Export::getInstance();
    }


    /**
     * @title 获取头部数据
     * @return array
     * @author millionmile
     * @time 2020/07/06 11:25
     */
    public function getHeader(): array
    {
        return $this->headerData;
    }

    /**
     * @title 设置头部信息
     * @param array $headerData
     * @return array
     * @author millionmile
     * @time 2020/07/06 10:50
     */
    public function setHeader(array $headerData): array
    {
        $colNum = 1;  //当前列数
        //支持3种类型的标题
        foreach ($headerData as $headerName => $headerItem) {
            switch (true) {
                case is_int($headerName) && is_string($headerItem):
                    $this->headerData[$headerItem] = null;
//                    $headerData = [
//                        'ID',
//                        '名称'
//                    ];
//                    //处理成
//                    $headerData = [
//                        'ID' => null,
//                        '名称' => null
//                    ];
                    break;
                case is_string($headerName) && is_string($headerItem):
                    $this->headerData[$headerName] = [
                        'field' => $headerItem
                    ];
//                    $headerData = [
//                        'ID' => 'title',
//                        '名称' => 'name'
//                    ];
                    break;
                case !is_int($headerName) && is_array($headerItem):
                    $this->headerData[$headerName] = [
                        'field' => $headerItem['field'],
                        'formatter' => $headerItem['formatter'] ?? null,
                    ];
                    if (isset($headerItem['style'])) {
                        //todo 处理style样式
                        $this->headerStyleArr[$colNum] = $headerItem['style'];
                    }
//                    $headerData = [
//                        'ID' => [
//                            'field' => 'title',
//                            'style' => [
//                                    'range'       //应用范围，默认是all，其他的写行号
//                                    'data'
//                            ],
//                        ],
//                        '名称' => [
//                            'field' => 'name',
//                            'style' => [
//                            ]
//                        ]
//                    ];
//                    //处理成
//                    $headerData = [
//                        'ID' => null,
//                        '名称' => null
//                    ];
                    break;
            }
            $colNum++;
        }
        return $this->headerData ?? [];
    }


    /**
     * @title 最终设置每列的样式
     * @author millionmile
     * @time 2020/07/06 16:47
     */
    public function setFinalStyle()
    {
        $this->headerStyle->setStyle($this->headerStyleArr);
    }
}