<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\Header;

use kxbrand\phplib\PHPSpreadsheet\core\Common;

class Header
{
    private $headerData = [];

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
                    $this->headerData[$headerItem] = [
                        'col' => Common::getColName($colNum),
                        'filter' => null,
                        'formatter' => null,
                    ];
//                    $headerData = [
//                        'id',
//                        'name'
//                    ];
//                    //处理成
//                    $headerData = [
//                        'id' => 'A',
//                        'name' => 'B'
//                    ];
                    break;
                case is_string($headerName) && is_string($headerItem):
                    $this->headerData[$headerName] = [
                        'col' => $headerItem,
                        'filter' => null,
                        'formatter' => null,
                    ];
//                    $headerData = [
//                        'id' => 'A',
//                        'name' => 'B'
//                    ];
                    break;
                case !is_int($headerName) && is_array($headerItem):
                    $this->headerData[$headerName] = [
                        'col' => $headerItem['col'] ?? Common::getColName($colNum),
                        'filter' => $headerItem['filter'] ?? null,
                        'formatter' => $headerItem['formatter'] ?? null,
                    ];
//                    $headerData = [
//                        'id' => [
//                            'col' => 'B',
//                            'filter' => [
//                                  'row_range'=>'1:5,30'           //行范围，同原来的style设置范围一样
//                                  'color'=>'#FF0000|#00FF00',       //应用范围，默认是all，其他的写行号
//                                  'font'=>'bold',             //一维中的是and的情况
//                                  'alignment'=>'',
//                                  [                           //二维数组中的是or的情况
//                                      'color'=>['#FF0000'],
//                                      'font'=>'bold',
//                                  ]
//                            ],
//                        ],
//                        'name' => [
//                            'col' => 'A',
//                            'filter' => [
//                            ]
//                        ]
//                    ];
                    break;
            }
            $colNum++;
        }
        return $this->headerData ?? [];
    }
}