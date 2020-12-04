<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\import\File;

use Exception;
use kxbrand\phplib\PHPSpreadsheet\extend\common\Code;

class File
{
    /**
     * @title 获取文件信息
     * @param $file
     * @return array
     * @author millionmile
     * @time 2020/09/19 16:39
     */
    public static function getFileData($file): array
    {
        try {
            switch (gettype($file)) {
                case 'string':
                    //判断文件是否存在
                    if (!file_exists($file)) {
                        throw new Exception('文件路径错误，文件不存在');
                    }
                    //打开文件，r表示以只读方式打开
                    $handle = fopen($file, "r");
                    //获取文件的统计信息
                    $fstat = fstat($handle);
                    $filePath = $file;
                    $fileName = basename($file);
                    $fileSize = $fstat["size"];
                    break;
                case 'array':
                    //判断是否是post提交的文件
                    $fileName = $file["name"];
                    $filePath = $file["tmp_name"];
                    $fileSize = $file['size'];
                    break;
                default:
                    throw new Exception('请传入文件地址或临时文件参数');
                    break;
            }

            //获取文件后缀名
            $ext = self::getFileExt($filePath);

            return [
                'code' => Code::SUCCESS,
                'msg' => '获取文件相关信息成功',
                'data' => [
                    'name' => $fileName,
                    'path' => $filePath,
                    'size' => $fileSize,
                    'ext' => $ext,
                ]
            ];
        } catch (Exception $e) {
            return [
                'code' => Code::ERROR,
                'msg' => $e->getMessage(),
                'data' => [],
            ];
        }
    }

    /**
     * @title 获取文件后缀类型
     * @param string $file
     * @return string
     * @author millionmile
     * @time 2020/09/18 10:31
     */
    protected static function getFileExt(string $file): string
    {
        $ext = explode('.', $file);
        if (empty($ext)) {
            return '';
        }
        return strtolower(end($ext));
    }


    /**
     * @title 验证文件是否满足条件
     * @param array $fileData 文件信息
     * @param array $validateConfig
     * @return array
     * @author millionmile
     * @time 2020/09/18 10:22
     */
    public static function validateFile(array $fileData, array $validateConfig = []): array
    {
        if (isset($validateConfig['allow_ext']) && !in_array($fileData['ext'], $validateConfig['allow_ext'])) {
            return [
                'code' => Code::ERROR,
                'msg' => '请上传 ' . (implode('/', $validateConfig['allow_ext'])) . ' 类型文件！',
                'data' => []
            ];
        }

        if (isset($validateConfig['size']) && $fileData['size'] > $validateConfig['size']) {
            return [
                'code' => Code::ERROR,
                'msg' => '文件不能超过' . (round($validateConfig['size'] / 1024 / 1024)) . 'M！',
                'data' => []
            ];
        }
        return [
            'code' => Code::SUCCESS,
            'msg' => '文件无误',
            'data' => []
        ];
    }
}