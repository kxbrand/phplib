<?php


namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Cache;


class CaChe
{
    /**
     * @title 预估将消耗的内存
     * @param int $colCount
     * @param int $rowCount
     * @return string
     * @author millionmile
     * @time 2020/07/07 14:21
     */
    public static function estimateConsumptionMemory(int $colCount, int $rowCount)
    {
        //PhpSpreadsheet在工作表中平均每个单元格使用约1k
        $cellMemory = 1 / 1024; //改为M为单位
        return '预计消耗内存' . round($cellMemory * $colCount * $rowCount, 2) . 'M';
    }

    /**
     * @title 使用缓存
     * @param string|object $cache
     * @return null
     * @author millionmile
     * @time 2020/07/07 14:52
     */
    public static function useCaChe($cache)
    {
        //如果传输字符串，使用内置的缓存对象方式
        switch (true) {
            //如果没有传输，代表不使用缓存
            case empty($cache):
                return null;
                break;
            case is_string($cache) && $cache === 'file':
                \PhpOffice\PhpSpreadsheet\Settings::setCache(new FileCache('/tmp/excel_cache'));
                return null;
                break;
            case is_string($cache) && $cache === 'redis':
                $client = new \Redis();
                $client->connect('127.0.0.1', 6379);
                $pool = new \Cache\Adapter\Redis\RedisCachePool($client);
                break;
            case is_string($cache) && $cache === 'apcu':
                $pool = new \Cache\Adapter\Apcu\ApcuCachePool();
                break;
            case is_string($cache) && $cache === 'memcache':
                $client = new \Memcache();
                $client->connect('localhost', 11211);
                $pool = new \Cache\Adapter\Memcache\MemcacheCachePool($client);
                break;
            case is_object($cache) && get_class($cache) == 'Redis':
                $pool = new \Cache\Adapter\Redis\RedisCachePool($cache);
                break;
            case is_object($cache) && get_class($cache) == 'Memcache':
                $pool = new \Cache\Adapter\Memcache\MemcacheCachePool($cache);
                break;
            default:
                die('设置的缓存方式未生效');
        }
        $simpleCache = new \Cache\Bridge\SimpleCache\SimpleCacheBridge($pool);
        \PhpOffice\PhpSpreadsheet\Settings::setCache($simpleCache);
    }
}