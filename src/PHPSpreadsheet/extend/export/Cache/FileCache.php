<?php
namespace kxbrand\phplib\PHPSpreadsheet\extend\export\Cache;

use Psr\SimpleCache\CacheInterface;
class FileCache implements CacheInterface {
    const FILE_SIZE = 3000; //读取时单次缓存行数（文件分割行数）
    private $cache_key = [];
    private $cache = [], $file_handles = [], $cache_dir, $file_prefix;
    private function delCacheDir($path) {
        if (is_dir($path)) {
            foreach (scandir($path) as $val) {
                if ($val != "." && $val != "..") {
                    if (is_dir($path . $val)) {
                        $this->delCacheDir($path . $val . '/');
                        @rmdir($path . $val . '/');
                    } else {
                        unlink($path . $val);
                    }
                }
            }
        }
    }
    private function getFilenameByKey($key) {
        $arr = explode('.', $key);
        $end = array_pop($arr);
        $dir = $this->cache_dir . implode('_', $arr);
        if (! is_dir($dir)) {
            mkdir($dir, 0777, true);
        }
        $line = '';
        $len = strlen($end);
        for ($i = 0; $i < $len; $i++) {
            if (is_numeric($end[$i])) {
                $line .= $end[$i];
            }
        }
        $suf = (int)round($line / self::FILE_SIZE);
        return $dir . '/' . $this->file_prefix . $suf;
    }
    private function getFileHandleByKey($key) {
        $filename = $this->getFilenameByKey($key);
        if (! array_key_exists($filename, $this->file_handles)) {
            $fp = fopen($filename, 'w+');
            if (! $fp) {
                throw new \Exception('生成缓存文件失败');
            }
            $this->file_handles[$filename] = $fp;
        }
        return $this->file_handles[$filename];
    }
    public function __construct($cache_dir) {
        $this->cache_dir = rtrim($cache_dir, '/') . '/';
        $this->file_prefix = uniqid();
    }
    public function __destruct() {
        $this->clear();
    }
    public function clear() {
        $this->cache_key = [];
        foreach ($this->file_handles as $file_handle) {
            isset($file_handle) && fclose($file_handle);
        }
        $this->delCacheDir($this->cache_dir);
        return true;
    }
    public function delete($key) {
        $key = $this->convertKey($key);
        unset($this->cache_key[$key]);
        return true;
    }
    public function deleteMultiple($keys) {
        foreach ($keys as $key) {
            $this->delete($key);
        }
        return true;
    }
    public function get($key, $default = null) {
        $key = $this->convertKey($key);
        if ($this->has($key)) {
            $seek = $this->cache_key[$key];
            if (array_key_exists($key, $this->cache) && $this->cache[$key]['seek'] == $seek) {
                return $this->cache[$key]['data'];
            }
            $fp = $this->getFileHandleByKey($key);
            $this->cache = [];
            fseek($fp, 0);
            while (! feof($fp)) {
                $data = fgets($fp);
                $data = json_decode(trim($data), 1);
                if ($data['key'] == $key && $data['seek'] == $seek) {
                    $default = unserialize($data['data']);
                }
                $this->cache[$data['key']] = [
                    'data' => unserialize($data['data']),
                    'seek' => $data['seek']
                ];
            }
        }
        return $default;
    }
    public function getMultiple($keys, $default = null) {
        $results = [];
        foreach ($keys as $key) {
            $results[$key] = $this->get($key, $default);
        }
        return $results;
    }
    public function has($key) {
        $key = $this->convertKey($key);
        return array_key_exists($key, $this->cache_key);
    }
    public function set($key, $value, $ttl = null) {
        $key = $this->convertKey($key);
        if ($this->has($key) && $this->get($key) == $value) {
            return true;
        }
        $fp = $this->getFileHandleByKey($key);
        fseek($fp, 0, SEEK_END);
        $seek = ftell($fp);
        $this->cache_key[$key] = $seek;
        fwrite($fp, json_encode([
                'key' => $key,
                'data' => serialize($value),
                'seek' => $seek
            ]) . PHP_EOL);
        unset($value);
        return true;
    }
    public function setMultiple($values, $ttl = null) {
        foreach ($values as $key => $value) {
            $this->set($key, $value);
        }
        return true;
    }
    private function convertKey($key) {
        return preg_replace('/^phpspreadsheet\./', '', $key); //remove prefix "phpspreadsheet."
    }
}
