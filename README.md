# kxbrand/phplib

凯旋工具包


## Composer 安装

```
composer require kxbrand/phplib
```

## Kxtools 使用
```php
use kxbrand\phplib\Tools as Tools;
$kxtools = new Tools();
$data = ['name'=>'kxbrand','email'=>'kxbrand@qq.com'];
$arr2xml = $kxtools->arr2xml($data);

//<xml>
//	<name><![CDATA[kxbrand]]></name>
//	<email><![CDATA[kxbrand@qq.com]]></email>
//</xml>

```