<h1 align="center"> excel </h1>

<p align="center"> .</p>


## Installing

```shell
$ composer require dongdavid/excel -vvv
```

## Usage

### 导入数据  

```php

```
\Dongdavid\Excel\QuickStart::getExcelRow($filename); // 获取当前sheet的行数
\Dongdavid\Excel\QuickStart::importByLimit($filename,$startRow,$endRow);

```php
// 读完后会自动释放内存
// 10万条数据 一次1000条 读一次要4秒， 14MB内存 全部读完要7分钟， 总消耗内存在40MB左右
// 10万条数据 一次3000条 读一次5秒 14MB 
$filename = "./output/quick100000.xlsx";
$debug = [];
$debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
$data = [];
for ($i = 0;$i < 100;$i++){
    $t = microtime(true);
    $tmp = \Dongdavid\Excel\QuickStart::importByLimit($filename,$i*10,($i+1)*1000);
    //$data = array_merge($data,$tmp);
    $t1 = microtime(true);
    echo $t1-$t .'s'.PHP_EOL;
    echo '数量'.$tmp['total'].PHP_EOL;
    echo memory_get_usage()/1024/1024 .'MB'.PHP_EOL;
}

```
### 导出数据  

更多的用法可以参考[xlsxwriter](https://packagist.org/packages/mk-j/php_xlsxwriter)

```php
// 10万行 18秒 80MB内存
// 1万行 2秒 14MB内存
$data = [
   ['我是表头1','我是表头2','我是表头3','我是表头4','我是表头5','我是表头6','我是表头7','我是表头8','我是表头9']
];
$line = [
    ['你好啊','我不好',23.434523,'TP239d2ojd0e3','40.0%','第六咧','哈哈哈','啊啊啊啊','0003232'],
    ['你好啊23','我不好',-23223,'TACOIS$$@#ojd0e3','40.0%','第六咧','哈哈哈','啊啊啊啊',10000000000032]
];
for ($i = 0;$i < 10000;$i++){
    $data[] = $line[0];
    $data[] = $line[1];
    $data[] = $line[0];
    $data[] = $line[1];
    $data[] = $line[0];
    $data[] = $line[1];
    $data[] = $line[1];
    $data[] = $line[0];
    $data[] = $line[0];
    $data[] = $line[1];
}
// 导出到文件
\Dongdavid\Excel\QuickStart::export('./output/quick.xlsx',$data);
// 导出到浏览器 会自己设置响应头
\Dongdavid\Excel\QuickStart::exportDown('quick.xlsx',$data);
```

```php
// 复杂调用
$rows = [
    [],
    [],
];
$col = 10; //列数
$sheetName = 'Sheet1'; //非必填 默认为Sheet1
$excel = new \Dongdavid\Excel\Excel();
$excel->init(false);
$excel->setTitle(['表头1','表头2','表头3','表头4']);
// 如果调用列setTitle 则无需在调用setFormate 
$excel->setFormate([],count($col)); // 设置所有列为文本格式
foreach($rows as $row){
    $excel->writeRow($row,$sheetName);
}
$excel->toFile("./output/result.xlsx");

```
每一行都设置不同的格式
```php
$row = [
    
];
$excel = new \Dongdavid\Excel\Excel();
$excel->init(false);
//$excel->writeRow($row);
//$excel->setFormate(['1'=>'string','23438932'=>'string','32'=>'0.00%']);
$excel->writeRow($row);
```
## Contributing

You can contribute in one of three ways:

1. File bug reports using the [issue tracker](https://github.com/dongdavid/excel/issues).
2. Answer questions or fix bugs on the [issue tracker](https://github.com/dongdavid/excel/issues).
3. Contribute new features or update the wiki.

_The code contribution process is not very formal. You just need to make sure that you follow the PSR-0, PSR-1, and PSR-2 coding guidelines. Any new code contributions must be accompanied by unit tests where applicable._

## License

MIT
