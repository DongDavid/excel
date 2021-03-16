<?php


use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase
{

    public function testNothing(){
        $a = 4;
        $this->assertIsInt($a);
    }
    //public function testCreateFile()
    //{
    //    $m = new \Dongdavid\Excel\Excel();
    //    //$this->expectException(\Dongdavid\Excel\Exceptions\IOException::class);
    //    //$this->expectDeprecationMessageMatches("/无写入权限/");
    //    $filename = './output/aa.xlsx';
    //    $m->setFilename($filename);
    //    $m->setCover(true);
    //    $this->assertEquals(true,$m->checkWriteable());
    //}

    //public function testWriteData()
    //{
    //    $m = new \Dongdavid\Excel\Excel(true);
    //    $rows = array(
    //        array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //        array('2003','=B1', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
    //        array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //        array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //        array('2003','=B1', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
    //        array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //        array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //        array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //    );
    //    $filename = "./output/test.xlsx";
    //    $m->init(true)->writeRows($rows);
    //    for ($i = 0 ;$i < 10;$i ++){
    //        $m->writeRows($rows);
    //    }
    //    $m->toFile($filename,true);
    //    $this->assertFileExists($filename);
    //}
    /**
     * @group add
     */
    //public function testQuickSave(){
    //    $data = [
    //       ['我是表头1','我是表头2','我是表头3','我是表头4','我是表头5','我是表头6','我是表头7','我是表头8','我是表头9','表头10']
    //    ];
    //    $line = [
    //        ['你好啊','我不好',23.434523,'TP239d2ojd0e3','40.0%','第六咧','哈哈哈','啊啊啊啊','0003232'],
    //        ['你好啊23','我不好',-23223,'TACOIS$$@#ojd0e3','40.0%','第六咧','哈哈哈','啊啊啊啊',10000000000032]
    //    ];
    //    $max = 10000;
    //    for ($i = 0;$i < $max;$i++){
    //        if ($i % 2 == 0){
    //            $line[1][0] = $i;
    //            $data[] = $line[1];
    //        }else{
    //            $line[0][0] = $i;
    //            $data[] = $line[0];
    //        }
    //    }
    //    $r = \Dongdavid\Excel\QuickStart::export('./output/quick'.$max.'.xlsx',$data);
    //    $this->assertEquals(true,$r);
    //}
    //public function testWriteFormate()
    //{
    //    $rows = array(
    //        array('2004','13323434354523243','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //        array('2005','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    //    );
    //    $title = ['111','112','113','114','115'];
    //    $m = new \Dongdavid\Excel\Excel(true);
    //    $filename = "./output/append.xlsx";
    //    $m->init(true)->writeRows($rows);
    //    $m->setTitle($title);
    //    $m->setFormate();
    //    $m->writeRows($rows);
    //    $m->toFile($filename,true);
    //    $this->assertFileExists($filename);
    //}
    /**
     * @group import
     */
    //public function testReader()
    //{
    //    $filename = "./output/quick100000.xlsx";
    //    $debug = [];
    //    $debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
    //    $data = \Dongdavid\Excel\QuickStart::getExcelRow($filename);
    //    $debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
    //    echo 'import:'.($debug[1][0] - $debug[0][0]).' -- '.$debug[1][1] .PHP_EOL;
    //    var_dump($data);
    //    $this->assertIsInt($data);
    //}
    /**
     * @group tfl
     */
    //public function testFirstLine()
    //{
    //    $filename = "./output/quick30.xlsx";
    //
    //    $res = \Dongdavid\Excel\QuickStart::getFirst($filename);
    //    $this->assertEquals(true,$res);
    //}
    /**
     * @group limit
     */
    public function testReaderLimit()
    {
        $filename = "./output/quick30.xlsx";
        $debug = [];
        $debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
        $data = [];
        for ($i = 0;$i < 1;$i++){
            $t = microtime(true);
            $tmp = \Dongdavid\Excel\QuickStart::importByLimit($filename,$i*10,($i+1)*100);
            //$data = array_merge($data,$tmp);
            $t1 = microtime(true);
            echo ($t1-$t) .'s'.PHP_EOL;
            echo '数量'.$tmp['total'].PHP_EOL;
            echo memory_get_usage()/1024/1024 .'MB'.PHP_EOL;
        }

        $debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
        echo PHP_EOL;
        //echo '数据数量'. $data['total'].PHP_EOL;
        echo 'import:'.($debug[1][0] - $debug[0][0]).' -- '.$debug[1][1] .PHP_EOL;
        //\Dongdavid\Excel\QuickStart::export('./limit.xlsx',$data['rows']);
        $this->assertArrayHasKey('total',$tmp);
    }

    /**
     * @group chunk
     * @throws \Dongdavid\Excel\Exceptions\IOException
     */
    //public function testReaderChunk()
    //{
    //    $filename = "./output/quick.xlsx";
    //    $debug = [];
    //
    //    $debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
    //    $data = \Dongdavid\Excel\QuickStart::importByChunk($filename,1,5000);
    //    $debug[] = [microtime(true),round(memory_get_usage()/1024/1024,2)];
    //    echo PHP_EOL;
    //    echo '数据数量'. $data['total'].PHP_EOL;
    //    echo 'import:'.($debug[1][0] - $debug[0][0]).' -- '.$debug[1][1] .PHP_EOL;
    //    \Dongdavid\Excel\QuickStart::export('./chunk.xlsx',$data['rows']);
    //    $this->assertArrayHasKey('total',$data);
    //}
}