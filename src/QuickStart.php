<?php


namespace Dongdavid\Excel;


use Dongdavid\Excel\Exceptions\Exception;
use Dongdavid\Excel\Exceptions\IOException;
use Dongdavid\Excel\XLSReader\XLSReaderChunk;
use Dongdavid\Excel\XLSReader\XLSReaderFIlter;
use Dongdavid\Excel\XLSReader\XLSReaderNone;
use PhpOffice\PhpSpreadsheet\IOFactory;

class QuickStart
{
    public static function exportDown($filename,$data,$title = [],$sheetName = 'Sheet1',)
    {

        if(empty($data)){
            return '请传入非空数组';
        }
        $excel = new Excel();
        try {
            $excel->init(false);
            if (!empty($title)){
                $excel->setTitle($title);
            }else{
                $title = $data[0];
                unset($data[0]);
            }
            $header = [];
            foreach ($title as $item) {
                $header[$item] = 'string';
            }
            $excel->setTitle($header);
            $excel->writeRows($data,$sheetName);
            $excel->toWeb($filename);
        }catch (Exception $e){
            return $e->getMessage();
        }
    }
    /**
     * 将数据写入excel，且格式为字符串格式保存
     * @param $filename 保存文件名
     * @param $data 保存数据
     * @param $title 表头
     * @param $cover 是否覆盖原文件
     */
    public static function export($filename,$data,$title = [],$sheetName = 'Sheet1',$cover = true)
    {
        if(empty($data)){
            return '请传入非空数组';
        }
        $excel = new Excel();
        try {
            $excel->init(false);
            if (!empty($title)){
                $excel->setTitle($title);
            }else{
                $title = $data[0];
                unset($data[0]);
            }
            $header = [];
            foreach ($title as $item) {
                $header[$item] = 'string';
            }
            $excel->setTitle($header);
            $excel->writeRows($data,$sheetName);
            $excel->toFile($filename,true);
        }catch (Exception $e){
            return $e->getMessage();
        }
        if (file_exists($filename)){
            return true;
        }
        return "保存失败";
    }

    /**
     * 读取excel中的指定行
     * @param $filename
     * @param  int  $startRow
     * @param  int  $endRow
     * @return array
     * @throws IOException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function importByLimit($filename,$startRow = 1,$endRow = 1)
    {

        $type = self::getFileType($filename);

        $reader = IOFActory::createReader($type);

        $filterSubset = new XLSReaderFIlter($startRow,$endRow,range('A','J'));
        $reader->setReadFilter($filterSubset);
        $reader->setReadDataOnly(true);
        $excel = $reader->load($filename);
        $sheet = $excel->getActiveSheet();
        $total = $sheet->getHighestRowAndColumn();
        $data = [];
        for ($j = 0;$j < $total['row'];$j++){
            $tmp = [];
            for ($c = 'A';$c<$total['column'];$c++){
                $tmp[] = $sheet->getCell($c.$j)->getValue();
            }
            $data[] = $tmp;
        }
        $excel->disconnectWorksheets();
        $filterSubset = null;
        $sheet = null;
        $excel = null;
        $reader = null;
        //var_dump($rows);
        //echo memory_get_peak_usage()/1024/1024 . 'MB';
        return ['total'=>count($data),'rows'=>$data];
    }
    // 高效获取前N行
    //public static function getFirst($filename)
    //{
    //    $type = self::getFileType($filename);
    //    $reader = IOFActory::createReader($type);
    //    $reader->setReadDataOnly(true);
    //    // 不读取任何内容 这里需要重写Reader类才能实现
    //    $filter = new XLSReaderFIrstLine();
    //
    //    $reader->setReadFilter($filter);
    //    try {
    //        $excel = $reader->load($filename);
    //    }catch(Exception $e){
    //        echo $e->getMessage();
    //    }
    //    $sheet = $excel->getActiveSheet();
    //    $excel->disconnectWorksheets();
    //    $sheet = null;
    //    $excel = null;
    //    $reader = null;
    //    var_dump($filter->record);
    //    return true;
    //}
    public static function getFileType($filename){
        if (!file_exists($filename)){
            throw new IOException("文件不存在".$filename);
        }
        $tmp = explode('.',$filename);
        if (count($tmp) > 1){
            $type  = end($tmp);
        }else{
            $info = finfo_open(FILEINFO_MIME);
            $type  = finfo_file($info,$filename);
            finfo_close($info);
        }
        $type = ucfirst($type);
        return $type;
    }

    /**
     * 获取excel的行数
     * @param $filename
     * @param false $all 默认只返回当前sheet的行数， 传入true则返回全部sheet
     * @return array
     * @throws IOException
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function getExcelRow($filename,$all = false)
    {
        //$filename = "./output/quick30.xlsx";
        $type = self::getFileType($filename);
        $reader = IOFActory::createReader($type);
        $reader->setReadDataOnly(true);
        // 不读取任何内容
        $filter = new XLSReaderNone();

        $reader->setReadFilter($filter);
        $excel = $reader->load($filename);
        $sheet = $excel->getActiveSheet();
        $title = $sheet->getTitle();
        $excel->disconnectWorksheets();
        $sheet = null;
        $excel = null;
        $reader = null;
        if ($all){
            return $filter->record;
        }else{
            return $filter->record[$title];
        }
    }

    public static function importByChunk($filename,$start = 5000,$end = 100000)
    {
        $chunkSize = 1000;
        $type = self::getFileType($filename);

        $reader = IOFActory::createReader($type);

        $chunkFilter = new XLSReaderChunk();
        $reader->setReadDataOnly(true);
        $reader->setReadFilter($chunkFilter);
        $data = [];
        for ($i = $start;$i <= $end;$i += $chunkSize){
            $chunkFilter->setRows($i,$chunkSize);
            $excel = $reader->load($filename);
            $sheet = $excel->getActiveSheet();
            $total = $sheet->getHighestRowAndColumn();

            for ($j = $i;$j <= $total['row'];$j++){
                if ($total['row'] > $end){
                    break;
                }
                $tmp = [];
                for ($c = 'A';$c<=$total['column'];$c++){
                    $tmp[] = $sheet->getCell($c.$j)->getValue();
                }
                $data[] = $tmp;
            }
            //echo memory_get_usage()/1024/1024 .'MB'.PHP_EOL;
            //echo "start $i;size $chunkSize-$j-".count($data).PHP_EOL;
            //if (memory_get_usage() > 104217728){
                //break;
            //}
            $excel->disconnectWorksheets();
            //$excel->discardMacros();
            $sheet = null;
            $excel = null;
        }
        $reader = null;
        //echo memory_get_peak_usage()/1024/1024 . 'MB'.PHP_EOL;

        return ['total'=>count($data),'rows'=>$data];
    }
}