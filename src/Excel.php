<?php


namespace Dongdavid\Excel;


use Dongdavid\Excel\Exceptions\Exception;
use Dongdavid\Excel\Exceptions\IOException;
use Dongdavid\Excel\XLSWriter\XLSXWriter;

class Excel
{
    private $filename;
    private $cover = false;
    private $title;
    private $writer;
    private $debugInfo = [];

    public function __construct($debug = false)
    {
        $this->debug = $debug;
    }

    public function __destruct()
    {
        // TODO: Implement __destruct() method.
        if ($this->debugInfo['debug']) {
            $this->debugInfo['end_time'] = microtime(true);
            $this->debugInfo['end_memory'] = memory_get_usage();
            $this->debugInfo['spend_time'] = $this->debugInfo['end_time'] - $this->debugInfo['start_time'];
            $this->debugInfo['max_memory'] = memory_get_peak_usage();
            $this->debugInfo['start_memory'] = $this->debugInfo['start_memory'] / 1024 / 1024 .'MB';
            $this->debugInfo['end_memory'] = $this->debugInfo['end_memory'] / 1024 / 1024 .'MB';
            $this->debugInfo['max_memory'] = $this->debugInfo['max_memory'] / 1024 / 1024 .'MB';
            var_dump($this->debugInfo);
        }
    }

    /**
     * 是否覆盖原文件
     * @param  bool  $cover
     */
    public function setCover(bool $cover)
    {
        $this->cover = $cover;
        return $this;
    }

    /**
     * @param  mixed  $title
     */
    public function setTitle($title,$sheet = 'Sheet1')
    {
        $this->writer->writeSheetHeader($sheet,$title);
        return $this;
    }

    /**
     * 设置格式 如果全部是string且有设置标题，则传入空数组，自动设置
     * 在写入每行数据前执行， 可以为每行设置不同的格式
     * @param  array  $formate
     * @param  int  $len 列数
     * @example ['string','integer','0','0.00','0%','0.00%','dollar','euro','date','datetime','YYYY-MM-DD','D-MMM-YYYY HH:MM AM/PM','HH:MM:SS']
     */
    public function setFormate($formate = [],$len = 0){
        if (empty($formate)){
            if ($len === 0){
                $len = count($this->title);
            }
            $formate = array_fill(0,$len,'string');
        }
        $this->writer->writeSheetHeader('Sheet1',$formate);
    }
    /**
     * @return mixed
     */
    public function getFilename()
    {
        return $this->filename;
    }

    /**
     * @param  mixed  $filename
     */
    public function setFilename($filename)
    {
        $this->filename = $filename;
        return $this;
    }

    /**
     * 是否开启调试信息
     * @param  false  $debug
     * @return $this
     */
    public function init($debug = false)
    {
        $this->debugInfo = [
            'debug' => $debug,
            'start_time' => microtime(true),
            'start_memory' => memory_get_usage(),
            'end_time' => 0,
            'end_memory' => 0,
            'spend_time' => 0,
            'spend_memory' => 0,
        ];
        $this->writer = new XLSXWriter();
        return $this;
    }

    /**
     * 写入数据 多行数据
     * @param $rows
     * @param  string  $sheetName
     * @return $this
     */
    public function writeRows($rows, $sheetName = 'Sheet1')
    {
        foreach ($rows as $row) {
            $this->writer->writeSheetRow($sheetName, $row);
        }
        return $this;
    }
    /**
     * 写入数据 单行数据
     * @param $row
     * @param  string  $sheetName
     * @return $this
     */
    public function writeRow($row, $sheetName = 'Sheet1')
    {
        $this->writer->writeSheetRow($sheetName, $row);
        return $this;
    }

    public function checkWriteable()
    {
        if (empty($this->filename)) {
            throw new IOException("文件名不能为空");
        }
        $dir = dirname($this->filename);
        if (!is_dir($dir)) {
            $r = mkdir($dir, 0777, true);
            if (!$r) {
                throw new IOException("创建目录失败:".$dir);
            }
        }
        //$filepath = realpath(dirname($this->filename));

        if (file_exists($this->filename)) {
            if ($this->cover) {
                if (!is_writable($this->filename)) {
                    throw new IOException("没有写入权限:".$this->filename);
                }
                @unlink($this->filename);
            } else {
                throw new IOException("文件已存在:".$this->filename);
            }
        } else {
            if ($fp = @fopen($this->filename, 'w')) {
                @fclose($fp);
                @unlink($this->filename);
            } else {
                throw new IOException("没有写入权限:".$this->filename);
            }
        }
        return true;
    }

    /**
     * 输出到本地文件
     * @param  string  $filename  文件名
     * @param  false  $cover  是否覆盖原文件
     * @throws Exception
     * @throws IOException
     */
    public function toFile($filename, $cover = false)
    {
        if (!$this->writer) {
            throw new Exception("请先执行init方法");
        }
        $this->filename = $filename;
        $this->cover = $cover;
        $this->checkWriteable();

        $this->writer->writeToFile($filename);
    }
    public function toWeb($filename){
        header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header('Content-Transfer-Encoding: binary');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        $this->writer->writeToStdOut();
        exit;
    }
}