<?php


namespace Dongdavid\Excel\XLSReader;


use Dongdavid\Excel\Exceptions\Exception;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class XLSReaderFIrstLine implements IReadFilter
{
    public $record = [];
    public function readCell($column, $row, $worksheetName = '')
    {
        // TODO: Implement readCell() method.
        if ($row > 1){
            //这里需要重写Reader类才能实现
            //throw new Exception("强制跳出");
        }
        $this->record[$row] = $row;
        return true;
    }
}