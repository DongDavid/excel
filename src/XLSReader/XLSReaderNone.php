<?php


namespace Dongdavid\Excel\XLSReader;


use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class XLSReaderNone implements IReadFilter
{
    public $record = array();
    private $lastRow = '';

    public function readCell($column, $row, $worksheetName = '')
    {
        // TODO: Implement readCell() method.
        if (isset($this->record[$worksheetName]) ) {
            if ($this->lastRow != $row) {
                $this->record[$worksheetName] ++;
                $this->lastRow = $row;
            }
        } else {
            $this->record[$worksheetName] = 1;
            $this->lastRow = $row;
        }
        return false;
    }
}