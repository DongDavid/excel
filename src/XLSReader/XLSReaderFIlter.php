<?php


namespace Dongdavid\Excel\XLSReader;


use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;

class XLSReaderFIlter implements IReadFilter
{

    /**  Get the list of rows and columns to read  */
    public function __construct($startRow, $endRow, $columns) {
        $this->startRow = $startRow;
        $this->endRow   = $endRow;
        $this->columns  = $columns;
    }

    public function readCell($column, $row, $worksheetName = '') {
        //  Only read the rows and columns that were configured
        if ($row >= $this->startRow && $row <= $this->endRow) {
            if (!empty($this->columns)&&in_array($column,$this->columns)) {
                return true;
            }
        }
        return false;
    }
}