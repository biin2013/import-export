<?php


use Biin2013\ImportExport\Import;

class ImportTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Calculation\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function testImport()
    {
        $path = './import/1.xlsx';
        $import = new Import($path);
        $import->setRows([3, 5, 9, 13, 15]);
        $import->setColumns(['a', 'b', 'c', 'd', 'e']);
        $import->setColumnFields(['A' => 'name', 'B' => 'title']);
        $import->setOnlyReadSheetIndex([0, 1, 3, 2]);
        print_r($import->getData());
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Calculation\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function testImport2()
    {
        $path = './import/2.xlsx';
        $import = new Import($path);
        print_r($import->getData());
    }
}