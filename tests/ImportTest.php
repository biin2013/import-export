<?php


class ImportTest extends \PHPUnit\Framework\TestCase
{
    public function testImport()
    {
        $path = './import/1.xlsx';
        $import = new ImportExport\Import($path);
        $import->setRows([3, 5, 9, 13, 15]);
        $import->setColumns(['a', 'b', 'c', 'd', 'e']);
        $import->setColumnFields(['A' => 'name', 'B' => 'title']);
        $import->setOnlyReadSheetIndex([0, 1, 3, 2]);
        print_r($import->getData());
    }
}