<?php


namespace ImportExport;


use Closure;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Reader\IReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class Import
{
    private string $file;
    private ?IReader $reader = null;
    private int $startRow = 1;
    private ?int $endRow;
    private string $startColumn = 'A';
    private ?string $endColumn;
    private array $rows;
    private array $columns;
    private bool $readDataOnly = false;
    private ?Spreadsheet $spreadsheet = null;
    private $onlyReadSheetName;
    private $onlyReadSheetIndex;
    private array $columnFields;

    public function __construct(
        $file,
        $startRow = 1,
        $endRow = null,
        $startColumn = 'A',
        $endColumn = null,
        $columnFields = [])
    {
        $this->file = $file;
        $this->startRow = $startRow;
        $this->endRow = $endRow;
        $this->startColumn = $startColumn;
        $this->endColumn = $endColumn;
        $this->columnFields = $columnFields;
    }

    /**
     * set start row
     * @param int $row
     * @return $this
     */
    public function setStartRow(int $row): self
    {
        $this->startRow = $row < 1 ? 1 : $row;

        return $this;
    }

    /**
     * set end row
     * @param int $row
     * @return $this
     */
    public function setEndRow(int $row): self
    {
        $this->endRow = $row;

        return $this;
    }

    /**
     * set start column
     * @param string $column
     * @return $this
     */
    public function setStartColumn(string $column): self
    {
        $this->startColumn = ucwords($column);

        return $this;
    }

    /**
     * set end column
     * @param string $column
     * @return $this
     */
    public function setEndColumn(string $column): self
    {
        $this->endColumn = ucwords($column);

        return $this;
    }

    /**
     * set rows
     * @param array $rows
     * @return $this
     */
    public function setRows(array $rows): self
    {
        $this->rows = $rows;

        return $this;
    }

    /**
     * set columns
     * @param array $column
     * @return $this
     */
    public function setColumns(array $column): self
    {
        $this->columns = array_map('ucwords', $column);

        return $this;
    }

    /**
     * set column relation field
     * @param array $fields
     * @return $this
     */
    public function setColumnFields(array $fields): self
    {
        $this->columnFields = $fields;

        return $this;
    }

    /**
     * @return IReader
     * @throws Exception
     */
    public function getReader(): IReader
    {
        if (is_null($this->reader)) {
            $fileType = IOFactory::identify($this->file);
            $this->reader = IOFactory::createReader(ucfirst($fileType));
            $this->reader->setReadDataOnly($this->readDataOnly);
        }

        return $this->reader;
    }

    /**
     * list worksheet names
     * @return mixed
     * @throws Exception
     */
    public function listWorksheetNames()
    {
        return $this->getReader()->listWorksheetNames($this->file);
    }

    /**
     * list worksheet info
     * @return mixed
     * @throws Exception
     */
    public function listWorksheetInfo()
    {
        return $this->getReader()->listWorksheetInfo($this->file);
    }

    /**
     * get spread
     * @return Spreadsheet
     * @throws Exception
     */
    public function getSpreadsheet(): Spreadsheet
    {
        if (is_null($this->spreadsheet)) {
            $this->spreadsheet = $this->getReader()->load($this->file);
        }

        return $this->spreadsheet;
    }

    /**
     * set only read sheet name
     * @param string|array $name
     * @return $this
     */
    public function setOnlyReadSheetName($name): self
    {
        $this->onlyReadSheetName = (array)$name;

        return $this;
    }

    /**
     * set only read sheet index
     * @param string|array $index
     * @return $this
     */
    public function setOnlyReadSheetIndex($index): self
    {
        $this->onlyReadSheetIndex = (array)$index;

        return $this;
    }

    /**
     * get data
     * @param Closure|null $callback
     * @return array
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Calculation\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function getData(Closure $callback = null): array
    {
        $this->getSpreadsheet();

        $sheetNames = $this->spreadsheet->getSheetNames();
        $sheetIndex = $this->resolveReadSheetIndex($sheetNames);

        $data = [];
        foreach ($sheetIndex as $index) {
            $worksheet = $this->spreadsheet->setActiveSheetIndex($index);
            $data[] = [
                'sheet_name' => $sheetNames[$index],
                'data' => $this->resolveCell($worksheet, $callback)
            ];
        }

        return $data;
    }

    /**
     * resolve read sheet index
     * @param array $sheetNames
     * @return array
     */
    protected function resolveReadSheetIndex(array $sheetNames): array
    {
        if (empty($this->onlyReadSheetIndex) && empty($this->onlyReadSheetName)) {
            return array_keys($sheetNames);
        }

        $intersectIndex = empty($this->onlyReadSheetIndex)
            ? array_keys($sheetNames)
            : array_intersect($this->onlyReadSheetIndex, array_keys($sheetNames));

        if (empty($this->onlyReadSheetName)) {
            return $intersectIndex;
        }

        $intersectName = array_intersect($this->onlyReadSheetName, $sheetNames);
        $intersectIndex2 = array_values(
            array_intersect_key(
                array_flip($sheetNames),
                array_flip($intersectName)
            )
        );

        return empty($this->onlyReadSheetIndex)
            ? array_intersect($intersectIndex, $intersectIndex2)
            : array_unique(array_merge($intersectIndex, $intersectIndex2));
    }

    /**
     * resolve cell data
     * @param Worksheet $worksheet
     * @param ?Closure $callback
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Calculation\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function resolveCell(Worksheet $worksheet, ?Closure $callback): array
    {
        $rows = $this->getReadRows($worksheet);
        $columns = $this->getReadColumns($worksheet);
        $callback = $callback ?: function ($field, $value, $cell, $worksheet) {
            return $value;
        };
        $data = [];
        foreach ($rows as $row) {
            $item = [];
            foreach ($columns as $column) {
                $field = $this->columnFields[$column] ?? $column;
                $cell = $worksheet->getCell($column . $row);
                $value = $cell->getCalculatedValue();
                $item[$field] = call_user_func($callback, $field, $value, $cell, $worksheet);
            }
            $data[] = $item;
        }

        return $data;
    }

    /**
     * get read rows
     * @param Worksheet $worksheet
     * @return array
     */
    protected function getReadRows(Worksheet $worksheet): array
    {
        if (!empty($this->rows)) {
            return $this->rows;
        }

        $end = $this->endRow ?: $worksheet->getHighestDataRow();

        return range($this->startRow, $end);
    }

    /**
     * get read columns
     * @param Worksheet $worksheet
     * @return array
     */
    protected function getReadColumns(Worksheet $worksheet): array
    {
        if (!empty($this->columns)) {
            return $this->columns;
        }

        $end = $this->endColumn ?: ucwords($worksheet->getHighestDataColumn());
        return range($this->startColumn, $end);
    }
}