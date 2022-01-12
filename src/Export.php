<?php


namespace Biin2013\ImportExport;

use Closure;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;

class Export
{
    private array $data;
    private array $enableTypes = ['Xlsx', 'Xls', 'Csv'];
    private string $defaultType = 'Xlsx';
    private string $type;
    private ?IWriter $writer = null;
    private ?Spreadsheet $spreadsheet = null;
    private array $title = [];
    private array $mergeCells = [];

    /**
     * Export constructor.
     * @param array $data
     * @param string $type
     */
    public function __construct(array $data = [], string $type = 'Xlsx')
    {
        $this->data = $data;
        $this->type = ucfirst($type);
    }

    /**
     * @param int $worksheetIndex
     * @param int $startColumn
     * @param int $startRow
     * @param int $endColumn
     * @param int $endRow
     * @return Export
     */
    public function addMergeCells(
        int $worksheetIndex,
        int $startColumn,
        int $startRow,
        int $endColumn,
        int $endRow
    ): self
    {
        if (!isset($this->mergeCells[$worksheetIndex])) {
            $this->mergeCells[$worksheetIndex] = [];
        }
        $this->mergeCells[$worksheetIndex][] = [
            $startColumn,
            $startRow,
            $endColumn,
            $endRow
        ];

        return $this;
    }

    /**
     * @param array $data
     * @return $this
     */
    public function data(array $data): self
    {
        $this->data = $data;

        return $this;
    }

    /**
     * set xlsx type
     * @return $this
     */
    public function xlsx(): self
    {
        $this->type = 'Xlsx';

        return $this;
    }

    /**
     * set xls type
     * @return $this
     */
    public function xls(): self
    {
        $this->type = 'Xls';

        return $this;
    }

    /**
     * set csv type
     * @return $this
     */
    public function csv(): self
    {
        $this->type = 'Csv';

        return $this;
    }

    /**
     * set title
     * @param array $title
     * [
     *  // sheet 1
     *  [
     *  'merge_row' => 2,
     *  'children' => [
     *          // field 1
     *          [
     *              'name' => 'ID'
     *          ],
     *          // field 2
     *          [
     *              'name' => '用户信息',
     *              'children' => ['姓名', '性别']
     *          ],
     *          // sheet 3
     *          [...]
     *      ]
     *  ],
     *  // sheet 2
     *  [...]
     * ]
     * @return $this
     */
    public function setTitle(array $title): self
    {
        $this->title = $title;

        return $this;
    }

    /**
     * @param string $rootPath
     * @param Closure|string|false $datePath
     * @param ?string $fileName
     * @param array $sheetConfig
     * @return array
     * @throws SpreadsheetException
     * @throws WriterException
     */
    public function save(
        string  $rootPath,
                $datePath = false,
        ?string $fileName = null,
        array   $sheetConfig = []
    ): array
    {
        $filePath = $this->resolvePath($rootPath, $datePath);
        $fileName = $this->resolveType()->resolveFileName($fileName);
        $title = $this->resolveTitle();
        $spreadsheet = $this->getSpreadsheet();
        $this->resolveSheetConfig($spreadsheet, $sheetConfig);

        foreach ($this->data as $index => $data) {
            if ($index > 0) {
                $spreadsheet->createSheet($index);
            }
            $worksheet = $spreadsheet->setActiveSheetIndex($index);
            $this->resolveMergeCell($worksheet, $index);
            $worksheet->setTitle($data['sheet_name']);
            $this->saveTitle($title[$index], $worksheet);
            $row = ($this->title[$index]['merge_row'] ?? 1) + 1;
            foreach ($data['data'] as $value) {
                foreach ($value as $field => $val) {
                    $fieldConfig = $title[$index][$field];
                    $fieldConfig['row'] = $row;
                    $currentColumn = ord($fieldConfig['column']) - ord('A') + 1;
                    $cell = $worksheet->getCellByColumnAndRow($currentColumn, $row);
                    $this->resolveCellConfig($worksheet, $fieldConfig, $cell, $row);
                    if (is_callable($val)) {
                        $result = call_user_func_array(
                            $val,
                            [$cell, $worksheet, $row, $currentColumn, $data, $this->data]
                        );
                    } elseif (is_callable($fieldConfig['format'] ?? null)) {
                        $result = call_user_func_array(
                            $fieldConfig['format'],
                            [$val, $fieldConfig, $cell, $worksheet, $value, $data, $this->data]
                        );
                    } else {
                        $result = call_user_func_array(
                            [$this, 'defaultCellFormat'],
                            [$val, $fieldConfig['format'] ?? null, $field, $cell, $worksheet]
                        );
                    }
                    if (!is_null($result)) {
                        $cell->setValue($result);
                    }

                }
                $row++;
            }
        }
        $this->getWriter()->save($filePath . $fileName);

        return [
            'path' => $filePath,
            'filename' => $fileName
        ];
    }

    /**
     * get spreadsheet instance
     * @return Spreadsheet
     */
    public function getSpreadsheet(): Spreadsheet
    {
        if (is_null($this->spreadsheet)) {
            $this->spreadsheet = new Spreadsheet();
        }

        return $this->spreadsheet;
    }

    /**
     * @param Spreadsheet|null $spreadsheet
     * @return IWriter
     * @throws WriterException
     */
    public function getWriter(?Spreadsheet $spreadsheet = null): IWriter
    {
        if (is_null($this->writer)) {
            $this->writer = IOFactory::createWriter(
                $spreadsheet ?? $this->getSpreadsheet(), $this->type
            );
        }

        return $this->writer;
    }

    /**
     * @param Worksheet $worksheet
     * @param int $index
     */
    protected function resolveMergeCell(Worksheet $worksheet, int $index)
    {
        if (empty($this->mergeCells[$index])) return;

        foreach ($this->mergeCells[$index] as $v) {
            $worksheet->mergeCellsByColumnAndRow(...$v);
        }
    }

    /**
     * resolve type
     * @return $this
     */
    protected function resolveType(): self
    {
        if (!in_array($this->type, $this->enableTypes)) {
            $this->type = $this->defaultType;
        }

        return $this;
    }

    /**
     * resolve path
     * @param string $rootPath
     * @param bool|string|Closure $datePath
     * @return string
     */
    protected function resolvePath(string $rootPath, $datePath = false): string
    {
        $path = $rootPath . DIRECTORY_SEPARATOR;
        $path .= is_bool($datePath)
            ? ($datePath ? date('Ymd') : '')
            : (
            is_callable($datePath)
                ? call_user_func($datePath)
                : date($datePath)
            );
        if (!is_dir($path)) {
            mkdir($path, 0777, true);
        }

        return rtrim($path, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR;
    }

    /**
     * resolve file name
     * @param ?string $fileName
     * @return string
     */
    protected function resolveFileName(string $fileName = null): string
    {
        $name = $fileName ?? md5(time());

        return $name . '.' . strtolower($this->type);
    }

    /**
     * resolve title
     * @return array
     */
    protected function resolveTitle(): array
    {
        $title = [];
        foreach ($this->title as $index => &$item) {
            $item['merge_row'] = $this->resolveTitleMergeRow($item['children']);
            $title[$index] = $this->resolveTitleItem(
                $item['children'],
                $item['merge_row']
            );
        }

        return $title;
    }

    /**
     * @param array $data
     * @return int
     */
    protected function resolveTitleMergeRow(array $data): int
    {
        $row = 1;
        $subRow = [0];
        foreach ($data as $v) {
            if (!empty($v['children'])) {
                $subRow[] = $this->resolveTitleMergeRow($v['children']);
            }
        }
        $row += max($subRow);
        return $row;
    }

    /**
     * resolve title item
     * @param array $item
     * @param int $mergeRow
     * @param string $column
     * @param int $currentRow
     * @return array
     */
    protected function resolveTitleItem(
        array  $item,
        int    $mergeRow,
        string $column = 'A',
        int    $currentRow = 1
    ): array
    {
        $data = [];
        foreach ($item as $value) {
            $data[$value['field']] = $value;
            $data[$value['field']]['column'] = $column;
            if (!empty($value['children'])) {
                $data[$value['field']] = array_merge(
                    $data[$value['field']],
                    $this->resolveTitleFieldColumnAndRow(
                        $column,
                        count($value['children']),
                        $currentRow,
                        1
                    ));
                $mergeRow--;
                if ($mergeRow < 1) {
                    return [];
                }
                $data = array_merge($data, $this->resolveTitleItem(
                    $value['children'],
                    $mergeRow,
                    $column,
                    ++$currentRow
                ));
            } else {
                $data[$value['field']] = array_merge(
                    $data[$value['field']],
                    $this->resolveTitleFieldColumnAndRow(
                        $column++,
                        1,
                        $currentRow,
                        $mergeRow
                    )
                );
            }
        }

        return $data;
    }

    /**
     * resolve title field column and row
     * @param string $column
     * @param int $columnStep
     * @param int $row
     * @param int $rowStep
     * @return array
     */
    protected function resolveTitleFieldColumnAndRow(
        string $column,
        int    $columnStep,
        int    $row,
        int    $rowStep
    ): array
    {
        if ($columnStep > 1 || $rowStep > 1) {
            $maxColumn = $columnStep > 1 ? chr(ord($column) + $columnStep - 1) : $column;
            $maxRow = $rowStep > 1 ? $row + $rowStep - 1 : $row;
            return [
                'merge' => true,
                'cell' => $column . $row . ':' . $maxColumn . $maxRow
            ];
        } else {
            return [
                'merge' => false,
                'cell' => $column . $row
            ];
        }
    }

    /**
     * save title
     * @param array $title
     * @param Worksheet $worksheet
     * @throws SpreadsheetException
     */
    protected function saveTitle(array $title, Worksheet $worksheet)
    {
        foreach ($title as $value) {
            if ($value['merge']) {
                $cells = explode(':', $value['cell']);
                $currentCell = $cells[0];
                $worksheet->mergeCells($value['cell']);
            } else {
                $currentCell = $value['cell'];
            }
            $worksheet->getStyle($currentCell)
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER)
                ->setVertical(Alignment::VERTICAL_CENTER);
            $worksheet->getCell($currentCell)->setValue($value['name']);
        }
    }

    /**
     * default cell format
     * @param mixed $value
     * @param mixed $format
     * @param string $field
     * @param Cell $cell
     * @param Worksheet $worksheet
     * @return mixed
     */
    protected function defaultCellFormat(
        $value,
        $format,
        string $field,
        Cell $cell,
        Worksheet $worksheet
    )
    {
        return $value;
    }

    /**
     * resolve cell config
     * @param Worksheet $worksheet
     * @param array $config
     * @param Cell $cell
     * @param int $row
     */
    protected function resolveCellConfig(
        Worksheet $worksheet,
        array     $config,
        Cell      $cell,
        int       $row
    )
    {
        if (isset($config['width'])) {
            $worksheet->getColumnDimension($config['column'])
                ->setWidth($config['width']);
        }
        if (isset($config['color'])) {
            $cell->getStyle()->getFont()->getColor()->setRGB(ltrim($config['color'], '#'));
        }
        $cell->getStyle()->getAlignment()->setHorizontal($config['horizontal'] ?? Alignment::HORIZONTAL_LEFT);
        $cell->getStyle()->getAlignment()->setVertical($config['vertical'] ?? Alignment::VERTICAL_CENTER);
    }

    /**
     * resolve sheet config
     * @param Spreadsheet $spreadsheet
     * @param array $sheetConfig
     * @throws SpreadsheetException
     */
    protected function resolveSheetConfig(Spreadsheet $spreadsheet, array $sheetConfig)
    {
        $sheetConfig['font_size'] = $sheetConfig['font_size'] ?? 16;
        $spreadsheet->getDefaultStyle()->getFont()->setSize($sheetConfig['font_size']);
    }
}