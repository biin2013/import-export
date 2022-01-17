<?php


use Biin2013\ImportExport\Export;

class ExportTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function testExport()
    {
        //$export = new Export($this->getData());
        $export = new Export();
        $path = $export->setData($this->getData())
            /*->addMergeCells(0, 1, 4, 1, 5)
            ->addMergeCells(1, 2, 2, 2, 3)*/
            ->setMergeCells([
                0 => [[1, 4, 1, 5]],
                1 => [[2, 2, 2, 3]]
            ])
            ->setTitle($this->getTitle())
            ->build('./export/test_export')
            ->save();
        print_r($path);
    }

    private function getTitle(): array
    {
        return [
            [
                'children' => [
                    [
                        'field' => 'id',
                        'name' => 'ID'
                    ],
                    [
                        'field' => 'userinfo',
                        'name' => '用户信息',
                        'children' => [
                            [
                                'field' => 'username',
                                'name' => '姓名'
                            ],
                            [
                                'field' => 'sex',
                                'name' => '性别',
                                'color' => '#0000ff',
                                'horizontal' => 'right',
                                'vertical' => 'center',
                                'format' => function ($value, $config, $cell) {
                                    return $value == 1
                                        ? '男'
                                        : ($value == 2 ? '女' : '保密');
                                }
                            ]
                        ]
                    ]
                ]
            ],
            [
                'children' => [
                    [
                        'field' => 'id',
                        'name' => 'ID'
                    ],
                    [
                        'field' => 'date',
                        'name' => '日期',
                        'width' => 20
                    ]
                ]
            ]
        ];
    }

    private function getData(): array
    {
        return [
            [
                'sheet_name' => '导出数据表1',
                'data' => [
                    [
                        'id' => 1,
                        // this callback will override title format
                        // when all cells are the same format, use title format
                        'username' => function ($cell, $worksheet, $row, $column) {
                            // merge cell
                            //$worksheet->mergeCellsByColumnAndRow($column, $row, $column + 1, $row + 2);
                            return '张三' . '-' . $row . '-' . $column;
                        },
                        'sex' => 1
                    ],
                    [
                        'id' => 2,
                        'username' => '小丽',
                        'sex' => 2
                    ],
                    [
                        'id' => 5,
                        'username' => '小明',
                        'sex' => 0
                    ],
                    [
                        'id' => 8,
                        'username' => '王五',
                        'sex' => 2
                    ]
                ]
            ],
            [
                'sheet_name' => '导出数据表2',
                'data' => [
                    [
                        'id' => '110',
                        'date' => '2020-04-23 19:19:20'
                    ],
                    [
                        'id' => '119',
                        'date' => '2020-04-22 10:38:05'
                    ],
                    [
                        'id' => '120',
                        'date' => '2020-05-22 22:58:25'
                    ]
                ]
            ]
        ];
    }
}