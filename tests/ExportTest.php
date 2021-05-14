<?php

use ImportExport\Export;

class ExportTest extends \PHPUnit\Framework\TestCase
{
    public function testExport()
    {
        $export = new Export($this->getData());
        $export->setTitle([
            [
                'merge_row' => 2,
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
                                'format' => function ($value) {
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
        ]);
        print_r($export->save('./export/test_export'));
    }

    private function getData(): array
    {
        return [
            [
                'sheet_name' => '导出数据表1',
                'data' => [
                    [
                        'id' => 1,
                        'username' => '张三',
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