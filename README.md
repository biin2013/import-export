## Import 使用说明

#### Import 构造函数参数

参数 | 类型 | 默认值 | 说明
---- | ----- | ----- | ----
file | string | - | 导入的文件路径
startRow | int | 1 | 开始读取数据的行数，从1开始
endRow | int | null | 结束读取数据的行数
startColumn | string | 'A' | 开始读取数据的列
endColumn | string | null | 结束读取数据的列
columnFields | array | [] | 列对应的字段名，在返回数据时自动替换列名

#### Import 方法

* setRows

  如果希望只读取指定行（非连续）的数据，在构造函数中指定开始和结束行就无法满足了， 此时可使用此方法自定义一个只需读取的行数组即可，且使用该方法指定后，构造函数中指定 的 `startRow` 与 `endRow` 失效。


* setColumns

  同 `setRows`，此处指定的为列。


* getReader

  使用该方法可得到 `PhpOffice\PhpSpreadsheet\Reader\IReader` 对象。


* getSpreadsheet

  使用该方法可得到 `PhpOffice\PhpSpreadsheet\Spreadsheet` 对象


* setOnlyReadSheetName

  使用该方法可指定只读取的 `sheet` 名称，可传入字符串或数组


* setOnlyReadSheetIndex

  使用该方法可指定只读取的 `sheet` 的index

> 注： 如 `setOnlyReadSheetName` 与 `setOnlyReadSheetIndex` 同时指定，则取并集

* getData

  使用该方法获取读取到的数据，可指定一个匿名回调函数来格式化每个单元格的数据。 回调函数的第一个参数为字段名（如未指定则使用列名），第二个参数为当前单元格的值，
  第三个参数为当前单元格对象，第四个参数为当前 `PhpOffice\PhpSpreadsheet\Worksheet\Worksheet` 对象
