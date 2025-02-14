1. 访问工作簿和工作表
context.workbook: 获取当前工作簿对象。
context.workbook.worksheets: 获取当前工作簿中的工作表集合。
context.workbook.worksheets.getActiveWorksheet(): 获取当前活动工作表。
context.workbook.worksheets.getItem(name): 根据名称获取特定工作表。
2. 读取和写入单元格
worksheet.getRange("A1"): 获取指定单元格或范围的对象。
range.load("values"): 加载单元格的值。
range.values = [[1, 2, 3]]: 修改单元格的值。
range.format.fill.color = "yellow": 设置单元格背景颜色。
3. 提交更改
context.sync(): 提交对 Excel 对象模型的更改。context.sync() 是将所有的操作提交到 Excel 的重要步骤，只有调用 sync，更改才会生效。
4. 处理表格数据
worksheet.tables: 获取工作表中的所有表格。
table.getRange(): 获取表格的范围。
table.rows.add(null, [[1, 2, 3]]): 向表格中添加行。
table.getDataBodyRange().load("values"): 加载表格的数据范围。

context.application: 获取 Excel 应用程序对象，进行如打开文件等操作。
context.document: 获取当前文档对象。
context.sync(): 提交更改并同步所有更改。