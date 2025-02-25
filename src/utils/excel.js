function readExcel(cells) {
    return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(cells);
        range.load("values");
        return context.sync().then(function () {
            return range.values
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function readUseExcel() {
    return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var usedRange = sheet.getUsedRange();
        usedRange.load("values");
        return context.sync().then(function () {
            return usedRange.values
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function getUseAddress() {
    return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var selectedRange = sheet.getRange();
        selectedRange.load('address');  // 加载选中区域的地址
        return context.sync().then(function () {
            const address = selectedRange.address
            writeExcel("D1", address)
            return address
        });
    });
}

// Excel.run(function (context) {
//     var sheet = context.workbook.worksheets.getActiveWorksheet();
//     var range = null
//     range = sheet.getRange("A1");
//     range.values = [["content"]];

//     context.sync().then(function () {
//     });
// }).catch(function (error) {
// });


function writeExcel(cell, content) {
    return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = null
        range = sheet.getRange(cell);
        range.values = [[content]];

        return context.sync().then(function () {
            return 1
        });
    }).catch(function (error) {
        return 0
    });
}

const writeSelectedRange = (content) => {
    return Excel.run(function (context) {
        var selectedRange = context.workbook.getSelectedRange();
        selectedRange.values = [[content]]

        return context.sync().then(function () {
            return 1
        });
    }).catch(function (error) {
        return 0
    });
}

function writeNonExcel(column, content) {
    Excel.run(function (context) {
        let columnInd = column.charCodeAt(0) - 65
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var usedRange = sheet.getUsedRange();  // 获取工作表的已使用范围
        usedRange.load("values,rowCount");  // 加载范围的值和行数
        return context.sync().then(function () {
            var lastRow = usedRange.rowCount;
            for (let i = lastRow - 1; i >= 0; i--) {
                const val = usedRange.values[i][columnInd]
                if (val !== null && val !== "") {
                    lastRow = i + 2;  // 行索引从1开始
                    break;
                }
            }
            // 将内容写入到最后一个有内容行的下一行
            var targetCell = column + lastRow;
            var targetRange = sheet.getRange(targetCell);
            targetRange.values = [[content]];

            return context.sync();
        });
    }).catch(function (error) {
        // 弹窗报错提醒
    });
}

function setCellStyle(cells) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(cells);

        range.format.font.bold = true;
        range.format.fill.color = "yellow";
        range.format.border.color = "black";

        return context.sync().then(function () {
            console.log(`Cell style applied to ${cells}`);
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function addFormula(cell) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(cell);
        range.formulas = [["=A2*B2"]];

        return context.sync().then(function () {
            console.log("Formula added to C2");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function createTable() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("A1:C4");
        var table = sheet.tables.add("MyTable", range);
        table.name = "Table1";
        table.getHeaderRowRange().values = [["Name", "Age", "Country"]];
        table.rows.add(null, [["John", 30, "USA"], ["Alice", 25, "UK"]]);

        return context.sync().then(function () {
            console.log("Table created!");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function createPivotTable() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("A1:D6");
        var pivotTable = sheet.pivotTables.add("PivotTable1", range, sheet.getRange("F1"));

        pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItemAt(0));
        pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItemAt(1));

        return context.sync().then(function () {
            console.log("Pivot Table created!");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function createChart() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("A1:B4");

        var chart = sheet.charts.add("ColumnClustered", range, "Auto");
        chart.title.text = "Sales Data";
        chart.legend.position = "Top";
        chart.legend.visible = true;

        return context.sync().then(function () {
            console.log("Chart created!");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function setWorkbookMetadata() {
    Excel.run(function (context) {
        var workbook = context.workbook;
        workbook.properties.title = "Sales Report";
        workbook.properties.author = "Your Name";
        workbook.properties.subject = "Monthly Sales Data";

        return context.sync().then(function () {
            console.log("Metadata set successfully!");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

const openMyDialog = () => {

}

// messageBox
const openMessageBox = (message, cell) => {
    Excel.run(context => {
        const app = context.application;
        writeExcel("B1", 1)
        app.messageBox("确认框", `您确定要将数据${message}插入到 ${cell} 单元格吗?`, ["是", "否"]);
        writeExcel("B1", 2)
        context.sync();
    }).catch(error => {
        console.error(error);
    });
}

export {
    readExcel, readUseExcel, writeExcel, writeSelectedRange, writeNonExcel, openMessageBox, openMyDialog, setCellStyle, addFormula
}