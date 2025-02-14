function readExcel() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("A1:B10");
        range.load("values");

        return context.sync().then(function () {
            console.log("Data from A1:B10:", range.values);
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function writeExcel(cell, content) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(cell);
        range.values = [[content]];

        return context.sync().then(function () {
            // console.log("Data written to A1");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function setCellStyle() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("A1:B1");

        range.format.font.bold = true;
        range.format.fill.color = "yellow";
        range.format.border.color = "black";

        return context.sync().then(function () {
            console.log("Cell style applied to A1:B1");
        });
    }).catch(function (error) {
        console.error(error);
    });
}

function addFormula() {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange("C2");
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

// 在 Excel 加载项中创建任务窗格
// Office.context.ui.displayDialogAsync('https://example.com/taskpane.html', { height: 50, width: 50 }, function (asyncResult) {
//     if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//         var dialog = asyncResult.value;
//         dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (message) {
//             console.log(message);
//         });
//     } else {
//         console.error('Error opening dialog:', asyncResult.error);
//     }
// });


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


export {
    readExcel, writeExcel, setCellStyle, addFormula
}