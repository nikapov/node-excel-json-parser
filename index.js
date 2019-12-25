const excel = require('node-excel-export');
const readXlsxFile = require('read-excel-file/node');
const XLSX = require('xlsx');
const fs = require('fs');

const consts = require('./consts');

var App = function () {

    //  Scope.
    var self = this;

    self.initialize = function () {
        console.log("Initialized!");
    }

    self.start = function () {
        console.log("Started!");
    }

    self.json_to_excel = function () {
        // Create the excel report.
        // This function will return Buffer
        const report = excel.buildExport(
            [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
                {
                    name: 'Report', // <- Specify sheet name (optional)
                    heading: consts.heading, // <- Raw heading array (optional)
                    merges: consts.merges, // <- Merge cell ranges
                    specification: consts.specification, // <- Report specification
                    data: JSON.parse(fs.readFileSync('./input/report.json', 'utf8')).data //consts.dataset // <-- Report data
                }
            ]
        );

        fs.writeFileSync("./output/report.xlsx", report);//, function (err) {
        //     if (err) {
        //         return console.log(err);
        //     }
        //     console.log("The file was saved!");
        // });

    }

    self.excel_to_json = function () {
        var schema = consts.schema;

        readXlsxFile('./input/report.xlsx', { schema }).then((rows) => {
            fs.writeFileSync("./output/report.json", JSON.stringify(rows, null, 2));
        });
    }

    self.excel_to_json_sheets = function () {
        var workbook = XLSX.readFile('./input/report.xlsx');
        // XLSX.utils.sheet_to_json(workbook);
        XLSX.writeFile(workbook, './output/report.xlsb');
    }

}

/**
 *  main():  Main code.
 */
var app = new App();
app.initialize();
app.start();

app.json_to_excel();
// app.excel_to_json();
