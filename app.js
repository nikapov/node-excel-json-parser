const excel = require('node-excel-export');
const readXlsxFile = require('read-excel-file/node');
const XLSX = require('xlsx');
const fs = require('fs');

const consts = require('./consts');

json_to_excel = function () {
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

    console.log('Xlsx exported to ./output/report.xlsx');

}

excel_to_json = function () {
    var schema = consts.schema;

    readXlsxFile('./input/report.xlsx',  { getSheets: true }).then((sheets) => {
        sheets.forEach(sheet => {
            readXlsxFile('./input/report.xlsx',  { sheet: sheet.name }).then((rows) => {
                fs.writeFileSync("./output/report_"+sheet.name+'.json', JSON.stringify(rows, null, 2));
            });
        });
        
    });

    console.log('Json exported to ./output/report.json');
}

excel_to_workbook_sheets = function () {
    var workbook = XLSX.readFile('./input/report.xlsx');
    // XLSX.utils.sheet_to_json(workbook);
    XLSX.writeFile(workbook, './output/report.xlsb');
}


module.exports = { json_to_excel, excel_to_json, excel_to_workbook_sheets };
