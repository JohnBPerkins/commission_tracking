import * as excel from 'exceljs'

function main() {
    var workbook = new excel.Workbook();
    workbook.xlsx.readFile('./temp/test.xlsx');


}