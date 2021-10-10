var Excel = require("exceljs");
var path = require("path");


var wb = new Excel.Workbook();
var filePath = path.resolve(__dirname, "ex.xlsx");

wb.xlsx.readFile(filePath).then(function () {
  var sh = wb.getWorksheet("Sheet1");

  sh.getRow(1).getCell(2).value = 32;
  wb.xlsx.writeFile("sample2.xlsx");
  console.log("Row-3 | Cell-2 - " + sh.getRow(3).getCell(2).value);

  console.log(sh.rowCount);
  //Get all the rows data [1st and 2nd column]
  for (i = 1; i <= sh.rowCount; i++) {
    console.log(
      sh.getRow(i).getCell(1).value + "\t" + sh.getRow(i).getCell(2).value
    ); 

  }
});
