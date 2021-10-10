// yêu cầu thư viện
const Excel =require("exceljs")
const path =require('path')

// tạo một workbook trống để bắt đầu làm việc với file excel
var wb= new Excel.Workbook();

// lấy đường dẫn file excel
var filepath=path.resolve(__dirname,'ex.xlsx');

// đọc file excel
wb.xlsx.readFile(filepath).then(function(){
    // lấy 1 sheet(trang tính) có tên là "Sheet1" trong workbook(wb) được mở
    var sh= wb.getWorksheet("Sheet1");

    // đếm số hàng
    console.log(sh.rowCount);

    //lấy giá trị ở ô vị trí 1 ,1
    console.log(sh.getCell(1,1).value);
});