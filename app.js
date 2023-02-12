const path = require('path')
const express = require('express')
const excel = require('exceljs')
var cons = require('consolidate');
const localStorage = require('local-storage');
const { json } = require('express');

const app = express()
const port = process.env.PORT || 3000

const basePath = path.join(__dirname, './public')

app.use(json())
app.use(express.static(basePath))
app.engine('html', cons.swig)
app.set('views', path.join(__dirname, 'public'));
app.set('view engine', 'html');

app.post("/excel", (req, res, next)=>{


    // Tạo một đối tượng Workbook
    const workbook = new excel.Workbook();

    // Tạo một worksheet mới
    const worksheet = workbook.addWorksheet('Sheet 1');

    // Thêm dữ liệu vào worksheet
    // const dataLocalStorage = localStorage.get('list_attendance')
    // console.log(dataLocalStorage);
    const dt = req.body || "[]"
    console.log("body", dt)
    const arrData = dt//JSON.parse(dt)
    worksheet.addRow(['name', 'time'])
    for (let e of arrData){
        worksheet.addRow([e.name, e.time])
    }

    // Lưu tập tin
    workbook.xlsx.writeFile('data.xlsx').then(function() {
        res.send("ok")
    });
})
app.get('/download', (req, res)=>{
    console.log("haha redic");
    const file = 'data.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=' + file);
    res.sendFile(file, { root: __dirname });
})

//start express server
app.listen(port, () => {
    console.log('Server started on post ' + port)
})