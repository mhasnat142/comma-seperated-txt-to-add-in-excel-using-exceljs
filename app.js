var express = require('express');
var path = require('path');
var app = express();
const fs = require('fs')
var exceljs = require('exceljs');
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');

app.use(express.json());
var ArrayOfData = [];
readFile()
async function readFile() {
  console.log('readinbgFile');
  var data =  await  fs.readFile(__dirname+'/data.txt', 'utf8' , (err, data) => {
    if (err) {
      console.error(err)
      return
    }
    else{
      const arr = data.toString().replace(/\r\n/g,'\n').split('\n');
      for(let i in arr) {
          var element = arr[i];
          var child = element.split(',');
          var newArr = []
          for (let i = 0; i < child.length; i++) {
            const childElement = child[i];
            newArr.push(childElement)
          }
          ArrayOfData.push(newArr)
      }
      // return ArrayOfData
      readExcel()
    }
  })
}
async function readExcel() {
  // var response = await readFile()
  let nameFileExcel = 'excelFileToUpload.xlsx'
  var workbook = new exceljs.Workbook();
  workbook.xlsx.readFile(nameFileExcel)
  .then(function()  {
    console.log('loop Started')
    console.log(ArrayOfData.length);
    for (let i = 380000; i < ArrayOfData.length; i++) {
      const element = ArrayOfData[i];
      var worksheet = workbook.getWorksheet(1);
      worksheet.addRow(element);
    }
    console.log('loop Ended')
    console.log('creating file');
    workbook.xlsx.writeFile(nameFileExcel);
    console.log('done');
  });
}



module.exports = app;