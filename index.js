const ExcelJS = require('exceljs');
const fs = require('fs');

// Load the Excel file
const workbook = new ExcelJS.Workbook();
const excelFilePath = './upload.xlsx'; // Replace with the path to your Excel file

let json = []
workbook.xlsx.readFile(excelFilePath)
  .then(() => {
    const sheetData = [];

    workbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name;
      const sheetRows = [];

      worksheet.eachRow((row, rowNumber) => {
        const rowData = [];

        row.eachCell((cell, colNumber) => {
          rowData.push(cell.value);
        });

        sheetRows.push(rowData);
      });

      const sheetJSON = {
        sheetName,
        data: sheetRows,
      };

      sheetData.push(sheetJSON);

    });
    const json = { satuan: [], barang: [], varian: [] }
    sheetData.forEach((sheetJSON, index) => {
      if (index > 0) {
        // console.log(`Saved ${sheetJSON.sheetName}`);
        const satuan = []
        if (sheetJSON.sheetName == 'Satuan') {
          const dataSatuan = sheetJSON.data;
          const headerSatuan = dataSatuan[0]
          const bodySatuan = dataSatuan.slice(1)
          const indexNamaSatuan = headerSatuan.findIndex(e => e === '*Nama Satuan')

          bodySatuan.forEach((e) => {
            json.satuan.push({
              nama_satuan: e[indexNamaSatuan]
            })
          })
        }
      }
    });
  })
  .catch((error) => {
    console.error('Error reading Excel file:', error);
  });




