const ExcelJS = require('exceljs');
const fs = require('fs');

// Load the Excel file
const workbook = new ExcelJS.Workbook();
const excelFilePath = './upload.xlsx'; // Replace with the path to your Excel file

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

    sheetData.forEach((sheetJSON, index) => {
      if (index > 0) {
        console.log(`Saved ${sheetJSON.sheetName}`);
        console.log(`Data`, sheetJSON.data)
      }
    });
  })
  .catch((error) => {
    console.error('Error reading Excel file:', error);
  });
