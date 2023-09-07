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
      console.log(sheetRows)

    });
    const json = { satuan: [], jenis_barang: [], barang: [], varian: [] }
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

        if (sheetJSON.sheetName == 'Jenis Barang') {
          const dataJenisBarang = sheetJSON.data;
          const headerJenisBarang = dataJenisBarang[0]
          const bodyJenisBarang = dataJenisBarang.slice(1)
          const indexNamaJenisBarang = headerJenisBarang.findIndex(e => e === '*Nama Jenis Barang')
          const indexKodeJenisBarang = headerJenisBarang.findIndex(e => e === '*Kode Jenis Barang')
          const indexTampilkanJenisBarang = headerJenisBarang.findIndex(e => e === '*Tampilkan di POS (Ya/Tidak)')

          bodyJenisBarang.forEach((e) => {
            json.jenis_barang.push({
              nama: e[indexNamaJenisBarang],
              kode: e[indexKodeJenisBarang],
              is_pos: e[indexTampilkanJenisBarang] == 'Ya' ? true : false
            })
          })

        }
        if (sheetJSON.sheetName == 'Varian & SKU') {

        }
        if (sheetJSON.sheetName == 'Data Barang') {

        }
      }
    });
    console.log(json)
  })
  .catch((error) => {
    console.error('Error reading Excel file:', error);
  });




