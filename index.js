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
          //'*Kategori Varian', 'Urutan Varian', '*Varian', 'Kode Varian'
          const dataVarian = sheetJSON.data;
          const headerVarian = dataVarian[0]
          const bodyVarian = dataVarian.slice(1)
          const indexKategoriVarian = headerVarian.findIndex(e => e === '*Kategori Varian')
          const indexUrutanVarian = headerVarian.findIndex(e => e === 'Urutan Varian')
          const indexVarian = headerVarian.findIndex(e => e === '*Varian')
          const indexKodeVarian = headerVarian.findIndex(e => e === 'Kode Varian')
          bodyVarian.forEach((e) => {
            json.varian.push({
              kategori_varian: e[indexKategoriVarian],
              urutan: e[indexUrutanVarian],
              nama_varian: e[indexVarian].split(','),
              kode_varian: e[indexKodeVarian].split(','),
            })
          })

        }
        if (sheetJSON.sheetName == 'Data Barang') {
          //*Mode Pengadaan	*Satuan Terkecil	Satuan Lain	*Barang dijual	*Tipe Barang	*Memiliki Varian	Varian SKU	Harga Jual Satuan	Harga Pokok/Standart	Harga Jual Minimum	Harga Beli Maksimum	Barcode	Kode SKU
          const dataBarang = sheetJSON.data;
          const headerBarang = dataBarang[0]
          const bodyBarang = dataBarang.slice(1)
          const indexNamaBarang = headerBarang.findIndex(e => e === '*Nama Barang')
          const indexJenisBarang = headerBarang.findIndex(e => e === '*Jenis Barang')
          const indexModePengadaan = headerBarang.findIndex(e => e === '*Mode Pengadaan')
          const indexSatuanTerkecil = headerBarang.findIndex(e => e === '*Satuan Terkecil')
          const indexSatuanLain = headerBarang.findIndex(e => e === 'Satuan Lain')
          const indexBarangDijual = headerBarang.findIndex(e => e === '*Barang dijual')
          const indexTipeBarang = headerBarang.findIndex(e => e === '*Tipe Barang')
          const indexMemilikiVarian = headerBarang.findIndex(e => e === '*Memiliki Varian')
          const indexVarianSKU = headerBarang.findIndex(e => e === 'Varian SKU')
          const indexHargaJualSatuan = headerBarang.findIndex(e => e === 'Harga Jual Satuan')
          const indexHargaPokok = headerBarang.findIndex(e => e === 'Harga Pokok/Standart')
          const indexHargaJualMinimum = headerBarang.findIndex(e => e === 'Harga Jual Minimum')
          const indexBarcode = headerBarang.findIndex(e => e === 'Barcode')
          const indexKodeSKU = headerBarang.findIndex(e => e === 'Kode SKU')
          bodyBarang.forEach((e) => {
            json.barang.push({
              nama_barang: e[indexNamaBarang],
            })
          })
        }
      }
    });
    console.log(json)
  })
  .catch((error) => {
    console.error('Error reading Excel file:', error);
  });




