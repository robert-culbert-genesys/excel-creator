const XLSX = require('xlsx');
const fs = require('fs');

function generateRandomString(length) {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * characters.length));
  }
  return result;
}

function generateRandomTimestamp() {
  const randomDate = new Date(+(new Date()) - Math.floor(Math.random() * 10000000000));
  return randomDate.toISOString();
}

function generateRandomFloat(min, max) {
  return (Math.random() * (max - min) + min).toFixed(2);
}

function generateExcel() {
  const workbook = XLSX.utils.book_new();

  // Define parameters
  const numWorksheets = 2;
  const rowsPerWorksheet = 80000;
  const colsPerRow = 10;
  const batchSize = 10;

  for (let sheetIndex = 0; sheetIndex < numWorksheets; sheetIndex++) {
    const worksheet = XLSX.utils.aoa_to_sheet([]);

    for (let batchIndex = 0; batchIndex < rowsPerWorksheet; batchIndex += batchSize) {
      const batch = [];

      for (let rowIndex = batchIndex; rowIndex < batchIndex + batchSize && rowIndex < rowsPerWorksheet; rowIndex++) {
        const row = [];
        for (let colIndex = 0; colIndex < colsPerRow; colIndex++) {
          const randomType = Math.floor(Math.random() * 3);

          switch (randomType) {
            case 0:
              row.push(generateRandomString(100));
              break;
            case 1:
              row.push(generateRandomTimestamp());
              break;
            case 2:
              row.push(generateRandomFloat(0, 100));
              break;
          }
        }

        batch.push(row);
      }

      XLSX.utils.sheet_add_aoa(worksheet, batch, { origin: -1 });
    }

    XLSX.utils.book_append_sheet(workbook, worksheet, `Sheet${sheetIndex + 1}`);
  }

  const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

  fs.writeFileSync('large_excel_file.xlsx', Buffer.from(excelBuffer));
  console.log('Excel file saved as "large_excel_file.xlsx"');
}

// Call the function to generate and save the Excel file
generateExcel();