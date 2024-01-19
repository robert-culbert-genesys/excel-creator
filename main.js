const ExcelJS = require('exceljs');
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

async function generateAndSaveExcel() {
  const stream = fs.createWriteStream('large_excel_file.xlsx');

  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream,
    useStyles: true,
  });

  const numWorksheets = 2;
  const rowsPerWorksheet = 1000000;

  for (let sheetIndex = 0; sheetIndex < numWorksheets; sheetIndex++) {
    const worksheet = workbook.addWorksheet(`Sheet${sheetIndex + 1}`);

    for (let rowIndex = 0; rowIndex < rowsPerWorksheet; rowIndex++) {
      const row = [];

      for (let colIndex = 0; colIndex < 10; colIndex++) {
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

      worksheet.addRow(row);
    }

    // Commit the rows to the worksheet to avoid excessive memory usage
    await worksheet.commit();
  }

  // Commit the workbook to save it
  await workbook.commit();
  stream.end();

  console.log('Excel file saved as "large_excel_file.xlsx"');
}

// Call the function to generate and save the Excel file
generateAndSaveExcel();
