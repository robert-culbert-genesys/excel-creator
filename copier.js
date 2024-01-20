const ExcelJS = require('exceljs');

async function copyWorksheet(originalWorkbook, targetWorkbook, numCopies) {
  // Ensure the original workbook has at least one worksheet
  if (originalWorkbook.worksheets.length < 1) {
    throw new Error('Original workbook does not contain any worksheets.');
  }

  const originalWorksheet = originalWorkbook.worksheets[0]; // Access the first worksheet

  for (let copyIndex = 0; copyIndex < numCopies; copyIndex++) {
    const copiedWorksheet = targetWorkbook.addWorksheet(`Copy${copyIndex + 1}`);

    originalWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const newRow = copiedWorksheet.getRow(rowNumber);
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        newRow.getCell(colNumber).value = cell.value;
        newRow.getCell(colNumber).style = cell.style;
      });
    });

    // Commit the copied worksheet to the target workbook to avoid excessive memory usage
    await copiedWorksheet.commit();
  }
}

async function createCopies() {

  let currentTime = new Date().toLocaleTimeString();

  console.log(`Starting at Time: ${currentTime}`);

  const originalWorkbook = new ExcelJS.Workbook();
  await originalWorkbook.xlsx.readFile('large_excel_file3.xlsx');

  const targetWorkbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: require('fs').createWriteStream('large_excel_file_with_copies50_10_6.xlsx'),
  });

  const numCopies = 50;

  for (let copyIndex = 0; copyIndex < numCopies; copyIndex++) {
    await copyWorksheet(originalWorkbook, targetWorkbook, 1);
  }

  // Commit the target workbook to save it
  await targetWorkbook.commit();

  currentTime = new Date().toLocaleTimeString();

  console.log(`Created ${numCopies} copies. Excel file saved as "large_excel_file_with_copies100_10_6.xlsx". Time: ${currentTime}`);
}

// Call the function to create copies
createCopies();
