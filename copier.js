const ExcelJS = require('exceljs');
const fs = require('fs');

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
  }
}

async function createCopies() {
  const originalStream = fs.createReadStream('large_excel_file3.xlsx');
  const originalWorkbook = new ExcelJS.Workbook();
  await originalWorkbook.xlsx.read(originalStream);

  const targetWorkbook = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: fs.createWriteStream('copies_large_excel_file.xlsx'),
  });

  const numCopies = 1;

  for (let copyIndex = 0; copyIndex < numCopies; copyIndex++) {
    await copyWorksheet(originalWorkbook, targetWorkbook, 1);
  }

  await targetWorkbook.commit();

  console.log(`Created ${numCopies} copies. Excel file saved as "copies_large_excel_file.xlsx"`);
}

// Call the function to create copies
createCopies();
