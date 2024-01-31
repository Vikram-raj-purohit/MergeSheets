const ExcelJS = require('exceljs');

async function mergeSheets(filePath) {
  // Load the workbook
  const workbook = new ExcelJS.Workbook();
  
  await workbook.xlsx.readFile(filePath);

  // Create a new workbook for the merged sheets
  const mergedWorkbook = new ExcelJS.Workbook();
  const mergedSheet = mergedWorkbook.addWorksheet('Merged Sheet');

  // Loop through each sheet in the original workbook
  workbook.eachSheet((worksheet, sheetId) => {
    // Copy the sheet data to the merged sheet
    worksheet.eachRow((row, rowNumber) => {
      const rowData = row.values;
      mergedSheet.addRow(rowData);
    });
  });

  // Save the merged workbook to a new file
  const outputFilePath = 'merged.xlsx';
  await mergedWorkbook.xlsx.writeFile(outputFilePath);
  console.log(`Merged sheets saved to ${outputFilePath}`);
}

const excelFilePath = 'Book2.xlsx';
mergeSheets(excelFilePath);
