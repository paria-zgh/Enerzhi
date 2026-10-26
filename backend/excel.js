const ExcelJS = require('exceljs');

async function createExcel(headers, rows, filePath) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Offers');

  if (headers.length > 0) {
    worksheet.addRow(headers);
  }

  for (const row of rows) {
    worksheet.addRow(row);
  }

  worksheet.getRow(1).font = { bold: true };

  await workbook.xlsx.writeFile(filePath);
}

module.exports = { createExcel };
