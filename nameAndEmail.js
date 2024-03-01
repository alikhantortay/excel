const path = require('path');

const ExcelJS = require('exceljs');

async function processNamesAndEmails(filename, outputFilename) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet(1);

    let data = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1) return;

        const name = row.getCell(1).value;
        const email = row.getCell(2).value;

        if (name && email) {
            data.push({ name, email });
        }
    });

    const outputPath = path.join(__dirname, 'newFiles', 'outputNamesAndEmails.xlsx');

    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Names and Emails');

    newWorksheet.addRow(['Имя пользователя', 'Почта']);

    data.forEach(item => {
        newWorksheet.addRow([item.name, item.email]);
    });

    await newWorkbook.xlsx.writeFile(outputPath);

}

processNamesAndEmails('file/excel.xlsx', 'outputNamesAndEmails.xlsx').catch(err => console.error(err));
