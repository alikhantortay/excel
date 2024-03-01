const path = require('path');

function normalizePhoneNumbers(phoneNumber) {
    let phoneStr = String(phoneNumber);

    if (phoneStr === "") {
        return "";
    }

    let digits = phoneStr.replace(/\D/g, "");

    if (digits.startsWith("8")) {
        digits = "7" + digits.slice(1);
    }

    if (!digits.startsWith("7")) {
        digits = "7" + digits;
    }

    return `+${digits.slice(0, 1)} (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7, 9)}-${digits.slice(9)}`;
}


const ExcelJS = require('exceljs');

async function processExcelFile(filename) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filename);
    const worksheet = workbook.getWorksheet(1);

    let data = [];
    const phoneNumbers = new Set();
    const duplicates = new Set();

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (rowNumber === 1) return;

        const name = row.getCell(1).value;
        const email = row.getCell(2).value;
        let phone = row.getCell(3).value;

        if (phone) {
            phone = normalizePhoneNumbers(phone); // Применение функции нормализации
            if (phoneNumbers.has(phone)) {
                duplicates.add(phone);
            } else {
                phoneNumbers.add(phone);
                data.push({ id: data.length + 1, name, email, phone });
            }
        }
    });

    data = data.filter(item => !duplicates.has(item.phone));

    const outputPath = path.join(__dirname, 'newFiles', 'Обработанные номера.xlsx');

    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Processed Data');

    newWorksheet.addRow(['ID', 'Имя', 'Email', 'Телефон']);

    data.forEach(item => {
        newWorksheet.addRow([item.id, item.name, item.email, item.phone]);
    });

    await newWorkbook.xlsx.writeFile(outputPath);
}

processExcelFile('file/excel.xlsx').catch(err => console.error(err));
