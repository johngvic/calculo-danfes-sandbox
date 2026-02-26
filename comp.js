const xlsx = require('xlsx');
const fs = require('fs');

// open huge.xlsx
const workbook = xlsx.readFile('result.xlsx');

// get the first sheet
const sheetName = workbook.SheetNames[2];
const sheet = workbook.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

const diff = [];

for (let i = 1; i < 797367; i++) {
  const [_, __, original, calculated] = rows[i];
  const originalNum = Number(original.toFixed(2));

  if (originalNum != calculated) {
    diff.push({
      row: i + 1,
      original: originalNum,
      calculated
    });
  }
}

// save diff to diff.json
fs.writeFileSync('diff.json', JSON.stringify(diff, null, 2));