import XLSX from 'xlsx';
import fs from 'fs';

const workbook = XLSX.readFile('/home/okachkac/container-optimizer/193所有产品清单.xlsx');

console.log('Sheet names:', workbook.SheetNames);

const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Get range
const range = XLSX.utils.decode_range(sheet['!ref']);
console.log('Range:', sheet['!ref'], '=> rows:', range.e.r + 1, 'cols:', range.e.c + 1);

// Print first 5 rows
console.log('\n--- First 5 rows ---');
const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
data.slice(0, 5).forEach((row, i) => {
  console.log('Row', i, ':', JSON.stringify(row));
});

// Print a few more rows
console.log('\n--- Rows 5-15 ---');
data.slice(5, 16).forEach((row, i) => {
  console.log('Row', i + 5, ':', JSON.stringify(row));
});

// Check for images
console.log('\n--- Images ---');
if (sheet['!images']) {
  console.log('Images found:', sheet['!images'].length);
} else {
  console.log('No !images property');
}

// Check workbook for media
if (workbook['!media']) {
  console.log('Media found:', workbook['!media'].length);
} else {
  console.log('No !media property');
}

// Total rows
console.log('\nTotal rows:', data.length);

// Print last product row
console.log('\n--- Last rows ---');
data.slice(-5).forEach((row, i) => {
  console.log('Row', data.length - 5 + i, ':', JSON.stringify(row));
});
