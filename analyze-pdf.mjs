import fs from 'fs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';

const pdfPath = '/home/okachkac/container-optimizer/193所有产品清单.pdf';
const data = new Uint8Array(fs.readFileSync(pdfPath));
const pdf = await pdfjs.getDocument({ data }).promise;

// Look for product 96 specifically
const page = await pdf.getPage(8); // Product 96 should be on page 8
const text = await page.getTextContent();

console.log('All items on page 8 with positions:');
text.items.forEach((item, i) => {
  if (item.str.includes('96') || item.str.includes('0.028') || item.str.includes('16.6') || item.str.includes('400')) {
    console.log(i, 'x=' + Math.round(item.transform[4]), 'y=' + Math.round(item.transform[5]), JSON.stringify(item.str));
  }
});

console.log('\n--- All items around y=193-197 (where 96 might be) ---');
text.items.filter(item => {
  const y = Math.round(item.transform[5]);
  return y >= 190 && y <= 200;
}).forEach(item => {
  console.log('x=' + Math.round(item.transform[4]), 'y=' + Math.round(item.transform[5]), JSON.stringify(item.str));
});

