import ExcelJS from 'exceljs';
import { buildContainerCSV, buildPackingWorkbook } from './excelExport';

const smallValueContainer = {
  id: 1,
  items: [{
    markNo: '1-1', description: 'Small values', pcsPerCtn: 1,
    pricePerPcs: 0.0001, gwPerCtn: 0.0001, nwPerCtn: 0.0001,
    cbmPerCtn: 0.0001, photo: ''
  }],
  totalCartons: 1,
  totalQty: 1,
  totalPrice: 0.0001,
  totalGW: 0.0001,
  totalNW: 0.0001,
  totalCBM: 0.0001
};

test('preserves four-decimal values in CSV exports', () => {
  const csv = buildContainerCSV(smallValueContainer);
  expect(csv).toContain('0.0001,0.0001,0.0001,0.0001,0.0001');
});

test('preserves and formats four-decimal values in workbook exports', async () => {
  const buffer = await buildPackingWorkbook([smallValueContainer], { maxWeight: 1, maxCbm: 1 }, new Date(0));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const row = workbook.worksheets[0].getRow(6);

  expect(row.getCell(9).value).toBe(0.0001);
  expect(row.getCell(13).value).toBe(0.0001);
  expect(row.getCell(9).numFmt).toBe('0.0000');
  expect(row.getCell(13).numFmt).toBe('0.0000');
});
