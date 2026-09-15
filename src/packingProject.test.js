import ExcelJS from 'exceljs';
import { importPackingWorkbook, validateProject } from './packingProject';

async function workbook(ctn = 3) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Container 1');
  ws.addRow(['CONTAINER #1 - PACKING LIST']);
  ws.addRow(['No', 'Photo', 'MARKS&NO', 'DESCRIPTION', 'CTN', 'PCS/CTN', 'UNIT', 'T/QTY', 'U/PRICE', 'AMOUNT', 'G.W.(KGS)', 'N.W.(KGS)', 'CBM']);
  ws.addRow([1, '', '10-1', 'Edited product', ctn, 12, 'PCS', 999, 2.5, 999, 60, 54, 0.9]);
  ws.addRow(['', '', 'TOTAL']);
  return wb.xlsx.writeBuffer();
}

test('reopens edited packing lists and recalculates stale Excel quantity and price totals', async () => {
  const project = await importPackingWorkbook(await workbook());
  expect(project.containers[0]).toMatchObject({ totalCartons: 3, totalQty: 36, totalPrice: 90, totalGW: 60, totalNW: 54 });
  expect(project.containers[0].totalCBM).toBeCloseTo(0.9);
  expect(project.containers[0].items[0].description).toBe('Edited product');
  expect(validateProject(project).containers[0].totalPrice).toBe(90);
});

test.each([-1, 1.5, 50001])('rejects invalid or excessive carton quantities: %s', async quantity => {
  await expect(importPackingWorkbook(await workbook(quantity))).rejects.toThrow();
});

test('rejects unrelated workbooks and unsupported backups', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Catalog').addRow(['Product', 'Quantity']);
  await expect(importPackingWorkbook(await wb.xlsx.writeBuffer())).rejects.toThrow('No packing-list');
  expect(() => validateProject({ version: 2 })).toThrow();
});

test('preserves embedded PNG photos', async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await workbook());
  const base64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aD1sAAAAASUVORK5CYII=';
  const imageId = wb.addImage({ base64, extension: 'png' });
  wb.worksheets[0].addImage(imageId, { tl: { col: 1.1, row: 2.1 }, br: { col: 1.9, row: 2.9 } });
  const project = await importPackingWorkbook(await wb.xlsx.writeBuffer());
  expect(project.catalog[0].photo).toBe(`data:image/png;base64,${base64}`);
});
