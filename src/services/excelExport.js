import ExcelJS from 'exceljs';

/**
 * Format a container as CSV, preserving group order and existing totals.
 * @param {Object} container Container with carton items and aggregate totals.
 * @returns {string} Packing-list CSV text.
 */
export const buildContainerCSV = (container) => {
  const groupedItems = {};
  const groupOrder = [];
  container.items.forEach(item => {
    const key = item.markNo + '|' + item.description;
    if (!groupedItems[key]) {
      groupedItems[key] = {
        count: 0, markNo: item.markNo, description: item.description,
        pcsPerCtn: item.pcsPerCtn, gwPerCtn: item.gwPerCtn,
        nwPerCtn: item.nwPerCtn, cbmPerCtn: item.cbmPerCtn,
        pricePerPcs: item.pricePerPcs || 0
      };
      groupOrder.push(key);
    }
    groupedItems[key].count++;
  });

  const headers = ['No', 'MARKS&NO', 'DESCRIPTION', 'CTN', 'PCS/CTN', 'UNIT', 'T/QTY', 'U/PRICE', 'AMOUNT', 'G.W.(KGS)', 'N.W.(KGS)', 'CBM'];
  const rows = [headers.join(',')];
  let idx = 0;
  for (const key of groupOrder) {
    const data = groupedItems[key];
    idx++;
    const totalQty = data.count * data.pcsPerCtn;
    const amount = data.pricePerPcs * totalQty;
    rows.push([
      idx,
      `"${data.markNo}"`,
      `"${data.description.replace(/"/g, '""')}"`,
      data.count,
      data.pcsPerCtn,
      'PCS',
      totalQty,
      data.pricePerPcs.toFixed(2),
      amount.toFixed(2),
      (data.gwPerCtn * data.count).toFixed(2),
      (data.nwPerCtn * data.count).toFixed(2),
      (data.cbmPerCtn * data.count).toFixed(4)
    ].join(','));
  }
  // Total row
  rows.push([
    '', '', 'TOTAL', container.totalCartons, '', '', container.totalQty || '',
    '', (container.totalPrice || 0).toFixed(2),
    container.totalGW.toFixed(2), container.totalNW.toFixed(2), container.totalCBM.toFixed(2)
  ].join(','));

  return rows.join('\n');
};

/**
 * Populate a worksheet with the existing packing-list layout and images.
 * Mutates the supplied ExcelJS workbook and worksheet.
 * @param {ExcelJS.Workbook} wb Workbook that owns embedded images.
 * @param {ExcelJS.Worksheet} ws Worksheet to populate.
 * @param {Object} container Container with carton items and aggregate totals.
 * @param {{maxWeight: number, maxCbm: number}} limits Numeric capacity limits.
 * @returns {void}
 */
export const buildContainerSheet = (wb, ws, container, { maxWeight, maxCbm }) => {
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E5090' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  const titleFont = { bold: true, size: 14, color: { argb: 'FF2E5090' } };
  const summaryFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F4FF' } };
  const totalFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCEAFF' } };
  const thinBorder = {
    top: { style: 'thin' }, left: { style: 'thin' },
    bottom: { style: 'thin' }, right: { style: 'thin' }
  };

  // Column widths
  ws.columns = [
    { key: 'no', width: 6 },
    { key: 'photo', width: 14 },
    { key: 'markNo', width: 12 },
    { key: 'description', width: 40 },
    { key: 'ctn', width: 8 },
    { key: 'pcsPerCtn', width: 10 },
    { key: 'unit', width: 7 },
    { key: 'totalQty', width: 10 },
    { key: 'uPrice', width: 12 },
    { key: 'amount', width: 14 },
    { key: 'gw', width: 12 },
    { key: 'nw', width: 12 },
    { key: 'cbm', width: 10 }
  ];

  // ── Title row ──
  const titleRow = ws.addRow([`CONTAINER #${container.id} — PACKING LIST`]);
  ws.mergeCells(titleRow.number, 1, titleRow.number, 13);
  titleRow.getCell(1).font = titleFont;
  titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
  titleRow.height = 30;

  // ── Summary rows ──
  const weightPct = maxWeight > 0 ? (container.totalGW / maxWeight * 100).toFixed(1) : '0';
  const cbmPct = maxCbm > 0 ? (container.totalCBM / maxCbm * 100).toFixed(1) : '0';

  const summaryData = [
    ['Total Cartons:', container.totalCartons, '', 'G.W.:', `${container.totalGW.toFixed(2)} KGS`, `(${weightPct}% of ${maxWeight} kg)`, '', 'CBM:', `${container.totalCBM.toFixed(2)} M³`, `(${cbmPct}% of ${maxCbm} m³)`, '', 'Total Price:', `¥${(container.totalPrice || 0).toFixed(2)}`],
    ['N.W.:', `${container.totalNW.toFixed(2)} KGS`, '', 'T/QTY:', container.totalQty || '', '', '', '', '', '', '', '', '']
  ];
  summaryData.forEach(rowData => {
    const row = ws.addRow(rowData);
    row.eachCell((cell) => {
      cell.fill = summaryFill;
      cell.font = { bold: true, size: 10 };
    });
    row.height = 22;
  });

  ws.addRow([]); // spacer

  // ── Header row ──
  const headers = ['No', 'Photo', 'MARKS&NO', 'DESCRIPTION', 'CTN', 'PCS/CTN', 'UNIT', 'T/QTY', 'U/PRICE', 'AMOUNT', 'G.W.(KGS)', 'N.W.(KGS)', 'CBM'];
  const headerRow = ws.addRow(headers);
  headerRow.height = 24;
  headerRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = thinBorder;
  });

  // ── Group items ──
  const groupedItems = {};
  const groupOrder = [];
  container.items.forEach(item => {
    const key = item.markNo + '|' + item.description;
    if (!groupedItems[key]) {
      groupedItems[key] = {
        count: 0, markNo: item.markNo, description: item.description,
        pcsPerCtn: item.pcsPerCtn, gwPerCtn: item.gwPerCtn,
        nwPerCtn: item.nwPerCtn, cbmPerCtn: item.cbmPerCtn,
        pricePerPcs: item.pricePerPcs || 0, photo: item.photo
      };
      groupOrder.push(key);
    }
    groupedItems[key].count++;
  });

  // ── Data rows ──
  const ROW_HEIGHT = 55;
  let idx = 0;
  for (const key of groupOrder) {
    const data = groupedItems[key];
    idx++;
    const totalQty = data.count * data.pcsPerCtn;
    const amount = data.pricePerPcs * totalQty;
    const row = ws.addRow([
      idx,
      '', // photo placeholder
      data.markNo,
      data.description,
      data.count,
      data.pcsPerCtn,
      'PCS',
      totalQty,
      parseFloat(data.pricePerPcs.toFixed(2)),
      parseFloat(amount.toFixed(2)),
      parseFloat((data.gwPerCtn * data.count).toFixed(2)),
      parseFloat((data.nwPerCtn * data.count).toFixed(2)),
      parseFloat((data.cbmPerCtn * data.count).toFixed(4))
    ]);
    row.height = ROW_HEIGHT;
    // eslint-disable-next-line no-loop-func
    row.eachCell((cell, colNumber) => {
      cell.border = thinBorder;
      cell.alignment = { vertical: 'middle', horizontal: colNumber <= 2 ? 'center' : (colNumber >= 5 ? 'center' : 'left') };
      if (idx % 2 === 0) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      }
    });
    row.getCell(4).alignment = { vertical: 'middle', wrapText: true };

    // Add photo
    if (data.photo) {
      try {
        const base64Match = data.photo.match(/^data:image\/(\w+);base64,(.+)/);
        if (base64Match) {
          const ext = base64Match[1] === 'png' ? 'png' : 'jpeg';
          const imageId = wb.addImage({
            base64: base64Match[2],
            extension: ext
          });
          ws.addImage(imageId, {
            tl: { col: 1.1, row: row.number - 1 + 0.1 },
            br: { col: 1.9, row: row.number - 0.1 },
            editAs: 'oneCell'
          });
        }
      } catch (imgErr) {
        console.warn('Could not embed image for', data.markNo, imgErr);
      }
    }
  }

  // ── Total row ──
  const totalRow = ws.addRow([
    '', '', 'TOTAL', '',
    container.totalCartons,
    '', '', container.totalQty || '',
    '', parseFloat((container.totalPrice || 0).toFixed(2)),
    parseFloat(container.totalGW.toFixed(2)),
    parseFloat(container.totalNW.toFixed(2)),
    parseFloat(container.totalCBM.toFixed(2))
  ]);
  totalRow.height = 26;
  totalRow.eachCell((cell) => {
    cell.fill = totalFill;
    cell.font = { bold: true, size: 11 };
    cell.border = thinBorder;
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  totalRow.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };
};

/**
 * Create a styled workbook buffer for containers in their supplied order.
 * @param {Array<Object>} containers Containers to export.
 * @param {{maxWeight: number, maxCbm: number}} limits Numeric capacity limits.
 * @param {Date} created Workbook creation timestamp.
 * @returns {Promise<Buffer>} Serialized XLSX workbook.
 */
export const buildPackingWorkbook = async (containers, limits, created) => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Container Optimizer';
  wb.created = created;

  for (const container of containers) {
    const ws = wb.addWorksheet(`Container ${container.id}`, {
      properties: { defaultRowHeight: 20 }
    });
    buildContainerSheet(wb, ws, container, limits);
  }

  return wb.xlsx.writeBuffer();
};
