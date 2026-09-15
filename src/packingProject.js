import ExcelJS from 'exceljs';

export function totalContainer(id, items) {
  return items.reduce((c, item) => ({ ...c,
    totalCartons: c.totalCartons + 1,
    totalGW: c.totalGW + item.gwPerCtn,
    totalNW: c.totalNW + item.nwPerCtn,
    totalCBM: c.totalCBM + item.cbmPerCtn,
    totalQty: c.totalQty + item.pcsPerCtn,
    totalPrice: c.totalPrice + item.pricePerPcs * item.pcsPerCtn
  }), { id, items, totalCartons: 0, totalGW: 0, totalNW: 0, totalCBM: 0, totalQty: 0, totalPrice: 0 });
}

function value(cell) {
  const v = cell.value;
  if (v && typeof v === 'object') {
    if ('formula' in v || 'sharedFormula' in v) {
      if (v.result === undefined) throw new Error('Recalculate formulas and save the workbook in Excel before importing.');
      return v.result;
    }
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.text) return v.text;
    throw new Error('Unsupported spreadsheet cell value.');
  }
  return v;
}

export async function importPackingWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const containers = [], containerNames = {}, catalog = [], shipment = [];
  let totalCartons = 0;
  for (const ws of wb.worksheets) {
    let header;
    ws.eachRow((row) => {
      if (!header && row.values.includes('MARKS&NO') && row.values.includes('PCS/CTN')) header = row;
    });
    if (!header) continue;
    const columns = {};
    header.eachCell((cell, col) => { columns[String(value(cell)).trim()] = col; });
    for (const name of ['MARKS&NO', 'DESCRIPTION', 'CTN', 'PCS/CTN', 'U/PRICE', 'G.W.(KGS)', 'N.W.(KGS)', 'CBM']) {
      if (!columns[name]) throw new Error(`${ws.name}: missing column ${name}`);
    }
    const id = containers.length + 1;
    const items = [];
    for (let r = header.number + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const get = name => value(row.getCell(columns[name]));
      const markNo = String(get('MARKS&NO') ?? '').trim();
      if (markNo === 'TOTAL') break;
      if (!markNo && !get('DESCRIPTION')) continue;
      const number = (name, integer = false) => {
        const raw = get(name);
        const n = Number(raw);
        if (raw === null || raw === undefined || raw === '' || !Number.isFinite(n) || n < 0 || (integer && !Number.isSafeInteger(n))) {
          throw new Error(`${ws.name}, row ${r}: ${name} must be a valid ${integer ? 'whole ' : ''}non-negative number.`);
        }
        return n;
      };
      const ctn = number('CTN', true);
      totalCartons += ctn;
      if (totalCartons > 50000) throw new Error('This workbook exceeds the 50,000 carton import limit.');
      if (!ctn) continue;
      let photo = '';
      const image = ws.getImages().find(img => Math.floor(img.range.tl.nativeRow) === r - 1);
      if (image) {
        const media = wb.getImage(image.imageId);
        if (media && ['png', 'jpeg'].includes(media.extension)) {
          const bytes = new Uint8Array(media.buffer || []);
          if (bytes.length > 2 * 1024 * 1024) throw new Error('An embedded image exceeds 2 MB.');
          photo = media.base64 || `data:image/${media.extension};base64,${btoa(Array.from(bytes, b => String.fromCharCode(b)).join(''))}`;
        }
      }
      const product = { id: `import-${id}-${r}`, catalogId: `import-${id}-${r}`, markNo,
        description: String(get('DESCRIPTION') ?? ''), source: 'Imported packing list', photo,
        pcsPerCtn: number('PCS/CTN', true), pricePerPcs: number('U/PRICE'),
        gwPerCtn: number('G.W.(KGS)') / ctn, nwPerCtn: number('N.W.(KGS)') / ctn,
        cbmPerCtn: number('CBM') / ctn, ctn };
      catalog.push(product);
      const totals = totalContainer(id, Array(ctn).fill(product));
      shipment.push({ ...product, totalQty: totals.totalQty, totalGW: totals.totalGW,
        totalNW: totals.totalNW, totalCBM: totals.totalCBM, totalPrice: totals.totalPrice });
      for (let i = 0; i < ctn; i++) items.push({ ...product });
    }
    containers.push(totalContainer(id, items));
    containerNames[id] = ws.name;
  }
  if (!containers.length) throw new Error('No packing-list sheets found. Import an XLSX exported by this app.');
  return { version: 1, containers, containerNames, catalog, shipment, catalogSources: ['Imported packing list'] };
}

export function downloadProject(project, name = 'container-project') {
  const url = URL.createObjectURL(new Blob([JSON.stringify({ ...project, version: 1 })], { type: 'application/json' }));
  const link = document.createElement('a');
  link.href = url;
  link.download = `${name.replace(/[^a-zA-Z0-9_-]/g, '_')}.json`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function validateProject(data) {
  if (!data || data.version !== 1 || !Array.isArray(data.containers) || !Array.isArray(data.shipment) || !Array.isArray(data.catalog)) throw new Error('Invalid project backup or unsupported version.');
  if (data.capacity && ['maxWeight', 'maxCbm'].some(key => !Number.isFinite(Number(data.capacity[key])) || Number(data.capacity[key]) <= 0)) throw new Error('Invalid container capacity.');
  if (!Array.isArray(data.catalogSources) || data.catalogSources.some(s => typeof s !== 'string')) throw new Error('Invalid catalog sources.');
  if (data.containerNames && Object.values(data.containerNames).some(s => typeof s !== 'string')) throw new Error('Invalid container names.');
  let count = 0;
  const ids = new Set();
  for (const c of data.containers) {
    if (!c || !Array.isArray(c.items) || !Number.isSafeInteger(c.id) || c.id <= 0 || ids.has(c.id)) throw new Error('Invalid container items or ID.');
    ids.add(c.id);
    count += c.items.length;
  }
  if (count > 50000) throw new Error('Project exceeds 50,000 cartons.');
  for (const p of [...data.catalog, ...data.shipment, ...data.containers.flatMap(c => c.items)]) {
    if (!p || typeof p.markNo !== 'string' || typeof p.description !== 'string' || typeof p.source !== 'string') throw new Error('Invalid product fields.');
    if (p.photo && (typeof p.photo !== 'string' || !/^data:image\/(png|jpeg|gif|webp);base64,/.test(p.photo))) throw new Error('Invalid product photo.');
    for (const key of ['pcsPerCtn', 'pricePerPcs', 'gwPerCtn', 'nwPerCtn', 'cbmPerCtn']) {
      if (typeof p[key] !== 'number' || !Number.isFinite(p[key]) || p[key] < 0) throw new Error(`Invalid product ${key}.`);
    }
  }
  const shipment = data.shipment.map(p => {
    if (!Number.isSafeInteger(p.ctn) || p.ctn < 0 || p.ctn > 50000) throw new Error('Invalid shipment quantity.');
    return { ...p, totalQty: p.ctn * p.pcsPerCtn, totalGW: p.ctn * p.gwPerCtn,
      totalNW: p.ctn * p.nwPerCtn, totalCBM: p.ctn * p.cbmPerCtn, totalPrice: p.ctn * p.pcsPerCtn * p.pricePerPcs };
  });
  return { ...data, shipment, containers: data.containers.map(c => totalContainer(c.id, c.items)) };
}

export function validateLibrary(data) {
  if (!data || data.type !== 'container-optimizer-library' || data.version !== 1 ||
      !Array.isArray(data.projects) || !Array.isArray(data.history) || !data.projects.length) {
    throw new Error('Invalid all-projects backup or unsupported version.');
  }
  const ids = new Set();
  const projects = data.projects.map(project => {
    if (typeof project.id !== 'string' || !project.id || ids.has(project.id)) throw new Error('Invalid or duplicate project ID.');
    ids.add(project.id);
    return validateProject(project);
  });
  const history = data.history.map(entry => {
    if (!entry || typeof entry.id !== 'string' || !ids.has(entry.projectId) || typeof entry.name !== 'string' ||
        typeof entry.createdAt !== 'string') throw new Error('Invalid project history.');
    return { ...entry, project: validateProject(entry.project) };
  });
  return { ...data, projects, history };
}
