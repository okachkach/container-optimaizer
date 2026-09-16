import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { roundDecimal } from '../utils/decimal';

/**
 * Read catalog rows and embedded images from the first workbook sheet.
 * Image extraction failures are logged and leave photos empty.
 * @param {File} file Workbook file; its name identifies the product source.
 * @param {() => number} [now=Date.now] Clock used for each product ID.
 * @returns {Promise<Array<Object>>} Parsed catalog products with per-carton values.
 */
export const extractXlsxData = async (file, now = Date.now) => {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  // Extract images
  const imageMap = {};
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const relsFile = zip.file('xl/drawings/_rels/drawing1.xml.rels');
    const drawingFile = zip.file('xl/drawings/drawing1.xml');
    if (relsFile && drawingFile) {
      const relsXml = await relsFile.async('string');
      const drawingXml = await drawingFile.async('string');
      const rIdMap = {};
      for (const m of relsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)) {
        rIdMap[m[1]] = m[2].replace('../', 'xl/');
      }
      const anchorPattern = /<xdr:oneCellAnchor>([\s\S]*?)<\/xdr:oneCellAnchor>/g;
      let anchor;
      while ((anchor = anchorPattern.exec(drawingXml))) {
        const content = anchor[1];
        const fromRow = content.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
        const rIdMatch = content.match(/r:embed="(rId\d+)"/);
        if (fromRow && rIdMatch) {
          const row = parseInt(fromRow[1]);
          const imagePath = rIdMap[rIdMatch[1]];
          if (imagePath) {
            const imgFile = zip.file(imagePath);
            if (imgFile) {
              const imgData = await imgFile.async('base64');
              const ext = imagePath.split('.').pop().toLowerCase();
              const mimeType = ext === 'png' ? 'image/png' : 'image/jpeg';
              imageMap[row] = `data:${mimeType};base64,${imgData}`;
            }
          }
        }
      }
    }
  } catch (imgErr) {
    console.warn('Could not extract images:', imgErr);
  }

  const allData = [];
  for (let i = 3; i < rows.length; i++) {
    const row = rows[i];
    const markNo = String(row[0] || '').trim();
    if (!markNo || !markNo.match(/^\d+-\d+$/)) continue;

    const description = [row[1], row[2], row[3]].map(v => String(v || '').trim()).filter(Boolean).join(' ') || markNo;
    const ctn = parseInt(row[5]) || 1;
    const pcsPerCtn = parseInt(row[6]) || 0;
    const priceRaw = String(row[8] || '').replace(/[^\d.]/g, '');
    const pricePerPcs = parseFloat(priceRaw) || 0;
    const cbm = parseFloat(row[10]) || 0;
    const weight = parseFloat(row[11]) || 0;
    const photo = imageMap[i] || '';

    allData.push({
      id: `${file.name}-${markNo}-${now()}-${i}`,
      markNo,
      description,
      photo,
      source: file.name,
      pcsPerCtn,
      pricePerPcs,
      gwPerCtn: roundDecimal(weight / ctn),
      nwPerCtn: roundDecimal(weight * 0.9 / ctn),
      cbmPerCtn: roundDecimal(cbm / ctn),
      originalCtn: ctn,
      originalWeight: weight,
      originalCbm: cbm
    });
  }
  return allData;
};
