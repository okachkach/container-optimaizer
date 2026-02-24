const fs = require('fs');
const JSZip = require('jszip');

async function main() {
  const data = fs.readFileSync('193所有产品清单.xlsx');
  const zip = await JSZip.loadAsync(data);
  
  // Read drawing XML
  const drawingXml = await zip.file('xl/drawings/drawing1.xml').async('string');
  
  // Check what anchor types exist
  const anchorTypes = drawingXml.match(/<xdr:\w+Anchor/g);
  console.log('Anchor types found:', [...new Set(anchorTypes || [])]);
  
  // Try oneCellAnchor
  const oneCell = (drawingXml.match(/<xdr:oneCellAnchor/g) || []).length;
  const twoCell = (drawingXml.match(/<xdr:twoCellAnchor/g) || []).length;
  const absAnchor = (drawingXml.match(/<xdr:absoluteAnchor/g) || []).length;
  console.log('oneCellAnchor:', oneCell, 'twoCellAnchor:', twoCell, 'absoluteAnchor:', absAnchor);
  
  // Extract first few anchors regardless of type
  const anchorPattern = /<xdr:oneCellAnchor>([\s\S]*?)<\/xdr:oneCellAnchor>/g;
  let anchor;
  let count = 0;
  
  // Read rels
  const relsXml = await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string');
  const rIdMap = {};
  const relMatches = relsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g);
  for (const m of relMatches) {
    rIdMap[m[1]] = m[2];
  }
  
  while ((anchor = anchorPattern.exec(drawingXml)) && count < 5) {
    const content = anchor[1];
    const fromRow = content.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/);
    const fromCol = content.match(/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/);
    const rIdMatch = content.match(/r:embed="(rId\d+)"/);
    if (fromRow && rIdMatch) {
      const row = parseInt(fromRow[1]);
      const col = fromCol ? parseInt(fromCol[1]) : '?';
      const imageFile = rIdMap[rIdMatch[1]];
      console.log('Image: row=' + row + ', col=' + col + ', file=' + imageFile);
    }
    count++;
  }
}
main();

