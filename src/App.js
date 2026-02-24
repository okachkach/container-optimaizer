import React, { useState, useMemo, useCallback } from 'react';
import { Plus, Download, Image, FileSpreadsheet, Upload, Search, ShoppingCart, Package, ChevronDown, ChevronUp, X } from 'lucide-react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';

const ContainerOptimizer = () => {
  // Catalog: all products from uploaded XLSX files
  const [catalog, setCatalog] = useState([]);
  const [catalogSources, setCatalogSources] = useState([]);
  // Shipment: selected products to ship
  const [shipment, setShipment] = useState([]);
  // Containers result
  const [containers, setContainers] = useState([]);
  // Search/filter
  const [searchTerm, setSearchTerm] = useState('');
  // UI state
  const [showCatalog, setShowCatalog] = useState(true);
  const [showManualAdd, setShowManualAdd] = useState(false);
  // Container capacity
  const [capacity, setCapacity] = useState({ maxWeight: 27000, maxCbm: 76 });
  const maxWeight = parseFloat(capacity.maxWeight) || 0;
  const maxCbm = parseFloat(capacity.maxCbm) || 0;

  // Manual add form
  const [newProduct, setNewProduct] = useState({
    markNo: '', description: '', photo: '', pcsPerCtn: '', gwPerCtn: '', nwPerCtn: '', cbmPerCtn: '', pricePerPcs: ''
  });

  // CTN input per catalog row
  const [ctnInputs, setCtnInputs] = useState({});

  // ─── XLSX PARSING ───────────────────────────────────────────────
  const extractXlsxData = async (file) => {
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
        id: `${file.name}-${markNo}-${Date.now()}-${i}`,
        markNo,
        description,
        photo,
        source: file.name,
        pcsPerCtn,
        pricePerPcs,
        gwPerCtn: parseFloat((weight / ctn).toFixed(2)),
        nwPerCtn: parseFloat((weight * 0.9 / ctn).toFixed(2)),
        cbmPerCtn: parseFloat((cbm / ctn).toFixed(4)),
        originalCtn: ctn,
        originalWeight: weight,
        originalCbm: cbm
      });
    }
    return allData;
  };

  const handleXlsxUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const data = await extractXlsxData(file);
      setCatalog(prev => [...prev, ...data]);
      setCatalogSources(prev => [...prev, file.name]);
      const withPhotos = data.filter(p => p.photo).length;
      alert(`Loaded ${data.length} products from "${file.name}" (${withPhotos} with photos)`);
    } catch (error) {
      alert('Failed to load XLSX: ' + error.message);
    }
    e.target.value = '';
  };

  const removeCatalogSource = (sourceName) => {
    setCatalog(prev => prev.filter(p => p.source !== sourceName));
    setCatalogSources(prev => prev.filter(s => s !== sourceName));
    setShipment(prev => prev.filter(s => s.source !== sourceName));
  };

  // ─── CATALOG ACTIONS ────────────────────────────────────────────
  const filteredCatalog = useMemo(() => {
    if (!searchTerm.trim()) return catalog;
    const term = searchTerm.toLowerCase();
    return catalog.filter(p =>
      p.markNo.toLowerCase().includes(term) ||
      p.description.toLowerCase().includes(term) ||
      p.source.toLowerCase().includes(term)
    );
  }, [catalog, searchTerm]);

  const addToShipment = (catalogProduct, ctn) => {
    const qty = parseInt(ctn) || 1;
    const existing = shipment.find(s => s.catalogId === catalogProduct.id);
    if (existing) {
      setShipment(shipment.map(s =>
        s.catalogId === catalogProduct.id
          ? {
            ...s,
            ctn: s.ctn + qty,
            totalQty: (s.ctn + qty) * s.pcsPerCtn,
            totalGW: s.gwPerCtn * (s.ctn + qty),
            totalNW: s.nwPerCtn * (s.ctn + qty),
            totalCBM: s.cbmPerCtn * (s.ctn + qty),
            totalPrice: (s.pricePerPcs || 0) * (s.ctn + qty) * s.pcsPerCtn
          }
          : s
      ));
    } else {
      setShipment([...shipment, {
        id: Date.now() + Math.random(),
        catalogId: catalogProduct.id,
        markNo: catalogProduct.markNo,
        description: catalogProduct.description,
        photo: catalogProduct.photo,
        source: catalogProduct.source,
        ctn: qty,
        pcsPerCtn: catalogProduct.pcsPerCtn,
        pricePerPcs: catalogProduct.pricePerPcs || 0,
        totalQty: qty * catalogProduct.pcsPerCtn,
        gwPerCtn: catalogProduct.gwPerCtn,
        nwPerCtn: catalogProduct.nwPerCtn,
        cbmPerCtn: catalogProduct.cbmPerCtn,
        totalGW: catalogProduct.gwPerCtn * qty,
        totalNW: catalogProduct.nwPerCtn * qty,
        totalCBM: catalogProduct.cbmPerCtn * qty,
        totalPrice: (catalogProduct.pricePerPcs || 0) * qty * catalogProduct.pcsPerCtn
      }]);
    }
  };

  const updateShipmentCtn = (id, newCtn) => {
    const ctn = parseInt(newCtn) || 0;
    if (ctn <= 0) {
      removeFromShipment(id);
      return;
    }
    setShipment(shipment.map(s =>
      s.id === id
        ? {
          ...s,
          ctn,
          totalQty: ctn * s.pcsPerCtn,
          totalGW: s.gwPerCtn * ctn,
          totalNW: s.nwPerCtn * ctn,
          totalCBM: s.cbmPerCtn * ctn,
          totalPrice: (s.pricePerPcs || 0) * ctn * s.pcsPerCtn
        }
        : s
    ));
  };

  const removeFromShipment = (id) => {
    setShipment(shipment.filter(s => s.id !== id));
  };

  // ─── MANUAL ADD ─────────────────────────────────────────────────
  const handlePhotoUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => setNewProduct({ ...newProduct, photo: reader.result });
      reader.readAsDataURL(file);
    }
  };

  const addManualProduct = () => {
    if (!newProduct.description) return;
    const pcsPerCtn = parseInt(newProduct.pcsPerCtn) || 0;
    const gwPerCtn = parseFloat(newProduct.gwPerCtn) || 0;
    const nwPerCtn = parseFloat(newProduct.nwPerCtn) || 0;
    const cbmPerCtn = parseFloat(newProduct.cbmPerCtn) || 0;
    const pricePerPcs = parseFloat(newProduct.pricePerPcs) || 0;

    const catProduct = {
      id: `manual-${Date.now()}`,
      markNo: newProduct.markNo || 'M-' + (catalog.length + 1),
      description: newProduct.description,
      photo: newProduct.photo,
      source: 'Manual',
      pcsPerCtn,
      pricePerPcs,
      gwPerCtn,
      nwPerCtn,
      cbmPerCtn,
      originalCtn: 1,
      originalWeight: gwPerCtn,
      originalCbm: cbmPerCtn
    };
    setCatalog(prev => [...prev, catProduct]);
    setNewProduct({ markNo: '', description: '', photo: '', pcsPerCtn: '', gwPerCtn: '', nwPerCtn: '', cbmPerCtn: '', pricePerPcs: '' });
  };

  // ─── CONTAINER OPTIMIZATION ─────────────────────────────────────
  const optimizeContainers = () => {
    const expandedProducts = [];
    shipment.forEach(product => {
      for (let i = 0; i < product.ctn; i++) {
        expandedProducts.push({
          ...product,
          singleCtnGW: product.gwPerCtn,
          singleCtnNW: product.nwPerCtn,
          singleCtnCBM: product.cbmPerCtn
        });
      }
    });

    expandedProducts.sort((a, b) => b.singleCtnGW - a.singleCtnGW);

    const newContainers = [];
    let currentContainer = { id: 1, items: [], totalGW: 0, totalNW: 0, totalCBM: 0, totalCartons: 0, totalPrice: 0, totalQty: 0 };

    expandedProducts.forEach(product => {
      if (currentContainer.totalGW + product.singleCtnGW <= maxWeight &&
        currentContainer.totalCBM + product.singleCtnCBM <= maxCbm) {
        currentContainer.items.push(product);
        currentContainer.totalGW += product.singleCtnGW;
        currentContainer.totalNW += product.singleCtnNW;
        currentContainer.totalCBM += product.singleCtnCBM;
        currentContainer.totalCartons += 1;
        currentContainer.totalPrice += (product.pricePerPcs || 0) * product.pcsPerCtn;
        currentContainer.totalQty += product.pcsPerCtn;
      } else {
        if (currentContainer.items.length > 0) newContainers.push(currentContainer);
        currentContainer = {
          id: newContainers.length + 2,
          items: [product],
          totalGW: product.singleCtnGW,
          totalNW: product.singleCtnNW,
          totalCBM: product.singleCtnCBM,
          totalCartons: 1,
          totalPrice: (product.pricePerPcs || 0) * product.pcsPerCtn,
          totalQty: product.pcsPerCtn
        };
      }
    });
    if (currentContainer.items.length > 0) newContainers.push(currentContainer);
    setContainers(newContainers);
  };

  // ─── CONTAINER MODIFICATION ─────────────────────────────────────
  const recalcContainerTotals = (container) => {
    let totalGW = 0, totalNW = 0, totalCBM = 0, totalCartons = 0, totalPrice = 0, totalQty = 0;
    container.items.forEach(item => {
      totalGW += item.gwPerCtn || item.singleCtnGW || 0;
      totalNW += item.nwPerCtn || item.singleCtnNW || 0;
      totalCBM += item.cbmPerCtn || item.singleCtnCBM || 0;
      totalCartons += 1;
      totalPrice += (item.pricePerPcs || 0) * (item.pcsPerCtn || 0);
      totalQty += item.pcsPerCtn || 0;
    });
    return { ...container, totalGW, totalNW, totalCBM, totalCartons, totalPrice, totalQty };
  };

  const moveProductBetweenContainers = (fromContainerId, productKey, ctnCount, toContainerId) => {
    setContainers(prev => {
      const updated = prev.map(c => ({ ...c, items: [...c.items] }));
      const fromIdx = updated.findIndex(c => c.id === fromContainerId);
      const toIdx = updated.findIndex(c => c.id === toContainerId);
      if (fromIdx === -1 || toIdx === -1 || fromIdx === toIdx) return prev;

      const [markNo, description] = productKey.split('|');
      // Collect matching items from source container
      const matchingIndices = [];
      updated[fromIdx].items.forEach((item, i) => {
        if (item.markNo === markNo && item.description === description) {
          matchingIndices.push(i);
        }
      });

      const moveCount = Math.min(ctnCount, matchingIndices.length);
      if (moveCount <= 0) return prev;

      // Move items (take from end to keep indices stable)
      const indicesToMove = matchingIndices.slice(-moveCount);
      const movedItems = indicesToMove.map(i => updated[fromIdx].items[i]);

      // Remove moved items from source
      const removeSet = new Set(indicesToMove);
      updated[fromIdx].items = updated[fromIdx].items.filter((_, i) => !removeSet.has(i));

      // Add to target
      updated[toIdx].items.push(...movedItems);

      // Recalculate totals
      updated[fromIdx] = recalcContainerTotals(updated[fromIdx]);
      updated[toIdx] = recalcContainerTotals(updated[toIdx]);

      // Remove empty containers
      return updated.filter(c => c.items.length > 0);
    });
  };

  const removeProductFromContainer = (containerId, productKey, ctnCount) => {
    setContainers(prev => {
      const updated = prev.map(c => ({ ...c, items: [...c.items] }));
      const idx = updated.findIndex(c => c.id === containerId);
      if (idx === -1) return prev;

      const [markNo, description] = productKey.split('|');
      const matchingIndices = [];
      updated[idx].items.forEach((item, i) => {
        if (item.markNo === markNo && item.description === description) {
          matchingIndices.push(i);
        }
      });

      const removeCount = Math.min(ctnCount, matchingIndices.length);
      if (removeCount <= 0) return prev;

      const indicesToRemove = new Set(matchingIndices.slice(-removeCount));
      updated[idx].items = updated[idx].items.filter((_, i) => !indicesToRemove.has(i));
      updated[idx] = recalcContainerTotals(updated[idx]);

      return updated.filter(c => c.items.length > 0);
    });
  };

  // ─── EXPORT HELPER: build worksheet for one container ─────────
  const buildContainerSheet = (wb, ws, container) => {
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

  // ─── EXPORT ALL CONTAINERS ─────────────────────────────────────
  const exportToXlsx = async () => {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Container Optimizer';
    wb.created = new Date();

    for (const container of containers) {
      const ws = wb.addWorksheet(`Container ${container.id}`, {
        properties: { defaultRowHeight: 20 }
      });
      buildContainerSheet(wb, ws, container);
    }

    // Generate and download
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'packing_list.xlsx';
    a.click();
    window.URL.revokeObjectURL(url);
  };

  // ─── EXPORT SINGLE CONTAINER ───────────────────────────────────
  const exportSingleContainer = useCallback(async (container) => {
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Container Optimizer';
    wb.created = new Date();

    const ws = wb.addWorksheet(`Container ${container.id}`, {
      properties: { defaultRowHeight: 20 }
    });
    buildContainerSheet(wb, ws, container);

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `packing_list_container_${container.id}.xlsx`;
    a.click();
    window.URL.revokeObjectURL(url);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [maxWeight, maxCbm]);

  // ─── SHIPMENT TOTALS ───────────────────────────────────────────
  const shipmentTotals = useMemo(() => {
    return shipment.reduce((acc, s) => ({
      ctn: acc.ctn + s.ctn,
      qty: acc.qty + s.totalQty,
      gw: acc.gw + s.totalGW,
      nw: acc.nw + s.totalNW,
      cbm: acc.cbm + s.totalCBM,
      price: acc.price + (s.totalPrice || 0)
    }), { ctn: 0, qty: 0, gw: 0, nw: 0, cbm: 0, price: 0 });
  }, [shipment]);

  // ─── GRAND TOTALS ACROSS ALL CONTAINERS ────────────────────────
  const grandTotals = useMemo(() => {
    return containers.reduce((acc, c) => ({
      cartons: acc.cartons + c.totalCartons,
      qty: acc.qty + (c.totalQty || 0),
      gw: acc.gw + c.totalGW,
      nw: acc.nw + c.totalNW,
      cbm: acc.cbm + c.totalCBM,
      price: acc.price + (c.totalPrice || 0)
    }), { cartons: 0, qty: 0, gw: 0, nw: 0, cbm: 0, price: 0 });
  }, [containers]);

  // ─── RENDER ─────────────────────────────────────────────────────
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 p-4">
      <div className="max-w-7xl mx-auto">

        {/* HEADER */}
        <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
          <div className="flex items-center gap-3 mb-2">
            <FileSpreadsheet className="w-8 h-8 text-blue-600" />
            <h1 className="text-2xl font-bold text-gray-800">Container Packing List Generator</h1>
          </div>
          <div className="flex flex-wrap gap-4 items-center text-sm">
            <div className="flex items-center gap-2">
              <label className="text-gray-600 font-medium">Max Weight (kg):</label>
              <input type="number" value={capacity.maxWeight} onChange={(e) => setCapacity({ ...capacity, maxWeight: e.target.value })}
                className="w-28 px-2 py-1 border rounded text-center" />
            </div>
            <div className="flex items-center gap-2">
              <label className="text-gray-600 font-medium">Max CBM (m³):</label>
              <input type="number" value={capacity.maxCbm} onChange={(e) => setCapacity({ ...capacity, maxCbm: e.target.value })}
                className="w-28 px-2 py-1 border rounded text-center" />
            </div>
            <div className="ml-auto flex items-center gap-2 text-gray-500">
              <Package className="w-4 h-4" />
              <span>{catalog.length} in catalog</span>
              <span className="mx-1">|</span>
              <ShoppingCart className="w-4 h-4" />
              <span>{shipment.length} in shipment ({shipmentTotals.ctn} CTN)</span>
            </div>
          </div>
        </div>

        {/* CATALOG SECTION */}
        <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-3 cursor-pointer" onClick={() => setShowCatalog(!showCatalog)}>
              {showCatalog ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
              <h2 className="text-xl font-bold text-gray-800">Product Catalog</h2>
            </div>
            <label className="flex items-center gap-2 cursor-pointer bg-blue-50 px-4 py-2 rounded border-2 border-blue-300 hover:bg-blue-100 transition">
              <Upload className="w-4 h-4 text-blue-600" />
              <span className="text-sm font-medium text-blue-700">Upload XLSX</span>
              <input type="file" accept=".xlsx,.xls" onChange={handleXlsxUpload} className="hidden" />
            </label>
          </div>

          {/* Source file tags */}
          {catalogSources.length > 0 && (
            <div className="flex flex-wrap gap-2 mb-3">
              {catalogSources.map((src, i) => (
                <span key={i} className="inline-flex items-center gap-1 bg-gray-100 text-gray-700 text-xs px-3 py-1 rounded-full">
                  {src}
                  <button onClick={() => removeCatalogSource(src)} className="text-gray-400 hover:text-red-500 ml-1">
                    <X className="w-3 h-3" />
                  </button>
                </span>
              ))}
            </div>
          )}

          {showCatalog && (
            <>
              {/* Search */}
              <div className="relative mb-4">
                <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input
                  type="text"
                  placeholder="Search by name, mark number, or source file..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>

              {/* Catalog Table */}
              {filteredCatalog.length > 0 ? (
                <div className="border rounded-lg overflow-x-auto max-h-[500px] overflow-y-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-50 sticky top-0">
                      <tr>
                        <th className="px-3 py-2 text-left font-bold border-b">Photo</th>
                        <th className="px-3 py-2 text-left font-bold border-b">Mark No</th>
                        <th className="px-3 py-2 text-left font-bold border-b">Description</th>
                        <th className="px-3 py-2 text-left font-bold border-b">PCS/CTN</th>
                        <th className="px-3 py-2 text-left font-bold border-b">U/Price</th>
                        <th className="px-3 py-2 text-left font-bold border-b">G.W./CTN</th>
                        <th className="px-3 py-2 text-left font-bold border-b">CBM/CTN</th>
                        <th className="px-3 py-2 text-left font-bold border-b">Source</th>
                        <th className="px-3 py-2 text-left font-bold border-b">Add to Shipment</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredCatalog.map((product) => {
                        const inShipment = shipment.find(s => s.catalogId === product.id);
                        return (
                          <tr key={product.id} className={`border-b hover:bg-blue-50 transition ${inShipment ? 'bg-green-50' : ''}`}>
                            <td className="px-3 py-2">
                              {product.photo ? (
                                <img src={product.photo} alt="" className="w-12 h-12 object-cover rounded border" />
                              ) : (
                                <div className="w-12 h-12 bg-gray-100 rounded border flex items-center justify-center">
                                  <Image className="w-4 h-4 text-gray-300" />
                                </div>
                              )}
                            </td>
                            <td className="px-3 py-2 font-mono text-xs">{product.markNo}</td>
                            <td className="px-3 py-2 font-medium max-w-48 truncate">{product.description}</td>
                            <td className="px-3 py-2">{product.pcsPerCtn}</td>
                            <td className="px-3 py-2 text-green-700 font-medium">{product.pricePerPcs ? `¥${product.pricePerPcs.toFixed(2)}` : '-'}</td>
                            <td className="px-3 py-2">{product.gwPerCtn}</td>
                            <td className="px-3 py-2">{product.cbmPerCtn}</td>
                            <td className="px-3 py-2 text-xs text-gray-500 max-w-24 truncate">{product.source}</td>
                            <td className="px-3 py-2">
                              <div className="flex items-center gap-1">
                                <input
                                  type="number"
                                  min="1"
                                  placeholder="CTN"
                                  value={ctnInputs[product.id] || ''}
                                  onChange={(e) => setCtnInputs({ ...ctnInputs, [product.id]: e.target.value })}
                                  className="w-16 px-2 py-1 border rounded text-center text-sm"
                                />
                                <button
                                  onClick={() => {
                                    addToShipment(product, ctnInputs[product.id] || 1);
                                    setCtnInputs({ ...ctnInputs, [product.id]: '' });
                                  }}
                                  className="bg-green-600 text-white px-2 py-1 rounded hover:bg-green-700 transition text-xs"
                                >
                                  <Plus className="w-3 h-3" />
                                </button>
                                {inShipment && (
                                  <span className="text-xs text-green-700 font-medium ml-1">({inShipment.ctn})</span>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              ) : catalog.length === 0 ? (
                <div className="text-center py-8 text-gray-400">
                  <Upload className="w-12 h-12 mx-auto mb-2" />
                  <p>Upload one or more XLSX files to build your product catalog</p>
                </div>
              ) : (
                <p className="text-center py-4 text-gray-400">No products match "{searchTerm}"</p>
              )}
            </>
          )}
        </div>

        {/* MANUAL ADD SECTION */}
        <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
          <div className="flex items-center gap-3 cursor-pointer mb-4" onClick={() => setShowManualAdd(!showManualAdd)}>
            {showManualAdd ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
            <h2 className="text-xl font-bold text-gray-800">Add Product Manually</h2>
          </div>

          {showManualAdd && (
            <div className="border rounded-lg p-4 bg-slate-50">
              <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-3 mb-3">
                <input type="text" placeholder="MARKS&NO" value={newProduct.markNo}
                  onChange={(e) => setNewProduct({ ...newProduct, markNo: e.target.value })}
                  className="px-3 py-2 border rounded text-sm" />
                <input type="text" placeholder="Description *" value={newProduct.description}
                  onChange={(e) => setNewProduct({ ...newProduct, description: e.target.value })}
                  className="px-3 py-2 border rounded text-sm col-span-2" />
                <input type="number" placeholder="PCS/CTN" value={newProduct.pcsPerCtn}
                  onChange={(e) => setNewProduct({ ...newProduct, pcsPerCtn: e.target.value })}
                  className="px-3 py-2 border rounded text-sm" />
                <input type="number" step="0.01" placeholder="U/Price (¥)" value={newProduct.pricePerPcs}
                  onChange={(e) => setNewProduct({ ...newProduct, pricePerPcs: e.target.value })}
                  className="px-3 py-2 border rounded text-sm" />
                <input type="number" step="0.01" placeholder="G.W./CTN (kg)" value={newProduct.gwPerCtn}
                  onChange={(e) => setNewProduct({ ...newProduct, gwPerCtn: e.target.value })}
                  className="px-3 py-2 border rounded text-sm" />
                <input type="number" step="0.01" placeholder="N.W./CTN (kg)" value={newProduct.nwPerCtn}
                  onChange={(e) => setNewProduct({ ...newProduct, nwPerCtn: e.target.value })}
                  className="px-3 py-2 border rounded text-sm" />
                <input type="number" step="0.0001" placeholder="CBM/CTN" value={newProduct.cbmPerCtn}
                  onChange={(e) => setNewProduct({ ...newProduct, cbmPerCtn: e.target.value })}
                  className="px-3 py-2 border rounded text-sm" />
              </div>
              <div className="flex items-center gap-3">
                <label className="flex items-center gap-2 cursor-pointer bg-white px-3 py-2 rounded border text-sm hover:bg-gray-50">
                  <Image className="w-4 h-4 text-gray-500" />
                  <span>Photo</span>
                  <input type="file" accept="image/*" onChange={handlePhotoUpload} className="hidden" />
                </label>
                {newProduct.photo && (
                  <img src={newProduct.photo} alt="" className="w-10 h-10 object-cover rounded border" />
                )}
                <button onClick={addManualProduct}
                  className="ml-auto bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700 transition text-sm flex items-center gap-1">
                  <Plus className="w-4 h-4" /> Add to Catalog
                </button>
              </div>
            </div>
          )}
        </div>

        {/* SHIPMENT LIST */}
        {shipment.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <ShoppingCart className="w-6 h-6 text-green-600" />
                Shipment List ({shipment.length} products, {shipmentTotals.ctn} CTN)
              </h2>
              <div className="flex gap-2 text-xs">
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded">G.W: {shipmentTotals.gw.toFixed(1)} kg</span>
                <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded">CBM: {shipmentTotals.cbm.toFixed(2)} m³</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded">¥{shipmentTotals.price.toFixed(2)}</span>
                <span className="bg-orange-100 text-orange-800 px-2 py-1 rounded">
                  ~{Math.ceil(Math.max(shipmentTotals.gw / (maxWeight || 1), shipmentTotals.cbm / (maxCbm || 1)))} container(s)
                </span>
              </div>
            </div>

            <div className="border rounded-lg overflow-x-auto mb-4">
              <table className="w-full text-sm">
                <thead className="bg-green-50">
                  <tr>
                    <th className="px-3 py-2 text-left font-bold border-b">Photo</th>
                    <th className="px-3 py-2 text-left font-bold border-b">Mark No</th>
                    <th className="px-3 py-2 text-left font-bold border-b">Description</th>
                    <th className="px-3 py-2 text-left font-bold border-b">CTN</th>
                    <th className="px-3 py-2 text-left font-bold border-b">PCS/CTN</th>
                    <th className="px-3 py-2 text-left font-bold border-b">T/QTY</th>
                    <th className="px-3 py-2 text-left font-bold border-b">U/Price</th>
                    <th className="px-3 py-2 text-left font-bold border-b">Amount</th>
                    <th className="px-3 py-2 text-left font-bold border-b">G.W.</th>
                    <th className="px-3 py-2 text-left font-bold border-b">N.W.</th>
                    <th className="px-3 py-2 text-left font-bold border-b">CBM</th>
                    <th className="px-3 py-2 text-left font-bold border-b">Source</th>
                    <th className="px-3 py-2 text-left font-bold border-b"></th>
                  </tr>
                </thead>
                <tbody>
                  {shipment.map((item) => (
                    <tr key={item.id} className="border-b hover:bg-gray-50">
                      <td className="px-3 py-2">
                        {item.photo ? (
                          <img src={item.photo} alt="" className="w-10 h-10 object-cover rounded border" />
                        ) : (
                          <div className="w-10 h-10 bg-gray-100 rounded border flex items-center justify-center">
                            <Image className="w-3 h-3 text-gray-300" />
                          </div>
                        )}
                      </td>
                      <td className="px-3 py-2 font-mono text-xs">{item.markNo}</td>
                      <td className="px-3 py-2 font-medium max-w-40 truncate">{item.description}</td>
                      <td className="px-3 py-2">
                        <input
                          type="number"
                          min="1"
                          value={item.ctn}
                          onChange={(e) => updateShipmentCtn(item.id, e.target.value)}
                          className="w-16 px-2 py-1 border rounded text-center text-sm"
                        />
                      </td>
                      <td className="px-3 py-2">{item.pcsPerCtn}</td>
                      <td className="px-3 py-2 font-medium">{item.totalQty}</td>
                      <td className="px-3 py-2 text-green-700">{item.pricePerPcs ? `¥${item.pricePerPcs.toFixed(2)}` : '-'}</td>
                      <td className="px-3 py-2 font-medium text-green-700">¥{(item.totalPrice || 0).toFixed(2)}</td>
                      <td className="px-3 py-2">{item.totalGW.toFixed(2)}</td>
                      <td className="px-3 py-2">{item.totalNW.toFixed(2)}</td>
                      <td className="px-3 py-2">{item.totalCBM.toFixed(4)}</td>
                      <td className="px-3 py-2 text-xs text-gray-400 max-w-20 truncate">{item.source}</td>
                      <td className="px-3 py-2">
                        <button onClick={() => removeFromShipment(item.id)} className="text-red-500 hover:text-red-700">
                          <X className="w-4 h-4" />
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <button
              onClick={optimizeContainers}
              className="bg-green-600 text-white px-8 py-3 rounded-lg hover:bg-green-700 transition font-semibold text-lg w-full"
            >
              Generate Container Loading Plan
            </button>
          </div>
        )}

        {/* CONTAINER RESULTS */}
        {containers.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800">
                Packing List: {containers.length} Container{containers.length > 1 ? 's' : ''}
              </h2>
              <button onClick={exportToXlsx}
                className="bg-blue-600 text-white px-6 py-2 rounded hover:bg-blue-700 transition flex items-center gap-2">
                <Download className="w-4 h-4" /> Export All XLSX
              </button>
            </div>

            {/* ── Grand Totals ── */}
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3 mb-5 text-sm">
              <div className="bg-slate-100 p-3 rounded border">
                <p className="text-gray-500 text-xs">Total Cartons</p>
                <p className="text-lg font-bold text-slate-800">{grandTotals.cartons}</p>
              </div>
              <div className="bg-slate-100 p-3 rounded border">
                <p className="text-gray-500 text-xs">Total T/QTY</p>
                <p className="text-lg font-bold text-slate-800">{grandTotals.qty}</p>
              </div>
              <div className="bg-green-50 p-3 rounded border border-green-200">
                <p className="text-gray-500 text-xs">Total G.W.</p>
                <p className="text-lg font-bold text-green-700">{grandTotals.gw.toFixed(2)} kg</p>
              </div>
              <div className="bg-purple-50 p-3 rounded border border-purple-200">
                <p className="text-gray-500 text-xs">Total N.W.</p>
                <p className="text-lg font-bold text-purple-700">{grandTotals.nw.toFixed(2)} kg</p>
              </div>
              <div className="bg-orange-50 p-3 rounded border border-orange-200">
                <p className="text-gray-500 text-xs">Total CBM</p>
                <p className="text-lg font-bold text-orange-700">{grandTotals.cbm.toFixed(2)} m³</p>
              </div>
              <div className="bg-yellow-50 p-3 rounded border border-yellow-200">
                <p className="text-gray-500 text-xs">Total Price</p>
                <p className="text-lg font-bold text-yellow-700">¥{grandTotals.price.toFixed(2)}</p>
              </div>
            </div>

            {containers.map((container) => {
              const weightPercent = maxWeight > 0 ? (container.totalGW / maxWeight * 100) : 0;
              const cbmPercent = maxCbm > 0 ? (container.totalCBM / maxCbm * 100) : 0;

              const groupedItems = {};
              container.items.forEach(item => {
                const key = item.markNo + '|' + item.description;
                if (!groupedItems[key]) {
                  groupedItems[key] = { count: 0, key, markNo: item.markNo, description: item.description, pcsPerCtn: item.pcsPerCtn, gwPerCtn: item.gwPerCtn, nwPerCtn: item.nwPerCtn, cbmPerCtn: item.cbmPerCtn, pricePerPcs: item.pricePerPcs || 0, photo: item.photo };
                }
                groupedItems[key].count++;
              });

              return (
                <div key={container.id} className="border-2 rounded-lg p-5 mb-4 shadow-sm">
                  <div className="flex items-center justify-between bg-slate-100 p-3 rounded mb-3">
                    <h3 className="text-lg font-bold text-gray-800">
                      CONTAINER #{container.id}
                    </h3>
                    <button
                      onClick={() => exportSingleContainer(container)}
                      className="bg-indigo-600 text-white px-4 py-1.5 rounded hover:bg-indigo-700 transition flex items-center gap-2 text-sm"
                    >
                      <Download className="w-3.5 h-3.5" /> Download This Container
                    </button>
                  </div>

                  <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-3 text-sm">
                    <div className="bg-blue-50 p-3 rounded">
                      <p className="text-gray-600">Cartons</p>
                      <p className="text-xl font-bold text-blue-700">{container.totalCartons}</p>
                    </div>
                    <div className="bg-green-50 p-3 rounded">
                      <p className="text-gray-600">G.W.</p>
                      <p className="text-xl font-bold text-green-700">{container.totalGW.toFixed(2)} kg</p>
                    </div>
                    <div className="bg-purple-50 p-3 rounded">
                      <p className="text-gray-600">N.W.</p>
                      <p className="text-xl font-bold text-purple-700">{container.totalNW.toFixed(2)} kg</p>
                    </div>
                    <div className="bg-orange-50 p-3 rounded">
                      <p className="text-gray-600">CBM</p>
                      <p className="text-xl font-bold text-orange-700">{container.totalCBM.toFixed(2)} m³</p>
                    </div>
                    <div className="bg-yellow-50 p-3 rounded">
                      <p className="text-gray-600">Total Price</p>
                      <p className="text-xl font-bold text-yellow-700">¥{(container.totalPrice || 0).toFixed(2)}</p>
                    </div>
                  </div>

                  {/* Utilization bars */}
                  <div className="grid grid-cols-2 gap-4 mb-4">
                    <div>
                      <div className="flex justify-between text-xs mb-1">
                        <span>Weight</span><span>{weightPercent.toFixed(1)}%</span>
                      </div>
                      <div className="w-full bg-gray-200 rounded-full h-2">
                        <div className={`h-2 rounded-full ${weightPercent > 95 ? 'bg-red-500' : weightPercent > 80 ? 'bg-yellow-500' : 'bg-green-500'}`}
                          style={{ width: `${Math.min(weightPercent, 100)}%` }}></div>
                      </div>
                    </div>
                    <div>
                      <div className="flex justify-between text-xs mb-1">
                        <span>Volume</span><span>{cbmPercent.toFixed(1)}%</span>
                      </div>
                      <div className="w-full bg-gray-200 rounded-full h-2">
                        <div className={`h-2 rounded-full ${cbmPercent > 95 ? 'bg-red-500' : cbmPercent > 80 ? 'bg-yellow-500' : 'bg-blue-500'}`}
                          style={{ width: `${Math.min(cbmPercent, 100)}%` }}></div>
                      </div>
                    </div>
                  </div>

                  <div className="border rounded overflow-x-auto">
                    <table className="w-full text-sm">
                      <thead className="bg-gray-50">
                        <tr>
                          <th className="px-3 py-2 text-left font-bold border-b">Photo</th>
                          <th className="px-3 py-2 text-left font-bold border-b">MARKS&NO</th>
                          <th className="px-3 py-2 text-left font-bold border-b">DESCRIPTION</th>
                          <th className="px-3 py-2 text-left font-bold border-b">CTN</th>
                          <th className="px-3 py-2 text-left font-bold border-b">PCS/CTN</th>
                          <th className="px-3 py-2 text-left font-bold border-b">T/QTY</th>
                          <th className="px-3 py-2 text-left font-bold border-b">U/PRICE</th>
                          <th className="px-3 py-2 text-left font-bold border-b">AMOUNT</th>
                          <th className="px-3 py-2 text-left font-bold border-b">G.W.</th>
                          <th className="px-3 py-2 text-left font-bold border-b">N.W.</th>
                          <th className="px-3 py-2 text-left font-bold border-b">CBM</th>
                          <th className="px-3 py-2 text-left font-bold border-b">Move</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.values(groupedItems).map((data, idx) => {
                          const totalQty = data.count * data.pcsPerCtn;
                          const amount = (data.pricePerPcs || 0) * totalQty;
                          return (
                            <tr key={idx} className="border-b">
                              <td className="px-3 py-2">
                                {data.photo ? (
                                  <img src={data.photo} alt="" className="w-10 h-10 object-cover rounded border" />
                                ) : (
                                  <div className="w-10 h-10 bg-gray-100 rounded border flex items-center justify-center">
                                    <Image className="w-3 h-3 text-gray-300" />
                                  </div>
                                )}
                              </td>
                              <td className="px-3 py-2 font-mono text-xs">{data.markNo}</td>
                              <td className="px-3 py-2 font-medium">{data.description}</td>
                              <td className="px-3 py-2 font-medium">{data.count}</td>
                              <td className="px-3 py-2">{data.pcsPerCtn}</td>
                              <td className="px-3 py-2 font-medium">{totalQty}</td>
                              <td className="px-3 py-2 text-green-700">{data.pricePerPcs ? `¥${data.pricePerPcs.toFixed(2)}` : '-'}</td>
                              <td className="px-3 py-2 font-medium text-green-700">¥{amount.toFixed(2)}</td>
                              <td className="px-3 py-2">{(data.gwPerCtn * data.count).toFixed(2)}</td>
                              <td className="px-3 py-2">{(data.nwPerCtn * data.count).toFixed(2)}</td>
                              <td className="px-3 py-2">{(data.cbmPerCtn * data.count).toFixed(2)}</td>
                              <td className="px-3 py-2">
                                <div className="flex flex-col gap-1">
                                  <div className="flex items-center gap-1">
                                    <input
                                      type="number"
                                      min="1"
                                      max={data.count}
                                      defaultValue={data.count}
                                      className="w-14 px-1 py-0.5 border rounded text-center text-xs"
                                      id={`move-qty-${container.id}-${idx}`}
                                    />
                                    <span className="text-xs text-gray-400">CTN</span>
                                  </div>
                                  <div className="flex items-center gap-1">
                                    <select
                                      className="text-xs border rounded px-1 py-0.5 bg-white"
                                      defaultValue=""
                                      onChange={(e) => {
                                        const targetId = parseInt(e.target.value);
                                        if (!targetId) return;
                                        const qtyEl = document.getElementById(`move-qty-${container.id}-${idx}`);
                                        const qty = parseInt(qtyEl?.value) || data.count;
                                        moveProductBetweenContainers(container.id, data.key, qty, targetId);
                                        e.target.value = '';
                                      }}
                                    >
                                      <option value="">Move to...</option>
                                      {containers.filter(c => c.id !== container.id).map(c => (
                                        <option key={c.id} value={c.id}>Container #{c.id}</option>
                                      ))}
                                    </select>
                                    <button
                                      onClick={() => {
                                        const qtyEl = document.getElementById(`move-qty-${container.id}-${idx}`);
                                        const qty = parseInt(qtyEl?.value) || data.count;
                                        removeProductFromContainer(container.id, data.key, qty);
                                      }}
                                      className="text-red-500 hover:text-red-700 p-0.5"
                                      title="Remove from container"
                                    >
                                      <X className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}
          </div>
        )}

      </div>
    </div>
  );
};

export default ContainerOptimizer;
