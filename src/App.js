import React, { useState, useMemo, useCallback, useEffect, useRef } from 'react';
import { Plus, Download, Image, FileSpreadsheet, Upload, Search, ShoppingCart, Package, ChevronDown, ChevronUp, X, Trash2, ArrowUpDown, Edit3, Undo2, PlusCircle, CheckCircle, AlertTriangle, FileText, Copy, Loader, BarChart3, Clock } from 'lucide-react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { readProjects, saveProject, updateHistory } from './projectStorage';
import { importPackingWorkbook, downloadProject, validateProject } from './packingProject';

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

  // ─── NEW STATE: Sorting ─────────────────────────────────────────
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });

  // ─── NEW STATE: Container deletion ──────────────────────────────
  const [deleteModal, setDeleteModal] = useState({ show: false, containerId: null });
  const [deletedContainers, setDeletedContainers] = useState([]);

  // ─── NEW STATE: Toast notifications ─────────────────────────────
  const [toastMessage, setToastMessage] = useState(null);
  const toastTimeout = useRef(null);

  // ─── NEW STATE: Container renaming ──────────────────────────────
  const [containerNames, setContainerNames] = useState({});
  const [editingContainerName, setEditingContainerName] = useState(null);
  const [tempContainerName, setTempContainerName] = useState('');

  // ─── NEW STATE: Add product to specific container ───────────────
  const [addToContainerModal, setAddToContainerModal] = useState({ show: false, containerId: null });
  const [addToContainerCtn, setAddToContainerCtn] = useState({});

  // ─── NEW STATE: Loading indicator ───────────────────────────────
  const [isGenerating, setIsGenerating] = useState(false);

  // ─── NEW STATE: Photo enlargement modal ─────────────────────────
  const [photoModal, setPhotoModal] = useState({ show: false, src: '', alt: '' });

  // ─── NEW STATE: Clear all confirmation ──────────────────────────
  const [clearAllModal, setClearAllModal] = useState(false);

  // ─── NEW STATE: Shipment search/filter ──────────────────────────
  const [shipmentSearchTerm, setShipmentSearchTerm] = useState('');

  // ─── NEW STATE: Last saved timestamp ────────────────────────────
  const [lastSaved, setLastSaved] = useState(null);
  const [lastSavedDisplay, setLastSavedDisplay] = useState('');
  const [storageReady, setStorageReady] = useState(false);
  const [storageError, setStorageError] = useState('');
  const [history, setHistory] = useState([]);
  const [showHistory, setShowHistory] = useState(false);
  const [projectName, setProjectName] = useState('My shipment');
  const [importing, setImporting] = useState(false);
  const previousContainers = useRef(null);
  const saveQueue = useRef(Promise.resolve());

  const restoreProject = useCallback((data) => {
    setShipment(data.shipment || []);
    setContainers(data.containers || []);
    setContainerNames(data.containerNames || {});
    setCapacity(data.capacity || { maxWeight: 27000, maxCbm: 76 });
    setCatalog(data.catalog || []);
    setCatalogSources(data.catalogSources || []);
    setProjectName(data.projectName || 'My shipment');
    setDeletedContainers([]);
  }, []);

  // ─── TOAST HELPER ───────────────────────────────────────────────
  const showToast = useCallback((message, type = 'success') => {
    if (toastTimeout.current) clearTimeout(toastTimeout.current);
    setToastMessage({ message, type });
    toastTimeout.current = setTimeout(() => setToastMessage(null), 3500);
  }, []);

  // Hydrate before enabling saves so an empty initial render cannot overwrite data.
  useEffect(() => {
    let active = true;
    readProjects().then(({ current, history: savedHistory }) => {
      if (!active) return;
      const legacy = !current && localStorage.getItem('containerOptimizer');
      const data = current || (legacy ? JSON.parse(legacy) : null);
      if (data) restoreProject(data);
      setHistory(savedHistory.sort((a, b) => b.createdAt.localeCompare(a.createdAt)));
      setStorageReady(true);
    }).catch(error => {
      if (active) setStorageError(`Could not open saved projects: ${error.message}. Reload to retry.`);
    });
    return () => { active = false; };
  }, [restoreProject]);

  useEffect(() => {
    if (!storageReady) return;
    const now = new Date();
    const current = { version: 1, shipment, containers, containerNames, capacity, catalog,
      catalogSources, projectName, lastSaved: now.toISOString() };
    const signature = JSON.stringify({ containers, containerNames, capacity });
    const snapshot = containers.length && signature !== previousContainers.current
      ? { id: window.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`, name: projectName, createdAt: now.toISOString(), project: current } : null;
    previousContainers.current = signature;
    setLastSavedDisplay('Saving...');
    saveQueue.current = saveQueue.current.catch(() => {}).then(() => saveProject(current, snapshot)).then(() => {
      if (snapshot) setHistory(prev => [snapshot, ...prev]);
      setLastSaved(now);
      setStorageError('');
    }).catch(error => {
      previousContainers.current = null;
      setStorageError(`Changes could not be saved: ${error.message}. Download a backup before closing.`);
    });
  }, [storageReady, shipment, containers, containerNames, capacity, catalog, catalogSources, projectName]);

  const importProject = async (event) => {
    const file = event.target.files[0];
    event.target.value = '';
    if (!file) return;
    setImporting(true);
    try {
      if (file.size > 20 * 1024 * 1024) throw new Error('Please use a file smaller than 20 MB.');
      const data = file.name.toLowerCase().endsWith('.json')
        ? validateProject(JSON.parse(await file.text()))
        : await importPackingWorkbook(await file.arrayBuffer());
      await saveQueue.current;
      restoreProject({ ...data, projectName: data.projectName || file.name.replace(/\.[^.]+$/, '') });
      showToast('Project opened. Excel totals were recalculated from the edited rows.');
    } catch (error) { showToast(error.message, 'error'); }
    finally { setImporting(false); }
  };

  const changeHistory = async (entry, remove = false) => {
    if (remove && !window.confirm(`Delete saved history entry "${entry.name}"?`)) return;
    const name = remove ? entry.name : window.prompt('History entry name', entry.name);
    if (!name?.trim()) return;
    const updated = { ...entry, name: name.trim() };
    try {
      await updateHistory(updated, remove);
      setHistory(prev => remove ? prev.filter(x => x.id !== entry.id) : prev.map(x => x.id === entry.id ? updated : x));
    } catch (error) { showToast(`History could not be updated: ${error.message}`, 'error'); }
  };

  // ─── LAST SAVED DISPLAY TIMER ──────────────────────────────────
  useEffect(() => {
    const updateDisplay = () => {
      if (!lastSaved) { setLastSavedDisplay(''); return; }
      const diff = Math.floor((Date.now() - lastSaved.getTime()) / 1000);
      if (diff < 5) setLastSavedDisplay('Just now');
      else if (diff < 60) setLastSavedDisplay(`${diff}s ago`);
      else if (diff < 3600) setLastSavedDisplay(`${Math.floor(diff / 60)}m ago`);
      else setLastSavedDisplay(`${Math.floor(diff / 3600)}h ago`);
    };
    updateDisplay();
    const interval = setInterval(updateDisplay, 10000);
    return () => clearInterval(interval);
  }, [lastSaved]);

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
    showToast(`Added ${qty} CTN of "${catalogProduct.markNo}" to shipment`);
  };

  const updateShipmentCtn = (id, newCtn) => {
    // Allow empty field while user is typing — don't remove the product
    if (newCtn === '' || newCtn === null || newCtn === undefined) {
      setShipment(shipment.map(s =>
        s.id === id ? { ...s, ctn: '' } : s
      ));
      return;
    }
    const ctn = parseInt(newCtn);
    if (isNaN(ctn) || ctn <= 0) return;
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
    showToast('Product removed from shipment');
  };

  // ─── DUPLICATE PRODUCT IN SHIPMENT ──────────────────────────────
  const duplicateShipmentProduct = (item) => {
    // Auto-increment: try to parse markNo and bump suffix
    let newMarkNo = item.markNo;
    const match = item.markNo.match(/^(.+?)(\d+)$/);
    if (match) {
      const prefix = match[1];
      const num = parseInt(match[2]) + 1;
      newMarkNo = `${prefix}${num}`;
    } else {
      newMarkNo = item.markNo + '-copy';
    }
    const dup = {
      ...item,
      id: Date.now() + Math.random(),
      catalogId: `dup-${Date.now()}`,
      markNo: newMarkNo
    };
    setShipment(prev => [...prev, dup]);
    showToast(`Duplicated "${item.markNo}" → "${newMarkNo}"`);
  };

  // ─── CLEAR ALL ──────────────────────────────────────────────────
  const executeClearAll = () => {
    setShipment([]);
    setContainers([]);
    setClearAllModal(false);
    showToast('All shipment products and containers cleared');
  };

  // ─── COPY CONTAINER INFO ───────────────────────────────────────
  const copyContainerInfo = async (container) => {
    const name = containerNames[container.id] || `Container #${container.id}`;
    const weightPct = maxWeight > 0 ? (container.totalGW / maxWeight * 100).toFixed(1) : '0';
    const cbmPct = maxCbm > 0 ? (container.totalCBM / maxCbm * 100).toFixed(1) : '0';
    const text = [
      `📦 ${name}`,
      `Cartons: ${container.totalCartons}`,
      `T/QTY: ${container.totalQty || 0}`,
      `G.W.: ${container.totalGW.toFixed(2)} kg (${weightPct}%)`,
      `N.W.: ${container.totalNW.toFixed(2)} kg`,
      `CBM: ${container.totalCBM.toFixed(2)} m³ (${cbmPct}%)`,
      `Total Price: ¥${(container.totalPrice || 0).toFixed(2)}`,
      `Products: ${Object.keys((() => { const g = {}; container.items.forEach(i => { g[i.markNo + '|' + i.description] = true; }); return g; })()).length} SKUs`
    ].join('\\n');
    try {
      await navigator.clipboard.writeText(text);
      showToast('Container info copied to clipboard!');
    } catch {
      showToast('Failed to copy to clipboard', 'error');
    }
  };

  // ─── CONTAINER COLOR BORDER ─────────────────────────────────────
  const getContainerBorderColor = (container) => {
    const weightPct = maxWeight > 0 ? (container.totalGW / maxWeight * 100) : 0;
    const cbmPct = maxCbm > 0 ? (container.totalCBM / maxCbm * 100) : 0;
    const maxPct = Math.max(weightPct, cbmPct);
    if (maxPct > 100) return 'border-blue-500';
    if (maxPct >= 95) return 'border-yellow-500';
    if (maxPct >= 80) return 'border-green-500';
    if (maxPct >= 70) return 'border-slate-300';
    return 'border-red-400';
  };

  // ─── SORTING ────────────────────────────────────────────────────
  const handleSort = (key) => {
    setSortConfig(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  const sortedShipment = useMemo(() => {
    if (!sortConfig.key) return shipment;
    const sorted = [...shipment].sort((a, b) => {
      let valA, valB;
      switch (sortConfig.key) {
        case 'markNo':
          valA = a.markNo.toLowerCase();
          valB = b.markNo.toLowerCase();
          return sortConfig.direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
        case 'description':
          valA = a.description.toLowerCase();
          valB = b.description.toLowerCase();
          return sortConfig.direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
        case 'ctn':
          valA = a.ctn; valB = b.ctn; break;
        case 'pcsPerCtn':
          valA = a.pcsPerCtn; valB = b.pcsPerCtn; break;
        case 'totalQty':
          valA = a.totalQty; valB = b.totalQty; break;
        case 'totalGW':
          valA = a.totalGW; valB = b.totalGW; break;
        case 'totalNW':
          valA = a.totalNW; valB = b.totalNW; break;
        case 'totalCBM':
          valA = a.totalCBM; valB = b.totalCBM; break;
        default: return 0;
      }
      if (valA === undefined) return 0;
      return sortConfig.direction === 'asc' ? valA - valB : valB - valA;
    });
    return sorted;
  }, [shipment, sortConfig]);

  // ─── SHIPMENT SEARCH FILTER ─────────────────────────────────────
  const filteredSortedShipment = useMemo(() => {
    let list = sortedShipment;
    if (shipmentSearchTerm.trim()) {
      const term = shipmentSearchTerm.toLowerCase();
      list = list.filter(s =>
        s.markNo.toLowerCase().includes(term) ||
        s.description.toLowerCase().includes(term)
      );
    }
    return list;
  }, [sortedShipment, shipmentSearchTerm]);

  const SortHeader = ({ label, sortKey, className = '' }) => {
    const isActive = sortConfig.key === sortKey;
    return (
      <th
        className={`px-3 py-2 text-left font-bold border-b cursor-pointer hover:bg-green-100 select-none transition ${className} ${isActive ? 'bg-green-100 text-green-800' : ''}`}
        onClick={() => handleSort(sortKey)}
      >
        <div className="flex items-center gap-1">
          {label}
          {isActive ? (
            sortConfig.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
          ) : (
            <ArrowUpDown className="w-3 h-3 text-gray-400" />
          )}
        </div>
      </th>
    );
  };

  // ─── DELETE CONTAINER ───────────────────────────────────────────
  const confirmDeleteContainer = (containerId) => {
    setDeleteModal({ show: true, containerId });
  };

  const executeDeleteContainer = () => {
    const { containerId } = deleteModal;
    const containerToDelete = containers.find(c => c.id === containerId);
    if (containerToDelete) {
      setDeletedContainers(prev => [...prev, containerToDelete]);
      setContainers(prev => prev.filter(c => c.id !== containerId));
      showToast(`Container #${containerNames[containerId] || containerId} deleted successfully. Products are marked as shipped.`);
    }
    setDeleteModal({ show: false, containerId: null });
  };

  const undoDeleteContainer = () => {
    if (deletedContainers.length === 0) return;
    const last = deletedContainers[deletedContainers.length - 1];
    setDeletedContainers(prev => prev.slice(0, -1));
    setContainers(prev => [...prev, last].sort((a, b) => a.id - b.id));
    showToast(`Container #${containerNames[last.id] || last.id} restored.`);
  };

  // ─── ADD EMPTY CONTAINER ────────────────────────────────────────
  const addEmptyContainer = () => {
    const maxId = containers.length > 0 ? Math.max(...containers.map(c => c.id)) : 0;
    const newContainer = {
      id: maxId + 1,
      items: [],
      totalGW: 0, totalNW: 0, totalCBM: 0, totalCartons: 0, totalPrice: 0, totalQty: 0
    };
    setContainers(prev => [...prev, newContainer]);
    showToast(`Empty Container #${maxId + 1} created. Add products from the shipment list.`);
  };

  // ─── ADD PRODUCT TO CONTAINER (from shipment) ──────────────────
  const addProductToContainer = (containerId, shipmentItem, ctnCount) => {
    const qty = Math.min(parseInt(ctnCount) || 1, shipmentItem.ctn);
    if (qty <= 0) return;

    setContainers(prev => {
      const updated = prev.map(c => ({ ...c, items: [...c.items] }));
      const idx = updated.findIndex(c => c.id === containerId);
      if (idx === -1) return prev;

      // Check limits
      const addedWeight = shipmentItem.gwPerCtn * qty;
      const addedCBM = shipmentItem.cbmPerCtn * qty;
      if (updated[idx].totalGW + addedWeight > maxWeight) {
        showToast(`Cannot add: would exceed weight limit (${maxWeight} kg)`, 'error');
        return prev;
      }
      if (updated[idx].totalCBM + addedCBM > maxCbm) {
        showToast(`Cannot add: would exceed volume limit (${maxCbm} m³)`, 'error');
        return prev;
      }

      // Add individual cartons as items
      for (let i = 0; i < qty; i++) {
        updated[idx].items.push({
          ...shipmentItem,
          singleCtnGW: shipmentItem.gwPerCtn,
          singleCtnNW: shipmentItem.nwPerCtn,
          singleCtnCBM: shipmentItem.cbmPerCtn
        });
      }
      updated[idx] = recalcContainerTotals(updated[idx]);
      return updated;
    });
    showToast(`Added ${qty} CTN of ${shipmentItem.markNo} to Container #${containerNames[containerId] || containerId}`);
  };

  // ─── CONTAINER RENAMING ─────────────────────────────────────────
  const startRenamingContainer = (containerId) => {
    setEditingContainerName(containerId);
    setTempContainerName(containerNames[containerId] || `Container ${containerId}`);
  };

  const saveContainerName = (containerId) => {
    setContainerNames(prev => ({ ...prev, [containerId]: tempContainerName.trim() || `Container ${containerId}` }));
    setEditingContainerName(null);
    showToast('Container renamed successfully');
  };

  // ─── EXPORT SINGLE CONTAINER CSV ───────────────────────────────
  const exportContainerCSV = useCallback((container) => {
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

    const csvContent = rows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `container_${container.id}_packing_list.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
    showToast(`CSV exported for Container #${containerNames[container.id] || container.id}`);
  }, [containerNames, showToast]);

  // ─── EXPORT ALL CSV ─────────────────────────────────────────────
  const exportAllCSV = useCallback(() => {
    containers.forEach(container => exportContainerCSV(container));
    showToast(`Exported CSV for all ${containers.length} containers`);
  }, [containers, exportContainerCSV, showToast]);

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
    showToast(`Product "${catProduct.description}" added to catalog`);
  };

  // ─── CONTAINER OPTIMIZATION ─────────────────────────────────────
  const optimizeContainers = () => {
    setIsGenerating(true);
    // Use setTimeout to allow the loading UI to render before crunching
    setTimeout(() => {
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
      setIsGenerating(false);
      showToast(`Generated ${newContainers.length} container(s) with ${expandedProducts.length} cartons`);
    }, 100);
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
        {storageError && <p role="alert" className="bg-red-100 text-red-800 p-3 rounded mb-3">{storageError}</p>}
        {!storageReady && <p role="status" className="p-3">Opening saved workspace...</p>}
        <fieldset disabled={!storageReady || importing} className="min-w-0">
        <section className="bg-white rounded-xl shadow p-4 mb-4">
          <div className="flex flex-wrap gap-3 items-center">
            <label className="text-sm font-medium">Project name
              <input className="border rounded px-2 py-1 ml-2" value={projectName} onChange={e => setProjectName(e.target.value)} />
            </label>
            <button type="button" className="border rounded px-3 py-2 text-sm" onClick={() => setShowHistory(!showHistory)}>History ({history.length})</button>
            <label className="border rounded px-3 py-2 text-sm cursor-pointer">{importing ? 'Opening...' : 'Open Excel / Backup'}
              <input type="file" accept=".xlsx,.json" className="hidden" onChange={importProject} />
            </label>
            <button type="button" className="border rounded px-3 py-2 text-sm" onClick={() => downloadProject({ shipment, containers, containerNames, capacity, catalog, catalogSources, projectName }, projectName)}>Download Backup</button>
          </div>
          <p className="text-xs text-gray-500 mt-2">History is saved in this browser. Download backups for another computer or before clearing browser data. Edit exported Excel rows, save as XLSX, then use Open Excel / Backup to reopen them.</p>
          {showHistory && <div className="mt-4 max-h-80 overflow-y-auto">
            {!history.length && <p className="text-sm text-gray-500">Your container history will appear here when you create containers.</p>}
            {history.map(entry => <div key={entry.id} className="border-t py-3 flex flex-wrap items-center gap-2 text-sm">
              <div className="flex-1 min-w-48"><strong>{entry.name}</strong><p className="text-xs text-gray-500">{new Date(entry.createdAt).toLocaleString()} · {entry.project.containers.length} containers · {entry.project.containers.reduce((n, c) => n + c.totalCartons, 0)} cartons</p></div>
              <button className="text-blue-700 px-2" onClick={async () => { try { await saveQueue.current; restoreProject({ ...entry.project, projectName: entry.name }); showToast('Saved project opened as a new working copy'); } catch { showToast('Save failed. Download a backup before opening history.', 'error'); } }}>Open / Duplicate</button>
              <button className="px-2" onClick={() => downloadProject(entry.project, entry.name)}>Backup</button>
              <button className="px-2" onClick={() => changeHistory(entry)}>Rename</button>
              <button className="text-red-600 px-2" onClick={() => changeHistory(entry, true)}>Delete</button>
            </div>)}
          </div>}
        </section>

        {/* HEADER */}
        <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
          <div className="flex items-center gap-3 mb-2">
            <FileSpreadsheet className="w-8 h-8 text-blue-600" />
            <h1 className="text-2xl font-bold text-gray-800">Container Packing List Generator</h1>
            {lastSavedDisplay && (
              <span className="ml-auto flex items-center gap-1.5 text-xs text-gray-400">
                <Clock className="w-3.5 h-3.5" /> Saved: {lastSavedDisplay}
              </span>
            )}
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

        {/* STATS DASHBOARD */}
        {shipment.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg p-4 mb-4">
            <div className="flex items-center gap-2 mb-3">
              <BarChart3 className="w-5 h-5 text-blue-600" />
              <h2 className="text-sm font-bold text-gray-700 uppercase tracking-wide">Shipment Overview</h2>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3 text-sm">
              <div className="bg-blue-50 p-3 rounded-lg border border-blue-200">
                <p className="text-xs text-gray-500">Products</p>
                <p className="text-xl font-bold text-blue-700">{shipment.length}</p>
              </div>
              <div className="bg-slate-50 p-3 rounded-lg border">
                <p className="text-xs text-gray-500">Total Cartons</p>
                <p className="text-xl font-bold text-slate-800">{shipmentTotals.ctn}</p>
              </div>
              <div className="bg-green-50 p-3 rounded-lg border border-green-200">
                <p className="text-xs text-gray-500">Total G.W.</p>
                <p className="text-xl font-bold text-green-700">{shipmentTotals.gw.toFixed(1)} kg</p>
              </div>
              <div className="bg-purple-50 p-3 rounded-lg border border-purple-200">
                <p className="text-xs text-gray-500">Total N.W.</p>
                <p className="text-xl font-bold text-purple-700">{shipmentTotals.nw.toFixed(1)} kg</p>
              </div>
              <div className="bg-orange-50 p-3 rounded-lg border border-orange-200">
                <p className="text-xs text-gray-500">Total CBM</p>
                <p className="text-xl font-bold text-orange-700">{shipmentTotals.cbm.toFixed(2)} m³</p>
              </div>
              <div className="bg-yellow-50 p-3 rounded-lg border border-yellow-200">
                <p className="text-xs text-gray-500">Total Value</p>
                <p className="text-xl font-bold text-yellow-700">¥{shipmentTotals.price.toFixed(0)}</p>
              </div>
              <div className="bg-indigo-50 p-3 rounded-lg border border-indigo-200">
                <p className="text-xs text-gray-500">Est. Containers</p>
                <p className="text-xl font-bold text-indigo-700">
                  ~{Math.ceil(Math.max(shipmentTotals.gw / (maxWeight || 1), shipmentTotals.cbm / (maxCbm || 1)))}
                </p>
              </div>
            </div>
          </div>
        )}

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
                                <img src={product.photo} alt="" className="w-12 h-12 object-cover rounded border cursor-pointer hover:opacity-80 transition"
                                  onClick={() => setPhotoModal({ show: true, src: product.photo, alt: product.description })} />
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
            <div className="flex items-center justify-between mb-4 flex-wrap gap-2">
              <h2 className="text-xl font-bold text-gray-800 flex items-center gap-2">
                <ShoppingCart className="w-6 h-6 text-green-600" />
                Shipment List
                <span className="bg-green-600 text-white text-xs font-bold px-2.5 py-0.5 rounded-full">{shipment.length} Products</span>
                <span className="text-sm font-normal text-gray-500">({shipmentTotals.ctn} CTN)</span>
              </h2>
              <div className="flex items-center gap-2">
                <div className="flex gap-2 text-xs flex-wrap">
                  <span className="bg-green-100 text-green-800 px-2 py-1 rounded">G.W: {shipmentTotals.gw.toFixed(1)} kg</span>
                  <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded">CBM: {shipmentTotals.cbm.toFixed(2)} m³</span>
                  <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded">¥{shipmentTotals.price.toFixed(2)}</span>
                  <span className="bg-orange-100 text-orange-800 px-2 py-1 rounded">
                    ~{Math.ceil(Math.max(shipmentTotals.gw / (maxWeight || 1), shipmentTotals.cbm / (maxCbm || 1)))} container(s)
                  </span>
                </div>
                <button onClick={() => setClearAllModal(true)}
                  className="bg-red-50 text-red-600 px-3 py-1.5 rounded border border-red-200 hover:bg-red-100 transition text-xs font-medium flex items-center gap-1.5"
                  title="Clear all products and containers">
                  <Trash2 className="w-3.5 h-3.5" /> Clear All
                </button>
              </div>
            </div>

            {/* Shipment Search Bar */}
            <div className="relative mb-3">
              <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
              <input
                type="text"
                placeholder="Filter shipment by MARKS&NO or description..."
                value={shipmentSearchTerm}
                onChange={(e) => setShipmentSearchTerm(e.target.value)}
                className="w-full pl-10 pr-20 py-2 border rounded-lg focus:ring-2 focus:ring-green-500 focus:border-transparent text-sm"
              />
              {shipmentSearchTerm && (
                <div className="absolute right-3 top-1/2 -translate-y-1/2 flex items-center gap-2">
                  <span className="text-xs text-gray-500">{filteredSortedShipment.length} of {shipment.length}</span>
                  <button onClick={() => setShipmentSearchTerm('')} className="text-gray-400 hover:text-gray-600">
                    <X className="w-4 h-4" />
                  </button>
                </div>
              )}
            </div>

            <div className="border rounded-lg overflow-x-auto mb-4">
              <table className="w-full text-sm">
                <thead className="bg-green-50">
                  <tr>
                    <th className="px-2 py-2 text-center font-bold border-b w-10">#</th>
                    <th className="px-3 py-2 text-left font-bold border-b">Photo</th>
                    <SortHeader label="Mark No" sortKey="markNo" />
                    <SortHeader label="Description" sortKey="description" />
                    <SortHeader label="CTN" sortKey="ctn" />
                    <SortHeader label="PCS/CTN" sortKey="pcsPerCtn" />
                    <SortHeader label="T/QTY" sortKey="totalQty" />
                    <th className="px-3 py-2 text-left font-bold border-b">U/Price</th>
                    <th className="px-3 py-2 text-left font-bold border-b">Amount</th>
                    <SortHeader label="G.W." sortKey="totalGW" />
                    <SortHeader label="N.W." sortKey="totalNW" />
                    <SortHeader label="CBM" sortKey="totalCBM" />
                    <th className="px-3 py-2 text-left font-bold border-b">Source</th>
                    <th className="px-3 py-2 text-center font-bold border-b">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSortedShipment.map((item, rowIndex) => (
                    <tr key={item.id} className="border-b hover:bg-gray-50">
                      <td className="px-2 py-2 text-center text-xs text-gray-400 font-mono">{rowIndex + 1}</td>
                      <td className="px-3 py-2">
                        {item.photo ? (
                          <img src={item.photo} alt="" className="w-10 h-10 object-cover rounded border cursor-pointer hover:opacity-80 transition"
                            onClick={() => setPhotoModal({ show: true, src: item.photo, alt: item.description })} />
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
                        <div className="flex items-center gap-1 justify-center">
                          <button onClick={() => duplicateShipmentProduct(item)} className="text-blue-500 hover:text-blue-700 p-0.5" title="Duplicate product">
                            <Copy className="w-4 h-4" />
                          </button>
                          <button onClick={() => removeFromShipment(item.id)} className="text-red-500 hover:text-red-700 p-0.5" title="Remove product">
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <button
              onClick={optimizeContainers}
              disabled={isGenerating}
              className={`text-white px-8 py-3 rounded-lg transition font-semibold text-lg w-full flex items-center justify-center gap-3 ${
                isGenerating ? 'bg-gray-400 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700'
              }`}
            >
              {isGenerating ? (
                <>
                  <Loader className="w-5 h-5 animate-spin" /> Generating Loading Plan...
                </>
              ) : (
                'Generate Container Loading Plan'
              )}
            </button>
          </div>
        )}

        {/* CONTAINER RESULTS */}
        {containers.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg p-6 mb-4">
            <div className="flex flex-wrap justify-between items-center mb-4 gap-2">
              <h2 className="text-xl font-bold text-gray-800">
                Packing List: {containers.length} Container{containers.length > 1 ? 's' : ''}
              </h2>
              <div className="flex flex-wrap gap-2">
                {deletedContainers.length > 0 && (
                  <button onClick={undoDeleteContainer}
                    className="bg-yellow-500 text-white px-4 py-2 rounded hover:bg-yellow-600 transition flex items-center gap-2 text-sm">
                    <Undo2 className="w-4 h-4" /> Undo Delete
                  </button>
                )}
                <button onClick={addEmptyContainer}
                  className="bg-green-600 text-white px-4 py-2 rounded hover:bg-green-700 transition flex items-center gap-2 text-sm">
                  <PlusCircle className="w-4 h-4" /> Add Empty Container
                </button>
                <button onClick={exportAllCSV}
                  className="bg-teal-600 text-white px-4 py-2 rounded hover:bg-teal-700 transition flex items-center gap-2 text-sm">
                  <FileText className="w-4 h-4" /> Export All CSV
                </button>
                <button onClick={exportToXlsx}
                  className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700 transition flex items-center gap-2 text-sm">
                  <Download className="w-4 h-4" /> Export All XLSX
                </button>
              </div>
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
                <div key={container.id} className={`border-2 rounded-lg p-5 mb-4 shadow-sm ${getContainerBorderColor(container)}`}>
                  <div className="flex items-center justify-between bg-slate-100 p-3 rounded mb-3 gap-2 flex-wrap">
                    <div className="flex items-center gap-2">
                      {editingContainerName === container.id ? (
                        <div className="flex items-center gap-2">
                          <input
                            type="text"
                            value={tempContainerName}
                            onChange={(e) => setTempContainerName(e.target.value)}
                            onKeyDown={(e) => { if (e.key === 'Enter') saveContainerName(container.id); if (e.key === 'Escape') setEditingContainerName(null); }}
                            className="px-2 py-1 border rounded text-sm font-bold"
                            autoFocus
                          />
                          <button onClick={() => saveContainerName(container.id)} className="text-green-600 hover:text-green-800">
                            <CheckCircle className="w-4 h-4" />
                          </button>
                          <button onClick={() => setEditingContainerName(null)} className="text-gray-400 hover:text-gray-600">
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      ) : (
                        <>
                          <h3 className="text-lg font-bold text-gray-800">
                            {containerNames[container.id] || `CONTAINER #${container.id}`}
                          </h3>
                          <button onClick={() => startRenamingContainer(container.id)}
                            className="text-gray-400 hover:text-blue-600 transition" title="Rename container">
                            <Edit3 className="w-4 h-4" />
                          </button>
                          {/* Utilization badge */}
                          {(() => {
                            const mp = Math.max(weightPercent, cbmPercent);
                            const label = mp > 100 ? 'Over' : mp >= 95 ? 'Near Full' : mp >= 80 ? 'Optimal' : mp >= 70 ? 'Moderate' : 'Low';
                            const cls = mp > 100 ? 'bg-blue-100 text-blue-700' : mp >= 95 ? 'bg-yellow-100 text-yellow-700' : mp >= 80 ? 'bg-green-100 text-green-700' : mp >= 70 ? 'bg-slate-100 text-slate-600' : 'bg-red-100 text-red-600';
                            return <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${cls}`}>{mp.toFixed(0)}% — {label}</span>;
                          })()}
                        </>
                      )}
                    </div>
                    <div className="flex items-center gap-2 flex-wrap">
                      <button
                        onClick={() => setAddToContainerModal({ show: true, containerId: container.id })}
                        className="bg-green-600 text-white px-3 py-1.5 rounded hover:bg-green-700 transition flex items-center gap-1.5 text-sm"
                      >
                        <Plus className="w-3.5 h-3.5" /> Add Products
                      </button>
                      <button
                        onClick={() => copyContainerInfo(container)}
                        className="bg-gray-600 text-white px-3 py-1.5 rounded hover:bg-gray-700 transition flex items-center gap-1.5 text-sm"
                      >
                        <Copy className="w-3.5 h-3.5" /> Copy Info
                      </button>
                      <button
                        onClick={() => exportContainerCSV(container)}
                        className="bg-teal-600 text-white px-3 py-1.5 rounded hover:bg-teal-700 transition flex items-center gap-1.5 text-sm"
                      >
                        <FileText className="w-3.5 h-3.5" /> CSV
                      </button>
                      <button
                        onClick={() => exportSingleContainer(container)}
                        className="bg-indigo-600 text-white px-3 py-1.5 rounded hover:bg-indigo-700 transition flex items-center gap-1.5 text-sm"
                      >
                        <Download className="w-3.5 h-3.5" /> XLSX
                      </button>
                      <button
                        onClick={() => confirmDeleteContainer(container.id)}
                        className="bg-red-600 text-white px-3 py-1.5 rounded hover:bg-red-700 transition flex items-center gap-1.5 text-sm"
                      >
                        <Trash2 className="w-3.5 h-3.5" /> Delete
                      </button>
                    </div>
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
                                  <img src={data.photo} alt="" className="w-10 h-10 object-cover rounded border cursor-pointer hover:opacity-80 transition"
                                    onClick={() => setPhotoModal({ show: true, src: data.photo, alt: data.description })} />
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

        {/* SHOW ADD EMPTY CONTAINER BUTTON WHEN NO CONTAINERS */}
        {containers.length === 0 && shipment.length > 0 && (
          <div className="text-center py-4">
            <button onClick={addEmptyContainer}
              className="bg-green-600 text-white px-6 py-3 rounded-lg hover:bg-green-700 transition font-semibold text-sm inline-flex items-center gap-2">
              <PlusCircle className="w-5 h-5" /> Create Empty Container Manually
            </button>
          </div>
        )}

        {/* ── TOAST NOTIFICATION ── */}
        {toastMessage && (
          <div className={`fixed bottom-6 right-6 z-50 flex items-center gap-3 px-5 py-3 rounded-lg shadow-xl text-white text-sm font-medium transition-all animate-fade-in ${
            toastMessage.type === 'error' ? 'bg-red-600' : 'bg-green-600'
          }`}>
            {toastMessage.type === 'error' ? (
              <AlertTriangle className="w-5 h-5 flex-shrink-0" />
            ) : (
              <CheckCircle className="w-5 h-5 flex-shrink-0" />
            )}
            {toastMessage.message}
            <button onClick={() => setToastMessage(null)} className="ml-2 hover:opacity-80">
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* ── DELETE CONFIRMATION MODAL ── */}
        {deleteModal.show && (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setDeleteModal({ show: false, containerId: null })}>
            <div className="bg-white rounded-xl shadow-2xl p-6 max-w-md w-full mx-4" onClick={e => e.stopPropagation()}>
              <div className="flex items-center gap-3 mb-4">
                <div className="w-10 h-10 bg-red-100 rounded-full flex items-center justify-center">
                  <Trash2 className="w-5 h-5 text-red-600" />
                </div>
                <h3 className="text-lg font-bold text-gray-800">Delete Container</h3>
              </div>
              <p className="text-gray-600 mb-2">
                Are you sure you want to delete <strong>{containerNames[deleteModal.containerId] || `Container #${deleteModal.containerId}`}</strong>?
              </p>
              <p className="text-sm text-gray-500 mb-6">
                The products in this container will be considered as shipped and will <strong>not</strong> return to the shipment list. You can undo this action.
              </p>
              <div className="flex justify-end gap-3">
                <button onClick={() => setDeleteModal({ show: false, containerId: null })}
                  className="px-4 py-2 rounded border text-gray-700 hover:bg-gray-50 transition text-sm font-medium">
                  Cancel
                </button>
                <button onClick={executeDeleteContainer}
                  className="px-4 py-2 rounded bg-red-600 text-white hover:bg-red-700 transition text-sm font-medium flex items-center gap-2">
                  <Trash2 className="w-4 h-4" /> Delete Container
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── ADD PRODUCTS TO CONTAINER MODAL ── */}
        {addToContainerModal.show && (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setAddToContainerModal({ show: false, containerId: null }); setAddToContainerCtn({}); }}>
            <div className="bg-white rounded-xl shadow-2xl p-6 max-w-4xl w-full mx-4 max-h-[80vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
              <div className="flex items-center justify-between mb-4">
                <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                  <Plus className="w-5 h-5 text-green-600" />
                  Add Products to {containerNames[addToContainerModal.containerId] || `Container #${addToContainerModal.containerId}`}
                </h3>
                <button onClick={() => { setAddToContainerModal({ show: false, containerId: null }); setAddToContainerCtn({}); }}
                  className="text-gray-400 hover:text-gray-600">
                  <X className="w-5 h-5" />
                </button>
              </div>
              {(() => {
                const targetContainer = containers.find(c => c.id === addToContainerModal.containerId);
                const remainWeight = maxWeight - (targetContainer?.totalGW || 0);
                const remainCBM = maxCbm - (targetContainer?.totalCBM || 0);
                return (
                  <div className="grid grid-cols-2 gap-3 mb-4 text-sm">
                    <div className="bg-green-50 p-2 rounded border border-green-200">
                      <span className="text-gray-600">Remaining Weight: </span>
                      <span className="font-bold text-green-700">{remainWeight.toFixed(2)} kg</span>
                    </div>
                    <div className="bg-blue-50 p-2 rounded border border-blue-200">
                      <span className="text-gray-600">Remaining Volume: </span>
                      <span className="font-bold text-blue-700">{remainCBM.toFixed(4)} m³</span>
                    </div>
                  </div>
                );
              })()}
              {shipment.length > 0 ? (
                <div className="border rounded-lg overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead className="bg-gray-50">
                      <tr>
                        <th className="px-3 py-2 text-left font-bold border-b">Mark No</th>
                        <th className="px-3 py-2 text-left font-bold border-b">Description</th>
                        <th className="px-3 py-2 text-left font-bold border-b">In Shipment</th>
                        <th className="px-3 py-2 text-left font-bold border-b">G.W./CTN</th>
                        <th className="px-3 py-2 text-left font-bold border-b">CBM/CTN</th>
                        <th className="px-3 py-2 text-left font-bold border-b">Add</th>
                      </tr>
                    </thead>
                    <tbody>
                      {shipment.map(item => (
                        <tr key={item.id} className="border-b hover:bg-blue-50 transition">
                          <td className="px-3 py-2 font-mono text-xs">{item.markNo}</td>
                          <td className="px-3 py-2 font-medium max-w-48 truncate">{item.description}</td>
                          <td className="px-3 py-2">{item.ctn} CTN</td>
                          <td className="px-3 py-2">{item.gwPerCtn} kg</td>
                          <td className="px-3 py-2">{item.cbmPerCtn} m³</td>
                          <td className="px-3 py-2">
                            <div className="flex items-center gap-1">
                              <input
                                type="number"
                                min="1"
                                max={item.ctn}
                                placeholder="CTN"
                                value={addToContainerCtn[item.id] || ''}
                                onChange={(e) => setAddToContainerCtn(prev => ({ ...prev, [item.id]: e.target.value }))}
                                className="w-16 px-2 py-1 border rounded text-center text-sm"
                              />
                              <button
                                onClick={() => {
                                  addProductToContainer(addToContainerModal.containerId, item, addToContainerCtn[item.id] || 1);
                                  setAddToContainerCtn(prev => ({ ...prev, [item.id]: '' }));
                                }}
                                className="bg-green-600 text-white px-2 py-1 rounded hover:bg-green-700 transition text-xs"
                              >
                                <Plus className="w-3 h-3" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                <p className="text-center py-6 text-gray-400">No products in shipment list. Add products to shipment first.</p>
              )}
            </div>
          </div>
        )}

        {/* ── PHOTO ENLARGEMENT MODAL ── */}
        {photoModal.show && (
          <div className="fixed inset-0 bg-black/70 flex items-center justify-center z-50 p-8" onClick={() => setPhotoModal({ show: false, src: '', alt: '' })}>
            <div className="relative max-w-3xl max-h-[85vh]" onClick={e => e.stopPropagation()}>
              <button onClick={() => setPhotoModal({ show: false, src: '', alt: '' })}
                className="absolute -top-3 -right-3 bg-white rounded-full p-1.5 shadow-lg text-gray-600 hover:text-gray-900 z-10">
                <X className="w-5 h-5" />
              </button>
              <img src={photoModal.src} alt={photoModal.alt} className="max-w-full max-h-[85vh] rounded-lg shadow-2xl object-contain bg-white" />
              {photoModal.alt && (
                <p className="text-center text-white text-sm mt-3 opacity-80">{photoModal.alt}</p>
              )}
            </div>
          </div>
        )}

        {/* ── CLEAR ALL CONFIRMATION MODAL ── */}
        {clearAllModal && (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setClearAllModal(false)}>
            <div className="bg-white rounded-xl shadow-2xl p-6 max-w-md w-full mx-4" onClick={e => e.stopPropagation()}>
              <div className="flex items-center gap-3 mb-4">
                <div className="w-10 h-10 bg-red-100 rounded-full flex items-center justify-center">
                  <Trash2 className="w-5 h-5 text-red-600" />
                </div>
                <h3 className="text-lg font-bold text-gray-800">Clear All Products</h3>
              </div>
              <p className="text-gray-600 mb-2">
                Are you sure you want to remove <strong>all products</strong> from the shipment and <strong>all containers</strong>?
              </p>
              <p className="text-sm text-red-500 mb-6 font-medium">
                This action cannot be undone.
              </p>
              <div className="flex justify-end gap-3">
                <button onClick={() => setClearAllModal(false)}
                  className="px-4 py-2 rounded border text-gray-700 hover:bg-gray-50 transition text-sm font-medium">
                  Cancel
                </button>
                <button onClick={executeClearAll}
                  className="px-4 py-2 rounded bg-red-600 text-white hover:bg-red-700 transition text-sm font-medium flex items-center gap-2">
                  <Trash2 className="w-4 h-4" /> Clear Everything
                </button>
              </div>
            </div>
          </div>
        )}

        </fieldset>
      </div>
    </div>
  );
};

export default ContainerOptimizer;
