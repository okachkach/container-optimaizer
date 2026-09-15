/**
 * Expand cartons and pack by descending weight using the existing sequential algorithm.
 * @param {Array<Object>} shipment Shipment products with carton quantities.
 * @param {number} maxWeight Maximum gross weight per container.
 * @param {number} maxCbm Maximum volume per container.
 * @returns {{containers: Array<Object>, cartonCount: number}} Generated containers and carton count.
 */
export const generateContainers = (shipment, maxWeight, maxCbm) => {
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
  return { containers: newContainers, cartonCount: expandedProducts.length };
};

/**
 * Recalculate totals with the existing per-carton fallback rules.
 * @param {Object} container Container whose items supply totals.
 * @returns {Object} New container object with recalculated totals.
 */
export const recalcContainerTotals = (container) => {
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

/**
 * Move matching cartons without capacity checks and filter empty containers.
 * @param {Array<Object>} prev Current containers.
 * @param {number} fromContainerId Source container ID.
 * @param {string} productKey Mark number and description joined with a pipe.
 * @param {number} ctnCount Requested carton quantity.
 * @param {number} toContainerId Destination container ID.
 * @returns {Array<Object>} Updated containers, or the original array for a no-op.
 */
export const moveProductBetweenContainers = (prev, fromContainerId, productKey, ctnCount, toContainerId) => {
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
};

/**
 * Remove matching cartons from the end and filter empty containers.
 * @param {Array<Object>} prev Current containers.
 * @param {number} containerId Container ID.
 * @param {string} productKey Mark number and description joined with a pipe.
 * @param {number} ctnCount Requested carton quantity.
 * @returns {Array<Object>} Updated containers, or the original array for a no-op.
 */
export const removeProductFromContainer = (prev, containerId, productKey, ctnCount) => {
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
};

/**
 * Add cartons after checking weight, then volume, without deducting shipment quantities.
 * @param {Array<Object>} prev Current containers.
 * @param {number} containerId Destination container ID.
 * @param {Object} shipmentItem Shipment product to expand into cartons.
 * @param {number} qty Quantity already normalized by the caller.
 * @param {number} maxWeight Maximum gross weight.
 * @param {number} maxCbm Maximum volume.
 * @returns {{containers: Array<Object>, error?: string}} Updated array or original array with a capacity error.
 */
export const addContainerProducts = (prev, containerId, shipmentItem, qty, maxWeight, maxCbm) => {
  const updated = prev.map(c => ({ ...c, items: [...c.items] }));
  const idx = updated.findIndex(c => c.id === containerId);
  if (idx === -1) return { containers: prev };

  // Check limits
  const addedWeight = shipmentItem.gwPerCtn * qty;
  const addedCBM = shipmentItem.cbmPerCtn * qty;
  if (updated[idx].totalGW + addedWeight > maxWeight) {
    return { containers: prev, error: 'weight' };
  }
  if (updated[idx].totalCBM + addedCBM > maxCbm) {
    return { containers: prev, error: 'volume' };
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
  return { containers: updated };
};
