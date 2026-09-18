// Presentation only: no assignments, product fields, or project totals are changed.
export const productKey = item => JSON.stringify([item.markNo || '', item.description || '']);

export function productColor(key) {
  let hash = 2166136261;
  for (let i = 0; i < key.length; i++) hash = Math.imul(hash ^ key.charCodeAt(i), 16777619);
  return `hsl(${(hash >>> 0) % 360}, 65%, 55%)`;
}

/**
 * Estimate equal-sided cartons from their CBM, then place them in rows and layers.
 * Positions are centers in meters: x = length, y = up, z = width; origin is a floor corner.
 * The envelope is illustrative, scaled to capacity, not a measured container specification.
 * Every input item appears in exactly one of placements or unplaced.
 */
export function buildContainerLayout(container, maxCbm) {
  const capacity = Number(maxCbm);
  const validCapacity = Number.isFinite(capacity) && capacity > 0;
  const scale = validCapacity ? Math.cbrt(capacity / 76) : 0;
  const containerSize = [12.16 * scale, 2.5 * scale, 2.5 * scale];
  const [length, height, width] = containerSize;
  const used = Number(container.totalCBM);
  const usedCbm = Number.isFinite(used) && used >= 0 ? used : null;
  const products = new Map();
  const cartons = container.items.map((item, itemIndex) => {
    const key = productKey(item);
    if (!products.has(key)) products.set(key, {
      productKey: key, markNo: item.markNo || '', description: item.description || 'Unnamed product',
      color: productColor(key), totalCartons: 0, placedCartons: 0,
    });
    products.get(key).totalCartons++;
    const cbm = Number(item.cbmPerCtn || item.singleCtnCBM || 0);
    return { itemIndex, productKey: key, cbm, side: Number.isFinite(cbm) && cbm > 0 ? Math.cbrt(cbm) : 0 };
  });
  // Sorting a derived array leaves the optimizer's order and objects untouched.
  cartons.sort((a, b) => b.side - a.side || a.itemIndex - b.itemIndex);
  const placements = [];
  const unplaced = [];
  const epsilon = Math.max(...containerSize) * 1e-10;
  let x = 0, y = 0, z = 0, rowDepth = 0, layerHeight = 0;

  cartons.forEach(({ itemIndex, productKey: key, cbm, side }) => {
    const reject = reason => unplaced.push({ itemIndex, productKey: key, reason });
    if (!validCapacity) return reject('Invalid container capacity');
    if (!side) return reject('Missing or invalid carton CBM');
    if (side > Math.min(length, height, width) + epsilon) return reject('Estimated carton exceeds container dimensions');
    let nextX = x, nextY = y, nextZ = z, nextRowDepth = rowDepth, nextLayerHeight = layerHeight;
    if (nextX + side > length + epsilon) {
      nextX = 0;
      nextZ += nextRowDepth;
      nextRowDepth = 0;
    }
    if (nextZ + side > width + epsilon) {
      nextX = 0;
      nextZ = 0;
      nextY += nextLayerHeight;
      nextRowDepth = 0;
      nextLayerHeight = 0;
    }
    if (nextY + side > height + epsilon) return reject('No room in estimated shelf layout');
    placements.push({ itemIndex, productKey: key, position: [nextX + side / 2, nextY + side / 2, nextZ + side / 2], size: [side, side, side], cbm });
    products.get(key).placedCartons++;
    x = nextX + side;
    y = nextY;
    z = nextZ;
    rowDepth = Math.max(nextRowDepth, side);
    layerHeight = Math.max(nextLayerHeight, side);
    return undefined;
  });

  return {
    containerId: container.id, geometryMode: 'estimated-from-cbm', units: 'meters', containerSize,
    capacityCbm: validCapacity ? capacity : null, usedCbm,
    fillPercent: validCapacity && usedCbm !== null ? usedCbm / capacity * 100 : null,
    placements, unplaced, products: [...products.values()],
  };
}
