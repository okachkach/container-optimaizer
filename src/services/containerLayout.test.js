import { buildContainerLayout, productColor, productKey } from './containerLayout';
import { generateContainers, moveProductBetweenContainers, removeProductFromContainer, addContainerProducts } from './optimizer';

const product = (markNo, cbmPerCtn, ctn = 1) => ({ markNo, description: `Product ${markNo}`, cbmPerCtn, ctn, gwPerCtn: 10, nwPerCtn: 8, pcsPerCtn: 5 });
const container = items => ({ id: 1, items, totalCBM: items.reduce((sum, item) => sum + (item.cbmPerCtn || 0), 0) });

function verifyLayout(layout, count) {
  const indices = [...layout.placements, ...layout.unplaced].map(item => item.itemIndex);
  expect(indices.sort((a, b) => a - b)).toEqual(Array.from({ length: count }, (_, i) => i));
  layout.placements.forEach((carton, index) => {
    expect(carton.size.reduce((volume, side) => volume * side, 1)).toBeCloseTo(carton.cbm, 10);
    carton.position.forEach((center, axis) => {
      expect(center - carton.size[axis] / 2).toBeGreaterThanOrEqual(-1e-8);
      expect(center + carton.size[axis] / 2).toBeLessThanOrEqual(layout.containerSize[axis] + 1e-8);
    });
    // Every pair must be separated on at least one axis (touching faces are allowed).
    for (let otherIndex = index + 1; otherIndex < layout.placements.length; otherIndex++) {
      const other = layout.placements[otherIndex];
      const separated = carton.position.some((center, axis) =>
        Math.abs(center - other.position[axis]) >= (carton.size[axis] + other.size[axis]) / 2 - 1e-8);
      if (!separated) throw new Error(`Cartons ${carton.itemIndex} and ${other.itemIndex} overlap`);
    }
  });
}

test('lays out hundreds of optimizer-assigned cartons deterministically without mutating data', () => {
  const { containers } = generateContainers([product('A', 0.125, 300), product('B', 0.064, 70)], 27000, 76);
  const packed = containers[0];
  const before = JSON.stringify(packed);
  packed.items.forEach(Object.freeze);
  Object.freeze(packed.items);
  Object.freeze(packed);
  const layout = buildContainerLayout(packed, 76);
  expect(layout).toEqual(buildContainerLayout(packed, 76));
  expect(JSON.stringify(packed)).toBe(before);
  expect(layout.placements).toHaveLength(370);
  expect(layout.containerSize.reduce((v, side) => v * side, 1)).toBeCloseTo(76);
  expect(layout.fillPercent).toBeCloseTo(packed.totalCBM / 76 * 100);
  expect(new Set(layout.placements.map(item => item.position[1])).size).toBeGreaterThan(1);
  expect(new Set(layout.placements.map(item => item.position[2])).size).toBeGreaterThan(1);
  verifyLayout(layout, packed.items.length);
});

test('reports every invalid, oversized, and overflowing carton without concealing assigned volume', () => {
  const items = [product('large', 30), product('zero', 0), product('negative', -1), product('invalid', NaN),
    ...Array.from({ length: 100 }, () => product('cube', 1))];
  const layout = buildContainerLayout({ ...container(items), totalCBM: 130 }, 76);
  verifyLayout(layout, items.length);
  expect(layout.unplaced.map(item => item.reason)).toEqual(expect.arrayContaining([
    'Estimated carton exceeds container dimensions', 'Missing or invalid carton CBM', 'No room in estimated shelf layout',
  ]));
  expect(layout.fillPercent).toBeCloseTo(130 / 76 * 100);
  expect(layout.products.reduce((sum, entry) => sum + entry.totalCartons, 0)).toBe(items.length);
  expect(layout.products.reduce((sum, entry) => sum + entry.placedCartons, 0)).toBe(layout.placements.length);
});

test.each([0, -10, NaN, Infinity])('handles invalid capacity %s without non-finite coordinates', capacity => {
  const layout = buildContainerLayout(container([product('A', 0.125)]), capacity);
  expect(layout.capacityCbm).toBeNull();
  expect(layout.fillPercent).toBeNull();
  expect(layout.placements).toEqual([]);
  expect(layout.unplaced).toHaveLength(1);
  expect(layout.containerSize.every(Number.isFinite)).toBe(true);
});

test('handles empty containers and the per-carton fallback field', () => {
  const empty = buildContainerLayout(container([]), 38);
  expect(empty.fillPercent).toBe(0);
  expect(empty.placements).toEqual([]);
  expect(empty.containerSize.reduce((v, side) => v * side, 1)).toBeCloseTo(38);
  const fallback = buildContainerLayout({ id: 1, totalCBM: 0.125, items: [{ ...product('A', 0), singleCtnCBM: 0.125 }] }, 76);
  expect(fallback.placements[0].size).toEqual([0.5, 0.5, 0.5]);
});

test('keeps product colors consistent across containers, order changes, and separator characters', () => {
  const a = product('A', 0.125), b = product('B', 0.064);
  const first = buildContainerLayout(container([a, b]), 76);
  const second = buildContainerLayout(container([b, a]), 76);
  expect(first.products[0].color).toBe(second.products[1].color);
  expect(productColor(productKey(a))).toBe(first.products[0].color);
  expect(productKey({ markNo: 'a|b', description: 'c' })).not.toBe(productKey({ markNo: 'a', description: 'b|c' }));
});

test('reflects existing move, remove, and add operations without changing them', () => {
  const a = product('A', 0.125, 4);
  const { containers } = generateContainers([a], 27000, 76);
  let next = [...containers, { ...container([]), id: 2 }];
  next = moveProductBetweenContainers(next, next[0].id, `${a.markNo}|${a.description}`, 2, 2);
  next.forEach(item => verifyLayout(buildContainerLayout(item, 76), 2));
  next = removeProductFromContainer(next, 2, `${a.markNo}|${a.description}`, 1);
  expect(buildContainerLayout(next.find(item => item.id === 2), 76).placements).toHaveLength(1);
  next = addContainerProducts(next, 2, a, 3, 27000, 76).containers;
  const layout = buildContainerLayout(next.find(item => item.id === 2), 76);
  expect(layout.placements).toHaveLength(4);
  expect(layout.usedCbm).toBe(0.5);
});
