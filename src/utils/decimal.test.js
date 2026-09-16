import { DECIMAL_STEP, formatDecimal, roundDecimal } from './decimal';

test('supports values as small as four decimal places', () => {
  expect(DECIMAL_STEP).toBe(0.0001);
  expect(formatDecimal(0.0001)).toBe('0.0001');
  expect(roundDecimal(0.0001)).toBe(0.0001);
});

test('rounds longer floating-point values consistently', () => {
  expect(formatDecimal(1.23456)).toBe('1.2346');
  expect(roundDecimal(1.23456)).toBe(1.2346);
});
