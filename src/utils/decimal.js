export const DECIMAL_PLACES = 4;
export const DECIMAL_STEP = 10 ** -DECIMAL_PLACES;

/**
 * Display measurements and prices without hiding values as small as 0.0001.
 */
export const formatDecimal = (value) => {
  const number = Number(value);
  return (Number.isFinite(number) ? number : 0).toFixed(DECIMAL_PLACES);
};

/**
 * Keep exported/imported numeric values at the precision supported by the UI.
 */
export const roundDecimal = (value) => Number(formatDecimal(value));
