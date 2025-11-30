/**
 * Convert a column index to an Excel column letter
 * @param column - The column index (1-based)
 * @returns The column letter
 */
export function numberToAlphabet(column: number) {
  if (!Number.isFinite(column) || column < 1) {
    throw new Error("Column index must be a positive integer");
  }

  let result = "";
  let current = Math.trunc(column);

  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }
  return result;
}

/**
 * Convert Excel-like column label (A, B, ... Z, AA, AB...) to number.
 * @param col - The column label
 * @returns The column index (1-based)
 */
export function columnToNumber(col: string): number {
  let result = 0;
  const chars = col.trim().toUpperCase();

  for (let i = 0; i < chars.length; i++) {
    result = result * 26 + (chars.charCodeAt(i) - 64);
  }

  return result;
}
