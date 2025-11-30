import ExcelJS from "exceljs";
import { applyCellStyle, StyleConfig } from "./style";
/**
 * Add text to a cell
 * @param sheet - The worksheet to add the text to
 * @param position - The position of the text
 * @param value - The value of the text
 * @param config - The configuration for the text
 */
export const text = (
  sheet: ExcelJS.Worksheet,
  position: { start: string; end: string } | string,
  value: ExcelJS.CellValue,
  config?: StyleConfig
) => {
  let cell = sheet.getCell(
    typeof position === "object" ? position.start : position
  );
  if (typeof position === "object") {
    sheet.mergeCells(`${position.start}:${position.end}`);
  }
  cell.value = value;
  applyCellStyle(cell, config);
};
