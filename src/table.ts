import ExcelJS from "exceljs";
import { applyCellStyle, StyleConfig } from "./style";
import { numberToAlphabet } from "./utils";
export const createTableExcel = (
  sheet: ExcelJS.Worksheet,
  x: number, // column position (1-based)
  y: number, // row position (1-based)
  columnData: {
    headers:
      | (string | { value: string; colspan?: number; rowspan?: number })[][]
      | (string | { value: string; colspan?: number; rowspan?: number })[];
    rows: (
      | ExcelJS.CellValue
      | {
          value: ExcelJS.CellValue;
          colspan?: number;
          rowspan?: number;
          cellStyle?: {
            borderAll?: Partial<ExcelJS.Border>;
            border?: Partial<ExcelJS.Borders>;
            alignment?: Partial<ExcelJS.Alignment>;
            fill?: ExcelJS.Fill;
            font?: Partial<ExcelJS.Font>;
            numFmt?: string;
          };
        }
    )[][];
  },
  config?: {
    headerStyle?: StyleConfig;
    cellStyle?: StyleConfig;
    columnStyles?: { [columnIndex: number]: StyleConfig };
    cellStyles?: { [key: string]: StyleConfig };
  }
) => {
  if (x < 1) {
    throw new Error("Column position 'x' must be a positive 1-based index");
  }

  const columnOffset = x - 1;
  const getActualColumn = (offset: number) => columnOffset + offset + 1;

  // Determine if headers is multi-dimensional array
  const isMultiHeader = Array.isArray(columnData.headers[0]);
  const headerRows = isMultiHeader
    ? (columnData.headers as (
        | string
        | { value: string; colspan?: number; rowspan?: number }
      )[][])
    : [
        columnData.headers as (
          | string
          | { value: string; colspan?: number; rowspan?: number }
        )[],
      ];

  // Create headers
  let totalHeaderRows = headerRows.length;
  const occupied = new Set<string>();
  const occupiedKey = (row: number, col: number) => `${row},${col}`;
  const hasIntersection = (
    row: number,
    col: number,
    colspan: number,
    rowspan: number
  ) => {
    for (let r = 0; r < rowspan; r++) {
      for (let c = 0; c < colspan; c++) {
        if (occupied.has(occupiedKey(row + r, col + c))) {
          return true;
        }
      }
    }
    return false;
  };

  headerRows.forEach((headerRow, rowIndex) => {
    let headerColOffset = 0;

    headerRow.forEach((header) => {
      const colspan =
        typeof header === "string" ? 1 : Math.max(header.colspan || 1, 1);
      const rowspan =
        typeof header === "string" ? 1 : Math.max(header.rowspan || 1, 1);
      const isPlaceholder =
        typeof header === "string" && header.trim().length === 0;

      if (isPlaceholder) {
        headerColOffset += colspan;
        return;
      }

      let startCol = headerColOffset;
      while (hasIntersection(rowIndex, startCol, colspan, rowspan)) {
        startCol++;
      }

      const actualCol = getActualColumn(startCol);
      const actualRow = y + rowIndex;
      const cellAddress = `${numberToAlphabet(actualCol)}${actualRow}`;
      const cell = sheet.getCell(cellAddress);

      if (typeof header === "string") {
        cell.value = header;
      } else {
        cell.value = header.value;

        if (rowspan > 1 || colspan > 1) {
          const endCol = actualCol + colspan - 1;
          const endRow = actualRow + rowspan - 1;
          const endAddress = `${numberToAlphabet(endCol)}${endRow}`;
          sheet.mergeCells(`${cellAddress}:${endAddress}`);
        }
      }

      for (let r = 1; r < rowspan; r++) {
        for (let c = 0; c < colspan; c++) {
          occupied.add(occupiedKey(rowIndex + r, startCol + c));
        }
      }

      headerColOffset = startCol + colspan;

      // Apply header styling
      applyCellStyle(cell, config?.headerStyle);
    });
  });

  // Create data rows
  columnData.rows.forEach((row, rowIndex) => {
    let colOffset = 0;
    row.forEach((cellValue, colIndex) => {
      const actualCol = getActualColumn(colOffset);
      const actualRow = y + totalHeaderRows + rowIndex;
      const cellAddress = `${numberToAlphabet(actualCol)}${actualRow}`;
      const cell = sheet.getCell(cellAddress);

      let inlineCellStyle = null;
      if (
        typeof cellValue === "object" &&
        cellValue !== null &&
        "value" in cellValue
      ) {
        cell.value = cellValue.value;
        inlineCellStyle = cellValue.cellStyle;

        // Handle colspan and rowspan
        if (cellValue.colspan || cellValue.rowspan) {
          const endCol = actualCol + (cellValue.colspan || 1) - 1;
          const endRow = actualRow + (cellValue.rowspan || 1) - 1;
          const endAddress = `${numberToAlphabet(endCol)}${endRow}`;
          sheet.mergeCells(`${cellAddress}:${endAddress}`);
        }

        colOffset += cellValue.colspan || 1;
      } else {
        cell.value = cellValue;
        colOffset++;
      }

      // Apply cell styling (priority: cellStyles > columnStyles > cellStyle)
      // Apply general cell style
      applyCellStyle(cell, config?.cellStyle);

      // Apply column-specific style
      if (config?.columnStyles?.[colOffset]) {
        applyCellStyle(cell, config.columnStyles[colOffset]);
      }

      // Apply cell-specific style (highest priority)
      const cellKey = `${rowIndex},${colOffset}`;
      if (config?.cellStyles?.[cellKey]) {
        applyCellStyle(cell, config.cellStyles[cellKey]);
      }

      // Apply inline cell style (highest priority)
      if (inlineCellStyle) {
        applyCellStyle(cell, inlineCellStyle);
      }
    });
  });

  // Return table dimensions for reference
  return {
    startColumn: x,
    startRow: y,
    endColumn: x + columnData.headers.length - 1,
    endRow: y + columnData.rows.length,
    totalColumns: columnData.headers.length,
    totalRows: columnData.rows.length + 1, // +1 for header
  };
};
