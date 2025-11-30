# ExcelJS Helper

A simple helper library for [ExcelJS](https://github.com/exceljs/exceljs#readme) that provides utilities for creating styled cells, tables, and formatting Excel workbooks.

## Installation

```bash
npm install exceljs-helper
```

## Quick Start

```typescript
import ExcelJS from "exceljs";
import { text, createTableExcel, applyCellStyle } from "exceljs-helper";

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Sheet1");

// Add styled text
text(sheet, "A1", "Hello World", {
  font: { bold: true },
  alignment: { horizontal: "center" },
});

// Save to file
await workbook.xlsx.writeFile("output.xlsx");
```

## API Reference

### `text(sheet, position, value, config?)`

Add text to a cell with optional styling and cell merging.

**Parameters:**

- `sheet` - ExcelJS Worksheet instance
- `position` - Cell position as string (e.g., `"A1"`) or object with `start` and `end` for merged cells
- `value` - The cell value (string, number, boolean, etc.)
- `config?` - Optional StyleConfig object

**Example:**

```typescript
import { text } from "exceljs-helper";

// Simple cell
text(sheet, "A1", "Title");

// Merged cells with styling
text(sheet, { start: "A1", end: "C1" }, "Header", {
  font: { bold: true, size: 14 },
  fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF0F172A" } },
  alignment: { horizontal: "center", vertical: "middle" },
});
```

### `createTableExcel(sheet, x, y, columnData, config?)`

Generate a structured table with headers and data rows, supporting merged cells and per-cell styling.

**Parameters:**

- `sheet` - ExcelJS Worksheet instance
- `x` - Column position (1-based index)
- `y` - Row position (1-based index)
- `columnData` - Table structure object
  - `headers` - Array of header values or 2D array for multi-row headers
    - Each header can be a string or object: `{ value: string, colspan?: number, rowspan?: number }`
  - `rows` - Array of data rows
    - Each cell can be a value or object: `{ value: any, colspan?: number, rowspan?: number, cellStyle?: StyleConfig }`
- `config?` - Optional styling configuration
  - `headerStyle?` - StyleConfig applied to all header cells
  - `cellStyle?` - StyleConfig applied to all data cells
  - `columnStyles?` - Object mapping column index to StyleConfig
  - `cellStyles?` - Object mapping cell key `"rowIndex,colIndex"` to StyleConfig

**Example:**

```typescript
import { createTableExcel } from "exceljs-helper";

createTableExcel(
  sheet,
  1,
  1,
  {
    headers: [
      [
        { value: "Region", rowspan: 2 },
        { value: "Sales", colspan: 2 },
        { value: "Growth", rowspan: 2 },
      ],
      ["Q1", "Q2"],
    ],
    rows: [
      ["North", 12000, 13500, { value: 0.12, cellStyle: { numFmt: "0.00%" } }],
      ["South", 9800, 11250, { value: 0.18, cellStyle: { numFmt: "0.00%" } }],
      ["West", 14300, 15670, 0.09],
    ],
  },
  {
    headerStyle: {
      font: { bold: true, color: { argb: "FFFFFFFF" } },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FF0F172A" },
      },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    cellStyle: {
      borderAll: { style: "thin", color: { argb: "FFCBD5F5" } },
      alignment: { vertical: "middle" },
    },
    columnStyles: {
      1: { alignment: { horizontal: "left" } },
    },
  }
);
```

### `applyCellStyle(cell, style?)`

Apply styling to an ExcelJS Cell instance.

**Parameters:**

- `cell` - ExcelJS Cell instance
- `style?` - StyleConfig object

**Example:**

```typescript
import { applyCellStyle } from "exceljs-helper";

const cell = sheet.getCell("A1");
applyCellStyle(cell, {
  borderAll: { style: "thin", color: { argb: "FF000000" } },
  fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF00" } },
  font: { bold: true, size: 12 },
  alignment: { horizontal: "center", vertical: "middle" },
  numFmt: "0.00%",
});
```

### `StyleConfig` Type

Configuration object for cell styling:

```typescript
type StyleConfig = {
  borderAll?: Partial<ExcelJS.Border>; // Apply border to all sides
  border?: Partial<ExcelJS.Borders>; // Custom borders per side
  alignment?: Partial<ExcelJS.Alignment>; // Text alignment
  fill?: ExcelJS.Fill; // Cell background fill
  font?: Partial<ExcelJS.Font>; // Font properties
  numFmt?: string; // Number format (e.g., "0.00%", "yyyy-mm-dd")
};
```

you can read more in [Exceljs Doc](https://github.com/exceljs/exceljs?tab=readme-ov-file#styles)

### Utility Functions

#### `numberToAlphabet(column: number): string`

Convert a column index (1-based) to Excel column letter.

```typescript
import { numberToAlphabet } from "exceljs-helper";

numberToAlphabet(1); // "A"
numberToAlphabet(26); // "Z"
numberToAlphabet(27); // "AA"
```

#### `columnToNumber(col: string): number`

Convert an Excel column letter to column index (1-based).

```typescript
import { columnToNumber } from "exceljs-helper";

columnToNumber("A"); // 1
columnToNumber("Z"); // 26
columnToNumber("AA"); // 27
```

## Examples

Clone this repo and install dependencies:

```bash
npm install
```

Generate the sample workbooks:

```bash
npm run build
node dist/examples/text-example.js
node dist/examples/table-example.js
```

The scripts create `.xlsx` files under `dist/examples/output/`.

> Prefer running the compiled files so the examples use the same code that ships to npm. If you want to run TypeScript directly during development, install a runner such as `tsx` and execute `npx tsx examples/table-example.ts`.

## License

MIT
