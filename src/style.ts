import ExcelJS from "exceljs";

export type StyleConfig = {
  borderAll?: Partial<ExcelJS.Border>;
  border?: Partial<ExcelJS.Borders>;
  alignment?: Partial<ExcelJS.Alignment>;
  fill?: ExcelJS.Fill;
  font?: Partial<ExcelJS.Font>;
  numFmt?: string;
};

export const applyCellStyle = (
  cell: ExcelJS.Cell,
  style?: StyleConfig
): void => {
  if (!style) return;

  if (style.borderAll) {
    cell.border = {
      top: style.borderAll,
      left: style.borderAll,
      bottom: style.borderAll,
      right: style.borderAll,
    };
  }
  if (style.border) {
    cell.border = style.border;
  }
  if (style.alignment) {
    cell.alignment = style.alignment;
  }
  if (style.fill) {
    cell.fill = style.fill;
  }
  if (style.font) {
    cell.font = style.font;
  }
  if (style.numFmt) {
    cell.numFmt = style.numFmt;
  }
};
