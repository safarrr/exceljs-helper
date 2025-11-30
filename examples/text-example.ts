import ExcelJS from "exceljs";
import { mkdir } from "node:fs/promises";
import path from "node:path";
import { text } from "../src/text";

(async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Invoice");

  text(sheet, { start: "A1", end: "C1" }, "Invoice #1024", {
    alignment: { horizontal: "center", vertical: "middle" },
    font: { bold: true, size: 14 },
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0F2FE" } },
    borderAll: { style: "thin", color: { argb: "FF93C5FD" } },
  });

  text(sheet, "A3", "Customer", { font: { bold: true } });
  text(sheet, "B3", "GadgetCo");
  text(sheet, "A4", "Amount", { font: { bold: true } });
  text(sheet, "B4", 2599, { numFmt: "$#,##0.00" });

  const outDir = path.resolve(__dirname, "output");
  await mkdir(outDir, { recursive: true });
  const filePath = path.join(outDir, "text-example.xlsx");

  await workbook.xlsx.writeFile(filePath);
  console.log(`Text example saved to ${filePath}`);
})().catch((error) => {
  console.error("Failed to generate text example", error);
  process.exit(1);
});
