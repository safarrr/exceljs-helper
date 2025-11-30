import ExcelJS from "exceljs";
import { mkdir } from "node:fs/promises";
import path from "node:path";
import { createTableExcel } from "../src/table";

(async () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sales Report");

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
        [
          "North",
          12000,
          13500,
          { value: 0.12, cellStyle: { numFmt: "0.00%" } },
        ],
        ["South", 9800, 11250, { value: 0.18, cellStyle: { numFmt: "0.00%" } }],
        [
          "West",
          14300,
          15670,
          {
            value: 0.09,
            cellStyle: {
              font: {
                color: { argb: "ffffff" },
              },
              numFmt: "0.00%",
              fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "f70202" },
              },
            },
          },
        ],
      ],
    },
    {
      headerStyle: {
        alignment: { horizontal: "center", vertical: "middle" },
        font: { bold: true, color: { argb: "FFFFFFFF" } },
        fill: {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0F172A" },
        },
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

  const outDir = path.resolve(__dirname, "output");
  await mkdir(outDir, { recursive: true });
  const filePath = path.join(outDir, "table-example.xlsx");

  await workbook.xlsx.writeFile(filePath);
  console.log(`Table example saved to ${filePath}`);
})().catch((error) => {
  console.error("Failed to generate table example", error);
  process.exit(1);
});
