import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { buildWorkbook, safeFileName } from "../src/xlsx.js";
import { WorkbookSpec } from "../src/schema.js";

const spec = {
  fileName: "budget",
  summary: "A small budget.",
  sheets: [
    {
      name: "Budget: 2026/Q1",
      columns: [
        { header: "Category", width: null, numberFormat: null, align: null },
        { header: "Planned", width: 12, numberFormat: "#,##0.00", align: "right" },
        { header: "Actual", width: 12, numberFormat: "#,##0.00", align: "right" },
        { header: "Diff", width: null, numberFormat: "#,##0.00", align: "right" },
      ],
      rows: [
        ["Rent", 1200, 1200, "=C2-B2"],
        ["Food", 400, 455.5, "=C3-B3"],
        ["Travel", 150, null, "=C4-B4"],
      ],
      freezeHeader: true,
      autoFilter: true,
      headerFill: "#1F4E78",
      zebra: true,
      totals: { label: "Total", columns: [{ column: 1, fn: "SUM" }, { column: 3, fn: "SUM" }] },
      conditionalFormats: [
        { column: 3, rule: "greaterThan", value: "0", value2: null, fillColor: "#F8CBAD" },
        { column: 0, rule: "containsText", value: "Food", value2: null, fillColor: "#FFEB9C" },
      ],
    },
    { name: "Budget: 2026/Q1", columns: [{ header: "A", width: null, numberFormat: null, align: null }], rows: [], freezeHeader: false, autoFilter: false, headerFill: null, zebra: false, totals: null, conditionalFormats: [] },
  ],
};

test("spec fixture matches the schema", () => {
  assert.ok(WorkbookSpec.safeParse(spec).success);
});

test("builds a workbook with formulas, totals and styling", async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await buildWorkbook(spec));

  assert.deepEqual(wb.worksheets.map((w) => w.name), ["Budget  2026 Q1", "Budget  2026 Q1 (2)"]);
  const ws = wb.worksheets[0];
  assert.equal(ws.getCell("A1").value, "Category");
  assert.equal(ws.getCell("A1").font.bold, true);
  assert.equal(ws.getCell("B3").value, 400);
  assert.equal(ws.getCell("D2").value.formula, "C2-B2");
  assert.equal(ws.getCell("A5").value, "Total");
  assert.equal(ws.getCell("B5").value.formula, "SUM(B2:B4)");
  assert.equal(ws.getCell("C5").value, null);
  assert.equal(ws.getCell("B2").numFmt, "#,##0.00");
  assert.equal(ws.views[0].state, "frozen");
});

test("safeFileName strips unsafe characters", () => {
  assert.equal(safeFileName('my "report"/2026'), "my report2026.xlsx");
  assert.equal(safeFileName(""), "workbook.xlsx");
  assert.equal(safeFileName("Übersicht Kosten"), "Übersicht Kosten.xlsx");
});

test("schema converts to a structured-output format", async () => {
  const { betaZodOutputFormat } = await import("@anthropic-ai/sdk/helpers/beta/zod");
  const fmt = betaZodOutputFormat(WorkbookSpec);
  assert.equal(fmt.type, "json_schema");
});
