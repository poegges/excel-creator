import ExcelJS from "exceljs";

const ALT_ROW_FILL = "FFF2F2F2";
const DEFAULT_HEADER_FILL = "FFD9D9D9";

function argb(hex, fallback) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex ?? "").trim());
  return m ? `FF${m[1].toUpperCase()}` : fallback;
}

function isDark(argbColor) {
  const r = parseInt(argbColor.slice(2, 4), 16);
  const g = parseInt(argbColor.slice(4, 6), 16);
  const b = parseInt(argbColor.slice(6, 8), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b < 140;
}

function solidFill(color) {
  return { type: "pattern", pattern: "solid", fgColor: { argb: color }, bgColor: { argb: color } };
}

function sheetName(name, used) {
  let base = String(name || "Sheet").replace(/[[\]:*?/\\]/g, " ").trim().slice(0, 31) || "Sheet";
  let candidate = base;
  for (let i = 2; used.has(candidate.toLowerCase()); i++) {
    const suffix = ` (${i})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function toCellValue(value) {
  if (typeof value === "string" && value.startsWith("=") && value.length > 1) {
    return { formula: value.slice(1) };
  }
  return value ?? null;
}

function autoWidth(sheet, colIndex) {
  let longest = String(sheet.columns[colIndex].header ?? "").length;
  for (const row of sheet.rows) {
    const v = row[colIndex];
    if (v == null || (typeof v === "string" && v.startsWith("="))) continue;
    longest = Math.max(longest, String(v).length);
  }
  return Math.min(Math.max(longest + 2, 8), 60);
}

function addSheet(workbook, sheet, used) {
  const ws = workbook.addWorksheet(sheetName(sheet.name, used));
  const colCount = sheet.columns.length;
  const lastDataRow = sheet.rows.length + 1;

  ws.columns = sheet.columns.map((col, i) => ({
    header: col.header,
    width: col.width ?? autoWidth(sheet, i),
    style: {
      ...(col.numberFormat ? { numFmt: col.numberFormat } : {}),
      ...(col.align ? { alignment: { horizontal: col.align } } : {}),
    },
  }));

  const headerFill = argb(sheet.headerFill, DEFAULT_HEADER_FILL);
  const header = ws.getRow(1);
  header.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: isDark(headerFill) ? "FFFFFFFF" : "FF000000" } };
    cell.fill = solidFill(headerFill);
    cell.alignment = { vertical: "middle", horizontal: cell.alignment?.horizontal ?? "left" };
    cell.border = { bottom: { style: "thin" } };
  });

  sheet.rows.forEach((values, r) => {
    const row = ws.addRow(values.slice(0, colCount).map(toCellValue));
    if (sheet.zebra && r % 2 === 1) {
      for (let c = 1; c <= colCount; c++) row.getCell(c).fill = solidFill(ALT_ROW_FILL);
    }
  });

  if (sheet.totals && sheet.rows.length > 0) {
    const values = new Array(colCount).fill(null);
    for (const { column, fn } of sheet.totals.columns) {
      if (column < 0 || column >= colCount) continue;
      const letter = ws.getColumn(column + 1).letter;
      values[column] = { formula: `${fn}(${letter}2:${letter}${lastDataRow})` };
    }
    if (values[0] == null) values[0] = sheet.totals.label;
    const row = ws.addRow(values);
    row.font = { bold: true };
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.border = { top: { style: "thin" }, bottom: { style: "double" } };
    });
  }

  if (sheet.freezeHeader) ws.views = [{ state: "frozen", ySplit: 1 }];
  if (sheet.autoFilter && colCount > 0) {
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: colCount } };
  }

  for (const cf of sheet.conditionalFormats) {
    if (cf.column < 0 || cf.column >= colCount || sheet.rows.length === 0) continue;
    const letter = ws.getColumn(cf.column + 1).letter;
    const style = { fill: solidFill(argb(cf.fillColor, "FFFFEB9C")) };
    const quote = (v) => (v !== "" && !Number.isNaN(Number(v)) ? v : `"${String(v).replace(/"/g, '""')}"`);
    const rule =
      cf.rule === "containsText"
        ? { type: "containsText", operator: "containsText", text: cf.value, style }
        : {
            type: "cellIs",
            operator: cf.rule,
            formulae: cf.rule === "between" ? [quote(cf.value), quote(cf.value2 ?? cf.value)] : [quote(cf.value)],
            style,
          };
    ws.addConditionalFormatting({ ref: `${letter}2:${letter}${lastDataRow}`, rules: [{ priority: 1, ...rule }] });
  }
}

/** Build an .xlsx file from a workbook spec and return it as a Buffer. */
export async function buildWorkbook(spec) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Excel Creator";
  workbook.created = new Date();
  const used = new Set();
  for (const sheet of spec.sheets) addSheet(workbook, sheet, used);
  if (spec.sheets.length === 0) workbook.addWorksheet("Sheet1");
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

export function safeFileName(name) {
  const cleaned = String(name || "workbook").replace(/[^\p{L}\p{N}._ -]+/gu, "").trim().slice(0, 80);
  return `${cleaned || "workbook"}.xlsx`;
}
