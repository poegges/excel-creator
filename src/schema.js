import { z } from "zod";

// The workbook "spec" Claude returns. Every field is required (nullable where
// optional) so the schema works with strict structured outputs.

const Cell = z
  .union([z.string(), z.number(), z.boolean(), z.null()])
  .describe('Cell value. Strings starting with "=" are Excel formulas, e.g. "=B2*C2".');

const Column = z.object({
  header: z.string(),
  width: z.number().nullable().describe("Column width in characters; null = auto-fit."),
  numberFormat: z
    .string()
    .nullable()
    .describe('Excel number format, e.g. "#,##0.00", "0%", "yyyy-mm-dd", "€#,##0.00". null = General.'),
  align: z.enum(["left", "center", "right"]).nullable(),
});

const TotalsColumn = z.object({
  column: z.number().int().describe("0-based column index."),
  fn: z.enum(["SUM", "AVERAGE", "COUNT", "MIN", "MAX"]),
});

const ConditionalFormat = z.object({
  column: z.number().int().describe("0-based column index the rule applies to (data rows only)."),
  rule: z.enum(["greaterThan", "lessThan", "between", "equal", "containsText"]),
  value: z.string().describe("Threshold / comparison value or text to search for."),
  value2: z.string().nullable().describe('Upper bound for "between", otherwise null.'),
  fillColor: z.string().describe('Hex background color, e.g. "#F8CBAD".'),
});

const Sheet = z.object({
  name: z.string().describe("Sheet tab name, max 31 characters."),
  columns: z.array(Column),
  rows: z.array(z.array(Cell)).describe("Data rows. Row 1 in Excel is the header, so data starts at row 2."),
  freezeHeader: z.boolean(),
  autoFilter: z.boolean(),
  headerFill: z.string().nullable().describe('Hex header background, e.g. "#1F4E78". null = light gray.'),
  zebra: z.boolean().describe("Alternate row shading."),
  totals: z
    .object({ label: z.string(), columns: z.array(TotalsColumn) })
    .nullable()
    .describe("Optional totals row appended below the data."),
  conditionalFormats: z.array(ConditionalFormat),
});

export const WorkbookSpec = z.object({
  fileName: z.string().describe('File name without extension, e.g. "monthly-budget".'),
  summary: z.string().describe("One or two sentences describing what was built or changed."),
  sheets: z.array(Sheet),
});
