# Excel Creator

Describe a spreadsheet in plain language, by typing or by voice, and get a real `.xlsx` file back.
Claude designs the workbook: sheets, columns, sample data, live formulas, number formats,
totals rows and conditional formatting. You can keep refining it with follow-up prompts
("add a VAT column", "make the header dark blue", "add a second sheet with a yearly summary").

## Features

- **Prompt field connected to Claude.** The server asks Claude for a structured workbook spec
  (validated with a Zod schema via structured outputs) and builds the file with ExcelJS.
- **Speech to text.** Click **Speak**, dictate your request and press **Stop**. It uses the browser's
  Web Speech API, so there's no extra key or service to set up. Several languages are available.
  Works in Chrome, Edge and Safari (Firefox doesn't support it).
- **Iterative refinement.** When "Modify current workbook" is checked, your next prompt edits the
  existing workbook instead of starting over.
- **Live preview.** Every sheet is previewed in the browser before you download it.
- Formulas stay live in Excel. Also supports frozen headers, auto-filters, zebra rows,
  totals rows (SUM/AVERAGE/COUNT/MIN/MAX) and conditional highlighting.

## Setup

Requires Node.js 20.12 or newer.

```bash
npm install
cp .env.example .env   # then put your Anthropic API key in .env
npm start              # open http://localhost:3000
```

| Variable | Default | Purpose |
|---|---|---|
| `ANTHROPIC_API_KEY` | (required) | Your key from https://console.anthropic.com/ |
| `CLAUDE_MODEL` | `claude-opus-5` | Model used to design workbooks |
| `CLAUDE_EFFORT` | `high` | `low` / `medium` / `high`. Lower is faster and cheaper |
| `PORT` | `3000` | HTTP port |

> The microphone only works on `localhost` or over HTTPS, which is a browser security rule.
> If you deploy this, put it behind HTTPS.

## How it works

```
browser (prompt / speech) ──POST /api/generate──▶ server ──▶ Claude (structured output: WorkbookSpec)
browser (preview)         ◀──── spec JSON ────────┘
browser ──POST /api/xlsx {spec}──▶ server (ExcelJS) ──▶ .xlsx download
```

- `src/schema.js`: the workbook spec schema Claude must fill
- `src/claude.js`: the prompt and the Claude call (includes server-side refusal fallback)
- `src/xlsx.js`: converts a spec into a styled Excel file
- `public/`: the web UI, including speech recognition

## Tests

```bash
npm test
```
