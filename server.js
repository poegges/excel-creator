import express from "express";
import { fileURLToPath } from "node:url";
import path from "node:path";
import Anthropic from "@anthropic-ai/sdk";

try {
  process.loadEnvFile();
} catch {
  // no .env file - rely on the real environment
}

const { generateSpec, GenerationError } = await import("./src/claude.js");
const { buildWorkbook, safeFileName } = await import("./src/xlsx.js");
const { WorkbookSpec } = await import("./src/schema.js");

const here = path.dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(here, "public")));

app.post("/api/generate", async (req, res) => {
  const prompt = String(req.body?.prompt ?? "").trim();
  if (!prompt) return res.status(400).json({ error: "Please describe the spreadsheet you want." });

  let current = null;
  if (req.body?.current) {
    const parsed = WorkbookSpec.safeParse(req.body.current);
    if (!parsed.success) return res.status(400).json({ error: "The current workbook is invalid." });
    current = parsed.data;
  }

  try {
    res.json({ spec: await generateSpec(prompt, current) });
  } catch (err) {
    console.error(err);
    if (err instanceof GenerationError) return res.status(422).json({ error: err.message });
    if (err instanceof Anthropic.AuthenticationError) {
      return res.status(500).json({ error: "Invalid or missing ANTHROPIC_API_KEY on the server." });
    }
    if (err instanceof Anthropic.RateLimitError) {
      return res.status(429).json({ error: "Rate limited by the Claude API - please wait a moment and retry." });
    }
    if (err instanceof Anthropic.APIError) {
      return res.status(502).json({ error: `Claude API error: ${err.message}` });
    }
    res.status(500).json({ error: "Unexpected server error." });
  }
});

app.post("/api/xlsx", async (req, res) => {
  const parsed = WorkbookSpec.safeParse(req.body?.spec);
  if (!parsed.success) return res.status(400).json({ error: "Invalid workbook spec." });
  try {
    const buffer = await buildWorkbook(parsed.data);
    const name = safeFileName(parsed.data.fileName);
    res
      .set({
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="workbook.xlsx"; filename*=UTF-8''${encodeURIComponent(name)}`,
      })
      .send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not build the Excel file." });
  }
});

const port = Number(process.env.PORT) || 3000;
app.listen(port, () => {
  console.log(`Excel Creator running at http://localhost:${port}`);
  if (!process.env.ANTHROPIC_API_KEY && !process.env.ANTHROPIC_AUTH_TOKEN) {
    console.warn("Warning: ANTHROPIC_API_KEY is not set - copy .env.example to .env and add your key.");
  }
});
