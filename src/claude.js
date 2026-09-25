import Anthropic from "@anthropic-ai/sdk";
import { betaZodOutputFormat } from "@anthropic-ai/sdk/helpers/beta/zod";
import { WorkbookSpec } from "./schema.js";

const MODEL = process.env.CLAUDE_MODEL || "claude-opus-5";
const EFFORT = process.env.CLAUDE_EFFORT || "high";

const SYSTEM = `You design Excel workbooks from a user's description (typed or dictated by voice, so expect
filler words and transcription slips - infer the intent).

Rules:
- Return the complete workbook spec every time. When a current workbook is provided, apply the
  requested change to it and keep everything else as it was.
- Row 1 of every sheet is the header row; data rows start at row 2. Use Excel formulas (strings
  starting with "=") for anything derived - totals, differences, percentages, lookups across sheets -
  so the file stays live when the user edits it. Reference other sheets as 'Sheet Name'!A2.
- Use realistic sample data when the user asks for an example or template; leave input cells empty
  (null) when they want a blank template to fill in.
- Pick number formats that fit the data (currency, percent, dates as yyyy-mm-dd strings with a date
  format). Keep row values aligned with the columns array.
- Write sheet names, headers and sample data in the language the user wrote or spoke in.`;

let client;
function getClient() {
  client ??= new Anthropic();
  return client;
}

export class GenerationError extends Error {}

/**
 * Turn a prompt (plus optionally the current spec, for refinements) into a workbook spec.
 */
export async function generateSpec(prompt, current = null) {
  const content = current
    ? `Current workbook:\n${JSON.stringify(current)}\n\nRequested change:\n${prompt}`
    : prompt;

  const stream = getClient().beta.messages.stream({
    model: MODEL,
    max_tokens: 64000,
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default",
    thinking: { type: "adaptive" },
    output_config: { effort: EFFORT, format: betaZodOutputFormat(WorkbookSpec) },
    system: SYSTEM,
    messages: [{ role: "user", content }],
  });
  const message = await stream.finalMessage();

  if (message.stop_reason === "refusal") {
    throw new GenerationError(
      `Claude declined this request${message.stop_details?.explanation ? `: ${message.stop_details.explanation}` : "."}`,
    );
  }
  if (message.stop_reason === "max_tokens") {
    throw new GenerationError("The workbook is too large to generate in one go. Try asking for fewer rows.");
  }
  if (!message.parsed_output) {
    throw new GenerationError("Claude returned a response that could not be parsed into a workbook.");
  }
  return message.parsed_output;
}
