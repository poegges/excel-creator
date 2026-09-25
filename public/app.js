const $ = (id) => document.getElementById(id);
const promptEl = $("prompt");
const micBtn = $("mic");
const micLabel = $("mic-label");
const langEl = $("lang");
const speechStatus = $("speech-status");
const generateBtn = $("generate");
const errorEl = $("error");

let spec = null;
let activeSheet = 0;
const history = [];

// ---------- Speech to text (Web Speech API) ----------

const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition = null;
let listening = false;
let baseText = ""; // text in the box before the current dictation started

function showSpeechStatus(text) {
  speechStatus.textContent = text;
  speechStatus.hidden = !text;
}

function setListening(on) {
  listening = on;
  micBtn.classList.toggle("listening", on);
  micBtn.setAttribute("aria-pressed", String(on));
  micLabel.textContent = on ? "Stop" : "Speak";
  showSpeechStatus(on ? "Listening… speak your request, then press Stop." : "");
}

if (!SpeechRecognition) {
  micBtn.disabled = true;
  langEl.disabled = true;
  showSpeechStatus("Speech input isn't supported in this browser. Use Chrome, Edge or Safari to dictate.");
} else {
  langEl.value = localStorage.getItem("speechLang") || (navigator.language?.startsWith("de") ? "de-DE" : "en-US");

  recognition = new SpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;

  recognition.onresult = (event) => {
    let finalText = "";
    let interim = "";
    for (let i = 0; i < event.results.length; i++) {
      const r = event.results[i];
      if (r.isFinal) finalText += r[0].transcript;
      else interim += r[0].transcript;
    }
    const sep = baseText && !/\s$/.test(baseText) ? " " : "";
    promptEl.value = baseText + sep + (finalText + interim).trim();
  };

  recognition.onerror = (event) => {
    if (event.error === "no-speech" || event.error === "aborted") return;
    const messages = {
      "not-allowed": "Microphone access was denied. Allow it in your browser's site settings.",
      "audio-capture": "No microphone was found.",
      network: "Speech recognition needs an internet connection.",
    };
    setListening(false);
    showSpeechStatus(messages[event.error] || `Speech recognition error: ${event.error}`);
  };

  // Browsers end recognition after a pause; keep going until the user presses Stop.
  recognition.onend = () => {
    if (!listening) return;
    baseText = promptEl.value;
    try {
      recognition.start();
    } catch {
      setListening(false);
    }
  };

  micBtn.addEventListener("click", () => {
    if (listening) {
      setListening(false);
      recognition.stop();
      return;
    }
    baseText = promptEl.value;
    recognition.lang = langEl.value;
    try {
      recognition.start();
      setListening(true);
    } catch (err) {
      showSpeechStatus(`Could not start the microphone: ${err.message}`);
    }
  });

  langEl.addEventListener("change", () => {
    localStorage.setItem("speechLang", langEl.value);
    if (listening) {
      recognition.stop(); // onend restarts with the new language
      recognition.lang = langEl.value;
    }
  });
}

// ---------- Generation ----------

function showError(msg) {
  errorEl.textContent = msg;
  errorEl.hidden = !msg;
}

async function generate() {
  const prompt = promptEl.value.trim();
  if (!prompt) {
    showError("Type or say what the spreadsheet should contain.");
    promptEl.focus();
    return;
  }
  if (listening) micBtn.click();
  showError("");

  const refine = spec && $("refine").checked;
  generateBtn.disabled = true;
  generateBtn.innerHTML = `<span class="spinner"></span>${refine ? "Updating…" : "Creating…"}`;

  try {
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt, current: refine ? spec : null }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `Request failed (${res.status})`);

    spec = data.spec;
    activeSheet = 0;
    history.push(prompt);
    promptEl.value = "";
    render();
  } catch (err) {
    showError(err.message);
  } finally {
    generateBtn.disabled = false;
    generateBtn.textContent = "Create spreadsheet";
  }
}

// ---------- Preview ----------

function columnLetter(n) {
  let s = "";
  for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function cell(tag, text, className) {
  const el = document.createElement(tag);
  el.textContent = text;
  if (className) el.className = className;
  return el;
}

function renderTable(sheet) {
  const table = $("preview");
  table.replaceChildren();

  const letters = document.createElement("tr");
  letters.append(cell("th", "", "rownum"));
  sheet.columns.forEach((_, i) => letters.append(cell("th", columnLetter(i), "rownum")));

  const head = document.createElement("tr");
  head.append(cell("th", "1", "rownum"));
  sheet.columns.forEach((c) => head.append(cell("th", c.header)));
  table.append(letters, head);

  sheet.rows.forEach((row, r) => {
    const tr = document.createElement("tr");
    tr.append(cell("td", String(r + 2), "rownum"));
    sheet.columns.forEach((_, c) => {
      const v = row[c];
      const isFormula = typeof v === "string" && v.startsWith("=");
      tr.append(cell("td", v == null ? "" : String(v), isFormula ? "formula" : typeof v === "number" ? "num" : ""));
    });
    table.append(tr);
  });

  if (sheet.totals && sheet.rows.length) {
    const tr = document.createElement("tr");
    tr.className = "totals";
    tr.append(cell("td", String(sheet.rows.length + 2), "rownum"));
    const last = sheet.rows.length + 1;
    sheet.columns.forEach((_, c) => {
      const t = sheet.totals.columns.find((x) => x.column === c);
      if (t) tr.append(cell("td", `=${t.fn}(${columnLetter(c)}2:${columnLetter(c)}${last})`, "formula"));
      else tr.append(cell("td", c === 0 ? sheet.totals.label : ""));
    });
    table.append(tr);
  }
}

function render() {
  $("result").hidden = false;
  $("refine-wrap").hidden = false;
  $("result-title").textContent = `${spec.fileName}.xlsx`;
  $("result-summary").textContent = spec.summary;

  const tabs = $("tabs");
  tabs.replaceChildren();
  spec.sheets.forEach((s, i) => {
    const b = cell("button", s.name);
    b.type = "button";
    b.setAttribute("role", "tab");
    b.setAttribute("aria-selected", String(i === activeSheet));
    b.addEventListener("click", () => {
      activeSheet = i;
      render();
    });
    tabs.append(b);
  });
  if (spec.sheets[activeSheet]) renderTable(spec.sheets[activeSheet]);

  $("history-wrap").hidden = history.length === 0;
  $("history").replaceChildren(...history.map((h) => cell("li", h)));
  generateBtn.textContent = "Create spreadsheet";
}

async function download() {
  if (!spec) return;
  const btn = $("download");
  btn.disabled = true;
  try {
    const res = await fetch("/api/xlsx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ spec }),
    });
    if (!res.ok) throw new Error((await res.json().catch(() => ({}))).error || "Download failed");
    const url = URL.createObjectURL(await res.blob());
    const a = Object.assign(document.createElement("a"), { href: url, download: `${spec.fileName || "workbook"}.xlsx` });
    document.body.append(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    showError(err.message);
  } finally {
    btn.disabled = false;
  }
}

generateBtn.addEventListener("click", generate);
$("download").addEventListener("click", download);
$("clear").addEventListener("click", () => {
  promptEl.value = "";
  baseText = "";
  promptEl.focus();
});
promptEl.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) generate();
});
