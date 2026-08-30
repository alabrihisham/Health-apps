import React, { useState, useEffect, useRef, useMemo } from "react";
import {
  Camera, Upload, Loader2, Check, X, Pencil, Trash2, Plus,
  MapPin, ClipboardList, BarChart3, AlertCircle,
  Stethoscope, ChevronRight, ChevronDown, ScanLine,
  FileSpreadsheet, FileText, Search, Mic, Square, Barcode, File as FileIcon, Images, Download, Share2, FileDown
} from "lucide-react";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip,
  BarChart, Bar, XAxis, YAxis, CartesianGrid, LineChart, Line
} from "recharts";
import * as XLSX from "xlsx";
import mammoth from "mammoth";

const CATEGORIES = [
  "General Surgery", "Orthopedics", "Cardiothoracic", "Neurosurgery",
  "Urology", "ENT", "Plastics", "Vascular", "OB/GYN", "Pediatric", "Other"
];

const CATEGORY_COLORS = {
  "General Surgery": "#2B4C43",
  "Orthopedics": "#C1502E",
  "Cardiothoracic": "#4A6FA5",
  "Neurosurgery": "#6B4E71",
  "Urology": "#B08968",
  "ENT": "#8FA998",
  "Plastics": "#A4907C",
  "Vascular": "#3E5C50",
  "OB/GYN": "#9C6644",
  "Pediatric": "#5C6B66",
  "Other": "#8C8577"
};

const LATERALITIES = ["Left", "Right", "Bilateral", "N/A"];
const ROLES = ["Primary Surgeon", "First Assist", "Assistant", "Supervised/Teaching", "Observed"];
const URGENCIES = ["Elective", "Urgent", "Emergency"];
const URGENCY_COLORS = { Elective: "#8FA998", Urgent: "#B08968", Emergency: "#C1502E" };
const SEXES = ["M", "F", "N/A"];
const ANAESTHESIA_TYPES = ["General", "Regional", "Spinal", "Local", "Topical", "Sedation", "N/A"];
const STORAGE_KEY = "surgical-cases";

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function monthKey(dateStr) {
  if (!dateStr) return "Undated";
  return dateStr.slice(0, 7);
}

function monthLabel(key) {
  if (key === "Undated") return "Undated";
  const [y, m] = key.split("-");
  const d = new Date(Number(y), Number(m) - 1, 1);
  return d.toLocaleDateString(undefined, { month: "short", year: "2-digit" });
}

function fmtDate(dateStr) {
  if (!dateStr) return "No date";
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
}

function blankCase() {
  return {
    id: uid(),
    date: todayStr(),
    patientRef: "",
    procedure: "",
    category: "General Surgery",
    surgeon: "",
    location: "",
    laterality: "N/A",
    role: "Primary Surgeon",
    urgency: "Elective",
    patientAge: "",
    patientSex: "N/A",
    anaesthesiaType: "N/A",
    notes: "",
    photos: {}
  };
}

async function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(",")[1]);
    reader.onerror = () => reject(new Error("Could not read the image file."));
    reader.readAsDataURL(file);
  });
}

const PHOTO_STAGES = [
  { key: "preOp", label: "Pre-op" },
  { key: "intraOp", label: "Intra-op" },
  { key: "postOp", label: "Post-op" }
];

function photoKey(caseId, stage, photoId) {
  return `photo:${caseId}:${stage}:${photoId}`;
}

function dataUrlToFile(dataUrl, filename) {
  const [header, base64] = dataUrl.split(",");
  const mimeMatch = header.match(/data:(.*);base64/);
  const mime = mimeMatch ? mimeMatch[1] : "image/jpeg";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new File([bytes], filename, { type: mime });
}

function downloadDataUrl(dataUrl, filename) {
  const a = document.createElement("a");
  a.href = dataUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

const EXPORT_HEADERS = ["Date", "Patient Ref", "Age", "Sex", "Procedure", "Surgeon", "Location", "Laterality", "Role", "Urgency", "Anaesthesia", "Notes"];

function caseToRow(c) {
  return [fmtDate(c.date), c.patientRef, c.patientAge, c.patientSex, c.procedure, c.surgeon, c.location, c.laterality, c.role, c.urgency, c.anaesthesiaType, c.notes];
}

function renderTableHtml(headers, rows) {
  let t = `<table><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>`;
  rows.forEach((r) => {
    t += `<tr>${r.map((v) => `<td>${escapeHtml(v)}</td>`).join("")}</tr>`;
  });
  t += `</table>`;
  return t;
}

function renderCaseTableHtml(rows) {
  return renderTableHtml(EXPORT_HEADERS, rows.map(caseToRow));
}

function buildSummaryHtml(caseList, statsObj) {
  let html = `<h2>Summary</h2>`;
  html += `<p>${caseList.length} case${caseList.length === 1 ? "" : "s"} total.</p>`;
  html += `<h3>By specialty</h3>` + renderTableHtml(["Specialty", "Count"], statsObj.categoryData.map((d) => [d.name, d.value]));
  html += `<h3>By month</h3>` + renderTableHtml(["Month", "Count"], statsObj.monthData.map((d) => [d.label, d.value]));
  html += `<h3>By location</h3>` + renderTableHtml(["Location", "Count"], statsObj.locationData.map((d) => [d.name, d.value]));
  if (statsObj.surgeonData.length > 1) {
    html += `<h3>By surgeon</h3>` + renderTableHtml(["Surgeon", "Count"], statsObj.surgeonData.map((d) => [d.name, d.value]));
  }
  if (statsObj.roleData.length > 1) {
    html += `<h3>By role</h3>` + renderTableHtml(["Role", "Count"], statsObj.roleData.map((d) => [d.name, d.value]));
  }
  if (statsObj.urgencyData.length > 1) {
    html += `<h3>By urgency</h3>` + renderTableHtml(["Urgency", "Count"], statsObj.urgencyData.map((d) => [d.name, d.value]));
  }
  const nonNaLaterality = statsObj.lateralityData.filter((d) => d.name !== "N/A");
  if (nonNaLaterality.length > 1) {
    html += `<h3>By laterality</h3>` + renderTableHtml(["Laterality", "Count"], statsObj.lateralityData.map((d) => [d.name, d.value]));
  }
  const nonNaAnaesthesia = (statsObj.anaesthesiaData || []).filter((d) => d.name !== "N/A");
  if (nonNaAnaesthesia.length > 1) {
    html += `<h3>By anaesthesia type</h3>` + renderTableHtml(["Anaesthesia", "Count"], statsObj.anaesthesiaData.map((d) => [d.name, d.value]));
  }
  const nonNaSex = (statsObj.sexData || []).filter((d) => d.name !== "N/A");
  if (nonNaSex.length > 1) {
    html += `<h3>By patient sex</h3>` + renderTableHtml(["Sex", "Count"], statsObj.sexData.map((d) => [d.name, d.value]));
  }
  return html;
}

function buildReportHtml(caseList, statsObj) {
  let html = `<h1>Surgical Case Log</h1>`;
  html += `<p>Generated ${escapeHtml(new Date().toLocaleString())}</p>`;
  html += buildSummaryHtml(caseList, statsObj);
  CATEGORIES.forEach((cat) => {
    const rows = caseList.filter((c) => c.category === cat);
    if (!rows.length) return;
    html += `<h2>${escapeHtml(cat)} (${rows.length})</h2>` + renderCaseTableHtml(rows);
  });
  return html;
}

function buildSlidesHtml(caseList, statsObj) {
  let html = `<div class="slide" style="display:flex;flex-direction:column;justify-content:center;height:100%;">
    <h1>Surgical Case Log</h1>
    <p style="font-size:18px;">${caseList.length} case${caseList.length === 1 ? "" : "s"} · Generated ${escapeHtml(new Date().toLocaleDateString())}</p>
  </div>`;
  html += `<div class="slide">${buildSummaryHtml(caseList, statsObj)}</div>`;
  CATEGORIES.forEach((cat) => {
    const rows = caseList.filter((c) => c.category === cat);
    if (!rows.length) return;
    html += `<div class="slide"><h2>${escapeHtml(cat)} — ${rows.length} case${rows.length === 1 ? "" : "s"}</h2>${renderCaseTableHtml(rows)}</div>`;
  });
  return html;
}

function openPrintWindow(bodyHtml, { title, landscape }) {
  const win = window.open("", "_blank");
  if (!win) throw new Error("Please allow pop-ups for this page, then try again.");
  win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
    @page { size: ${landscape ? "landscape" : "portrait"}; margin: ${landscape ? "12mm" : "18mm"}; }
    * { box-sizing: border-box; }
    body { font-family: Arial, Helvetica, sans-serif; color: #1E2624; margin: 0; padding: ${landscape ? "0" : "0 0 30px"}; }
    h1 { font-family: Georgia, 'Times New Roman', serif; margin-bottom: 4px; }
    h2 { margin-top: 26px; border-bottom: 2px solid #2B4C43; padding-bottom: 5px; }
    h3 { margin-top: 16px; margin-bottom: 6px; color: #5C6B66; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 16px; font-size: 12px; }
    th, td { border: 1px solid #ccc; padding: 5px 7px; text-align: left; }
    th { background: #F0EEE6; }
    .slide { page-break-after: always; padding: ${landscape ? "30px 46px" : "0"}; }
    .slide:last-child { page-break-after: auto; }
  </style></head><body>${bodyHtml}</body></html>`);
  win.document.close();
  win.focus();
  setTimeout(() => {
    win.print();
  }, 350);
}

function exportToExcel(caseList, statsObj) {
  const wb = XLSX.utils.book_new();

  const summaryRows = [
    ["Surgical Case Log — Summary"],
    ["Generated", new Date().toLocaleString()],
    ["Total cases", caseList.length],
    [],
    ["By specialty"], ["Specialty", "Count"], ...statsObj.categoryData.map((d) => [d.name, d.value]),
    [],
    ["By month"], ["Month", "Count"], ...statsObj.monthData.map((d) => [d.label, d.value]),
    [],
    ["By location"], ["Location", "Count"], ...statsObj.locationData.map((d) => [d.name, d.value]),
    [],
    ["By role"], ["Role", "Count"], ...statsObj.roleData.map((d) => [d.name, d.value]),
    [],
    ["By urgency"], ["Urgency", "Count"], ...statsObj.urgencyData.map((d) => [d.name, d.value]),
    [],
    ["By laterality"], ["Laterality", "Count"], ...statsObj.lateralityData.map((d) => [d.name, d.value]),
    [],
    ["By anaesthesia type"], ["Anaesthesia", "Count"], ...statsObj.anaesthesiaData.map((d) => [d.name, d.value]),
    [],
    ["By patient sex"], ["Sex", "Count"], ...statsObj.sexData.map((d) => [d.name, d.value])
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryRows), "Summary");

  const header = ["Date", "Patient Ref", "Age", "Sex", "Procedure", "Specialty", "Surgeon", "Location", "Laterality", "Role", "Urgency", "Anaesthesia", "Notes"];
  const toRow = (c) => [c.date, c.patientRef, c.patientAge, c.patientSex, c.procedure, c.category, c.surgeon, c.location, c.laterality, c.role, c.urgency, c.anaesthesiaType, c.notes];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([header, ...caseList.map(toRow)]), "All Cases");

  const usedNames = new Set(["Summary", "All Cases"]);
  CATEGORIES.forEach((cat) => {
    const rows = caseList.filter((c) => c.category === cat);
    if (!rows.length) return;
    let name = cat.replace(/[\\/?*[\]:]/g, "-").slice(0, 31);
    while (usedNames.has(name)) name = (name.slice(0, 28) + "-2");
    usedNames.add(name);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([header, ...rows.map(toRow)]), name);
  });

  XLSX.writeFile(wb, `surgical-case-log-${todayStr()}.xlsx`);
}

function exportToWord(caseList, statsObj) {
  const body = buildReportHtml(caseList, statsObj);
  const docHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>Surgical Case Log</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
  body { font-family: Calibri, Arial, sans-serif; color: #1E2624; }
  h1 { font-family: Georgia, 'Times New Roman', serif; }
  h2 { border-bottom: 2px solid #2B4C43; padding-bottom: 4px; margin-top: 24px; }
  h3 { color: #5C6B66; margin-bottom: 4px; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 14px; font-size: 12px; }
  th, td { border: 1px solid #ccc; padding: 5px 7px; text-align: left; }
  th { background: #F0EEE6; }
</style>
</head><body>${body}</body></html>`;
  const blob = new Blob(["\ufeff", docHtml], { type: "application/msword" });
  downloadBlob(blob, `surgical-case-log-${todayStr()}.doc`);
}

async function compressImageFile(file, maxDim = 1000, quality = 0.72) {
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Could not read that image."));
    reader.readAsDataURL(file);
  });
  const img = await new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Could not open that image."));
    image.src = dataUrl;
  });
  let { width, height } = img;
  if (width > maxDim || height > maxDim) {
    if (width >= height) {
      height = Math.round((height * maxDim) / width);
      width = maxDim;
    } else {
      width = Math.round((width * maxDim) / height);
      height = maxDim;
    }
  }
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0, width, height);
  return canvas.toDataURL("image/jpeg", quality);
}

const FIELD_SPEC = `- "date": best-guess date in YYYY-MM-DD format, or "" if none is visible
- "patientRef": a short DE-IDENTIFIED reference only, such as initials, a case/patient number, or MRN as given. NEVER write out a full patient name, even if one appears in the source — use initials or the patient number instead.
- "patientAge": the patient's age as given (e.g. "45" or "45Y"), or "" if not given
- "patientSex": one of "M", "F", "N/A"
- "procedure": the procedure or operation name
- "category": your best single classification (specialty), chosen exactly from this list: ${CATEGORIES.join(", ")}
- "surgeon": surgeon or consultant name or initials, or "" if not given
- "location": hospital, site, ward, or OR/theatre room, or "" if not given
- "laterality": exactly one of "Left", "Right", "Bilateral", "N/A"
- "role": the role of the person logging the case, chosen exactly from: ${ROLES.join(", ")}. Default to "Primary Surgeon" if nothing suggests otherwise.
- "urgency": chosen exactly from: ${URGENCIES.join(", ")}. Default to "Elective" if nothing suggests otherwise (an "Emergency" list heading or similar means "Emergency").
- "anaesthesiaType": chosen exactly from: ${ANAESTHESIA_TYPES.join(", ")}. Use "N/A" if not given.
- "notes": any brief extra detail such as diagnosis, indication, or complication, or ""`;

const RETURN_SPEC = `Return ONLY a raw JSON array, nothing else. No markdown code fences, no commentary before or after.`;


function normalizeExtractedRows(parsed) {
  if (!Array.isArray(parsed)) return [];
  return parsed.map((c) => ({
    id: uid(),
    date: typeof c.date === "string" ? c.date : "",
    patientRef: typeof c.patientRef === "string" ? c.patientRef : "",
    patientAge: typeof c.patientAge === "string" ? c.patientAge : (typeof c.patientAge === "number" ? String(c.patientAge) : ""),
    patientSex: SEXES.includes(c.patientSex) ? c.patientSex : "N/A",
    procedure: typeof c.procedure === "string" ? c.procedure : "",
    category: CATEGORIES.includes(c.category) ? c.category : "Other",
    surgeon: typeof c.surgeon === "string" ? c.surgeon : "",
    location: typeof c.location === "string" ? c.location : "",
    laterality: LATERALITIES.includes(c.laterality) ? c.laterality : "N/A",
    role: ROLES.includes(c.role) ? c.role : "Primary Surgeon",
    urgency: URGENCIES.includes(c.urgency) ? c.urgency : "Elective",
    anaesthesiaType: ANAESTHESIA_TYPES.includes(c.anaesthesiaType) ? c.anaesthesiaType : "N/A",
    notes: typeof c.notes === "string" ? c.notes : "",
    photos: {}
  }));
}

function parseJsonRows(rawText, notUnderstoodMessage) {
  let cleaned = rawText.trim();
  cleaned = cleaned.replace(/^```(json)?/i, "").replace(/```$/, "").trim();
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error(notUnderstoodMessage);
  }
  if (!Array.isArray(parsed) || parsed.length === 0) {
    throw new Error(notUnderstoodMessage);
  }
  return normalizeExtractedRows(parsed);
}

async function callClaude(contentBlocks) {
  // Calls our own serverless proxy (api/claude.js) instead of Anthropic directly,
  // so the API key never has to live in browser code.
  const response = await fetch("/api/claude", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      messages: [{ role: "user", content: contentBlocks }]
    })
  });
  if (!response.ok) {
    throw new Error("The extraction service did not respond. Please try again.");
  }
  const data = await response.json();
  const textBlock = (data.content || []).find((b) => b.type === "text");
  if (!textBlock) throw new Error("No readable data came back.");
  return textBlock.text;
}

async function extractCasesFromImage(base64, mediaType) {
  const prompt = `You are reading a photo of a printed or handwritten surgical case list, OR schedule, or logbook page. Find every distinct surgical case on it and return them as a JSON array.

For each case, return an object with exactly these fields:
${FIELD_SPEC}

${RETURN_SPEC}`;

  const text = await callClaude([
    { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
    { type: "text", text: prompt }
  ]);

  return parseJsonRows(text, "Couldn't understand the list in that photo. Try a clearer, flatter shot.");
}

const MAX_PDF_BYTES = 20 * 1024 * 1024;

async function extractCasesFromPdf(base64) {
  const prompt = `You are reading a PDF containing a surgical case list, OR schedule, or logbook. It may be a typed document, a scanned page, or a table, and may run to several pages. Find every distinct surgical case in it and return them as a JSON array.

For each case, return an object with exactly these fields:
${FIELD_SPEC}

${RETURN_SPEC}`;

  const text = await callClaude([
    { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } },
    { type: "text", text: prompt }
  ]);

  return parseJsonRows(text, "Couldn't find a case list in that PDF. If it's a scan, try a clearer copy.");
}

const MAX_TEXT_CHARS = 14000;

async function extractCasesFromText(rawText, sourceKind) {
  let content = rawText;
  let truncated = false;
  if (content.length > MAX_TEXT_CHARS) {
    content = content.slice(0, MAX_TEXT_CHARS);
    truncated = true;
  }

  const sourceDescription =
    sourceKind === "spreadsheet"
      ? "a spreadsheet (given below as CSV) containing a surgical case list or logbook. The first row is likely a header row — column names may differ from the field names below (for example a column called \"Op\", \"Theatre\", or \"Consultant\") — map them sensibly by meaning, not by exact name."
      : sourceKind === "dictation"
      ? "a spoken voice transcript of a surgeon dictating one or more case notes out loud. It may include filler words (\"um\", \"so\", \"next patient\"), run-on sentences, and casual phrasing — ignore filler and split it into separate cases wherever the speaker clearly moves to a new one (cues like \"next case\", \"case two\", \"and then\", or a new patient/procedure being named)."
      : "a Word document containing a surgical case list, OR schedule, or logbook. It may be a table or a series of free-text lines, one case per line or paragraph.";

  const prompt = `You are reading ${sourceDescription} Find every distinct surgical case in it and return them as a JSON array.

For each case, return an object with exactly these fields:
${FIELD_SPEC}

${truncated ? "Note: this content was cut off because it was very long, so only work with what is shown below.\n\n" : ""}${RETURN_SPEC}

CONTENT:
"""
${content}
"""`;

  const text = await callClaude([{ type: "text", text: prompt }]);

  const notUnderstood =
    sourceKind === "spreadsheet"
      ? "Couldn't find case data in that spreadsheet. Check it has one row per case."
      : sourceKind === "dictation"
      ? "Couldn't make out any cases in that recording. Try speaking each detail clearly, one case at a time."
      : "Couldn't find a case list in that document.";

  return parseJsonRows(text, notUnderstood);
}

async function fileToCsvText(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_csv(sheet);
}

async function fileToDocxText(file) {
  const buffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer: buffer });
  return result.value;
}

function StitchDivider({ style }) {
  return (
    <svg viewBox="0 0 400 10" preserveAspectRatio="none" style={{ width: "100%", height: 10, display: "block", ...style }}>
      {Array.from({ length: 26 }).map((_, i) => (
        <line
          key={i}
          x1={i * 16 + 2}
          y1={2}
          x2={i * 16 + 10}
          y2={8}
          stroke="#C7C2B4"
          strokeWidth="1.6"
          strokeLinecap="round"
        />
      ))}
    </svg>
  );
}

function Pill({ children, color }) {
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        padding: "3px 9px",
        borderRadius: 999,
        fontSize: 11.5,
        fontWeight: 600,
        letterSpacing: 0.2,
        color: "#fff",
        background: color || "#5C6B66",
        whiteSpace: "nowrap"
      }}
    >
      {children}
    </span>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "flex", flexDirection: "column", gap: 5, fontSize: 12.5, color: "#5C6B66", fontWeight: 600 }}>
      {label}
      {children}
    </label>
  );
}

const inputStyle = {
  border: "1px solid #DDD8CB",
  borderRadius: 8,
  padding: "9px 10px",
  fontSize: 14.5,
  fontFamily: "'IBM Plex Sans', sans-serif",
  color: "#1E2624",
  background: "#fff",
  outline: "none"
};

const BARCODE_FORMATS = [
  "code_128", "code_39", "code_93", "codabar", "ean_13", "ean_8",
  "itf", "upc_a", "upc_e", "qr_code", "data_matrix", "pdf417"
];

function BarcodeScanButton({ onScan }) {
  const supported = typeof window !== "undefined" && "BarcodeDetector" in window;
  const [open, setOpen] = useState(false);
  const [error, setError] = useState(null);
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const rafRef = useRef(null);
  const detectorRef = useRef(null);

  const stopCamera = () => {
    if (rafRef.current) cancelAnimationFrame(rafRef.current);
    rafRef.current = null;
    if (streamRef.current) {
      streamRef.current.getTracks().forEach((t) => t.stop());
      streamRef.current = null;
    }
  };

  useEffect(() => () => stopCamera(), []);

  const closeScanner = () => {
    stopCamera();
    setOpen(false);
    setError(null);
  };

  const scanLoop = () => {
    const tick = async () => {
      if (!videoRef.current || videoRef.current.readyState < 2) {
        rafRef.current = requestAnimationFrame(tick);
        return;
      }
      try {
        const codes = await detectorRef.current.detect(videoRef.current);
        if (codes && codes.length > 0) {
          const value = codes[0].rawValue;
          stopCamera();
          setOpen(false);
          setError(null);
          onScan(value);
          return;
        }
      } catch (e) {
        // transient detection miss — keep trying
      }
      rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
  };

  const openScanner = async () => {
    setError(null);
    setOpen(true);
    try {
      detectorRef.current = new window.BarcodeDetector({ formats: BARCODE_FORMATS });
    } catch (e) {
      detectorRef.current = new window.BarcodeDetector();
    }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
      streamRef.current = stream;
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();
      }
      scanLoop();
    } catch (e) {
      setError("Couldn't access the camera. Check permissions, or type the ID in by hand.");
    }
  };

  if (!supported) return null;

  return (
    <>
      <button
        type="button"
        onClick={openScanner}
        aria-label="Scan barcode"
        style={{ ...iconBtn, width: "auto", padding: "0 10px", gap: 5, display: "inline-flex", flexShrink: 0 }}
      >
        <Barcode size={14} /> <span style={{ fontSize: 11.5, fontWeight: 700 }}>Scan</span>
      </button>
      {open && (
        <div style={{ position: "fixed", inset: 0, background: "#000", zIndex: 1000, display: "flex", flexDirection: "column" }}>
          <video ref={videoRef} muted playsInline style={{ flex: 1, width: "100%", objectFit: "cover" }} />
          <div style={{ position: "absolute", top: 0, left: 0, right: 0, padding: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ color: "#fff", fontSize: 12.5, fontWeight: 600, background: "rgba(0,0,0,0.45)", padding: "6px 10px", borderRadius: 8 }}>
              Point at the patient's barcode
            </span>
            <button
              onClick={closeScanner}
              aria-label="Cancel scan"
              style={{ background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 8, width: 34, height: 34, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}
            >
              <X size={18} />
            </button>
          </div>
          <div style={{ position: "absolute", left: "12%", right: "12%", top: "38%", height: 90, border: "2px solid #C1502E", borderRadius: 10, pointerEvents: "none" }} />
          {error && (
            <div style={{ position: "absolute", bottom: 24, left: 16, right: 16, background: "#FBEDE7", color: "#8A3A20", borderRadius: 9, padding: "10px 12px", fontSize: 13, textAlign: "center" }}>
              {error}
            </div>
          )}
        </div>
      )}
    </>
  );
}

function CaseForm({ initial, onCancel, onSave, saveLabel }) {
  const [form, setForm] = useState(initial);
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Date">
          <input type="date" style={inputStyle} value={form.date} onChange={set("date")} />
        </Field>
        <Field label="Patient ref. (initials/ID)">
          <div style={{ display: "flex", gap: 6 }}>
            <input style={{ ...inputStyle, flex: 1, minWidth: 0 }} value={form.patientRef} onChange={set("patientRef")} placeholder="e.g. J.D. or #4471" />
            <BarcodeScanButton onScan={(value) => setForm((f) => ({ ...f, patientRef: value }))} />
          </div>
        </Field>
      </div>
      <Field label="Procedure">
        <input style={inputStyle} value={form.procedure} onChange={set("procedure")} placeholder="e.g. Laparoscopic cholecystectomy" />
      </Field>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
        <Field label="Age">
          <input style={inputStyle} value={form.patientAge} onChange={set("patientAge")} placeholder="e.g. 45" />
        </Field>
        <Field label="Sex">
          <select style={inputStyle} value={form.patientSex} onChange={set("patientSex")}>
            {SEXES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Anaesthesia">
          <select style={inputStyle} value={form.anaesthesiaType} onChange={set("anaesthesiaType")}>
            {ANAESTHESIA_TYPES.map((a) => <option key={a} value={a}>{a}</option>)}
          </select>
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Category">
          <select style={inputStyle} value={form.category} onChange={set("category")}>
            {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </Field>
        <Field label="Laterality">
          <select style={inputStyle} value={form.laterality} onChange={set("laterality")}>
            {LATERALITIES.map((l) => <option key={l} value={l}>{l}</option>)}
          </select>
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Surgeon">
          <input style={inputStyle} value={form.surgeon} onChange={set("surgeon")} placeholder="e.g. Dr. Alavi" />
        </Field>
        <Field label="Location / site">
          <input style={inputStyle} value={form.location} onChange={set("location")} placeholder="e.g. St. Mary's OR 3" />
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Your role">
          <select style={inputStyle} value={form.role} onChange={set("role")}>
            {ROLES.map((r) => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="Urgency">
          <select style={inputStyle} value={form.urgency} onChange={set("urgency")}>
            {URGENCIES.map((u) => <option key={u} value={u}>{u}</option>)}
          </select>
        </Field>
      </div>
      <Field label="Notes (optional)">
        <input style={inputStyle} value={form.notes} onChange={set("notes")} placeholder="Indication, complications..." />
      </Field>
      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 4 }}>
        <button onClick={onCancel} style={btnGhost}>Cancel</button>
        <button
          onClick={() => onSave(form)}
          disabled={!form.procedure.trim()}
          style={{ ...btnPrimary, opacity: form.procedure.trim() ? 1 : 0.5 }}
        >
          {saveLabel || "Save case"}
        </button>
      </div>
    </div>
  );
}

const btnPrimary = {
  background: "#2B4C43",
  color: "#fff",
  border: "none",
  borderRadius: 9,
  padding: "10px 16px",
  fontSize: 14,
  fontWeight: 700,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  gap: 7
};

const btnGhost = {
  background: "transparent",
  color: "#5C6B66",
  border: "1px solid #DDD8CB",
  borderRadius: 9,
  padding: "10px 16px",
  fontSize: 14,
  fontWeight: 600,
  cursor: "pointer"
};

const btnAccent = {
  ...btnPrimary,
  background: "#C1502E"
};

const exportCardStyle = {
  display: "flex",
  alignItems: "center",
  gap: 12,
  width: "100%",
  textAlign: "left",
  background: "#fff",
  border: "1px solid #EAE5D8",
  borderRadius: 12,
  padding: "13px 14px",
  cursor: "pointer",
  fontFamily: "'IBM Plex Sans', sans-serif"
};

const miniIconBtn = {
  background: "rgba(30,38,36,0.62)",
  border: "none",
  borderRadius: 5,
  width: 19,
  height: 19,
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  color: "#fff",
  cursor: "pointer",
  padding: 0
};

function PhotoThumb({ photo, url, onOpen }) {
  return (
    <button
      type="button"
      onClick={onOpen}
      style={{
        position: "relative",
        aspectRatio: "1",
        borderRadius: 7,
        overflow: "hidden",
        border: "1px solid #EAE5D8",
        background: "#F7F4EE",
        padding: 0,
        cursor: "pointer",
        display: "flex",
        alignItems: "center",
        justifyContent: "center"
      }}
    >
      {url ? (
        <img src={url} alt="Case photo" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
      ) : (
        <Loader2 size={14} color="#8C8577" style={{ animation: "spin 1s linear infinite" }} />
      )}
    </button>
  );
}

function PhotoStageGrid({ label, photos, photoCache, caseId, stageKey, saving, error, onPickFiles, onOpen }) {
  const inputRef = useRef(null);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 5, flex: 1, minWidth: 0 }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: "#8C8577", textTransform: "uppercase", letterSpacing: 0.4 }}>
        {label}{photos.length > 0 ? ` (${photos.length})` : ""}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 4 }}>
        {photos.map((p, idx) => (
          <PhotoThumb key={p.id} photo={p} url={photoCache[photoKey(caseId, stageKey, p.id)]} onOpen={() => onOpen(idx)} />
        ))}
        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          style={{
            aspectRatio: "1",
            borderRadius: 7,
            border: "1px dashed #C7C2B4",
            background: "#F7F4EE",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            cursor: "pointer",
            color: "#8C8577",
            padding: 0
          }}
        >
          {saving ? <Loader2 size={14} style={{ animation: "spin 1s linear infinite" }} /> : <Plus size={16} />}
        </button>
      </div>
      {error && <div style={{ fontSize: 10, color: "#C1502E" }}>{error}</div>}
      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        multiple
        style={{ display: "none" }}
        onChange={(e) => {
          const files = Array.from(e.target.files || []);
          if (files.length) onPickFiles(files);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function PhotoViewer({ stageLabel, photos, photoCache, caseId, stageKey, index, onClose, onNavigate, onDateChange, onDelete }) {
  const shareSupported = typeof navigator !== "undefined" && !!navigator.share;
  const [actionError, setActionError] = useState(null);
  if (index == null || !photos[index]) return null;
  const p = photos[index];
  const url = photoCache[photoKey(caseId, stageKey, p.id)];
  const filename = `${stageKey}-${p.date || "photo"}-${index + 1}.jpg`;

  const handleDownload = () => {
    if (!url) return;
    downloadDataUrl(url, filename);
  };

  const handleShare = async () => {
    if (!url) return;
    setActionError(null);
    try {
      const file = dataUrlToFile(url, filename);
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: `${stageLabel} photo` });
      } else if (navigator.share) {
        await navigator.share({ title: `${stageLabel} photo` });
      } else {
        setActionError("Sharing isn't supported on this device. Use Download instead.");
      }
    } catch (e) {
      if (e && e.name !== "AbortError") {
        setActionError("Couldn't share that photo. Try Download instead.");
      }
    }
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(10,12,11,0.95)", zIndex: 1100, display: "flex", flexDirection: "column" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "14px 16px" }}>
        <span style={{ color: "#fff", fontSize: 13, fontWeight: 600 }}>
          {stageLabel} · {index + 1} of {photos.length}
        </span>
        <button onClick={onClose} aria-label="Close" style={{ background: "rgba(255,255,255,0.15)", border: "none", borderRadius: 8, width: 32, height: 32, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
          <X size={17} />
        </button>
      </div>

      <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", position: "relative", overflow: "hidden", padding: "0 8px" }}>
        {photos.length > 1 && (
          <button
            onClick={() => onNavigate(index === 0 ? photos.length - 1 : index - 1)}
            aria-label="Previous photo"
            style={{ position: "absolute", left: 6, background: "rgba(255,255,255,0.15)", border: "none", borderRadius: "50%", width: 34, height: 34, color: "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}
          >
            <ChevronRight size={18} style={{ transform: "rotate(180deg)" }} />
          </button>
        )}
        {url ? (
          <img src={url} alt={`${stageLabel} photo`} style={{ maxWidth: "100%", maxHeight: "100%", objectFit: "contain", borderRadius: 6 }} />
        ) : (
          <Loader2 size={28} color="#fff" style={{ animation: "spin 1s linear infinite" }} />
        )}
        {photos.length > 1 && (
          <button
            onClick={() => onNavigate(index === photos.length - 1 ? 0 : index + 1)}
            aria-label="Next photo"
            style={{ position: "absolute", right: 6, background: "rgba(255,255,255,0.15)", border: "none", borderRadius: "50%", width: 34, height: 34, color: "#fff", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}
          >
            <ChevronRight size={18} />
          </button>
        )}
      </div>

      <div style={{ padding: "12px 16px 20px", display: "flex", flexDirection: "column", gap: 10 }}>
        {actionError && (
          <div style={{ background: "rgba(251,237,231,0.95)", color: "#8A3A20", borderRadius: 9, padding: "8px 12px", fontSize: 12.5, textAlign: "center" }}>
            {actionError}
          </div>
        )}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
          <span style={{ color: "#D8D4C8", fontSize: 12 }}>Taken</span>
          <input
            type="date"
            value={p.date || ""}
            onChange={(e) => onDateChange(index, e.target.value)}
            style={{ ...inputStyle, padding: "6px 8px", fontSize: 12.5 }}
          />
        </div>
        <div style={{ display: "flex", gap: 8, justifyContent: "center", flexWrap: "wrap" }}>
          <button onClick={handleDownload} style={{ ...btnGhost, background: "rgba(255,255,255,0.1)", color: "#fff", borderColor: "rgba(255,255,255,0.25)" }}>
            <Download size={15} /> Save
          </button>
          {shareSupported && (
            <button onClick={handleShare} style={btnPrimary}>
              <Share2 size={15} /> Share
            </button>
          )}
          <button onClick={() => onDelete(index)} style={{ ...btnGhost, background: "rgba(193,80,46,0.15)", color: "#F0AF97", borderColor: "rgba(193,80,46,0.4)" }}>
            <Trash2 size={15} /> Delete
          </button>
        </div>
      </div>
    </div>
  );
}

function CaseCard({ c, onEdit, onDelete, confirmingDelete, onConfirmDelete, onCancelDelete, photoCache, onLoadPhoto, onSavePhoto, onDeletePhoto, onUpdatePhotoDate }) {
  const [photosOpen, setPhotosOpen] = useState(false);
  const [savingStage, setSavingStage] = useState(null);
  const [stageErrors, setStageErrors] = useState({});
  const [viewer, setViewer] = useState(null); // { stageKey, index } | null
  const photos = c.photos || {};
  const photoCount = PHOTO_STAGES.reduce((n, s) => n + ((photos[s.key] || []).length), 0);

  useEffect(() => {
    if (!photosOpen) return;
    PHOTO_STAGES.forEach((s) => {
      (photos[s.key] || []).forEach((p) => {
        if (photoCache[photoKey(c.id, s.key, p.id)] === undefined) {
          onLoadPhoto(s.key, p.id);
        }
      });
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [photosOpen, c.id, photos, photoCache]);

  const handlePickFiles = async (stage, files) => {
    setSavingStage(stage);
    setStageErrors((e) => ({ ...e, [stage]: null }));
    try {
      for (const file of files) {
        await onSavePhoto(stage, file);
      }
    } catch (err) {
      setStageErrors((e) => ({ ...e, [stage]: err.message || "Couldn't save one of those photos." }));
    } finally {
      setSavingStage(null);
    }
  };

  const handleViewerDelete = async (stageKey, index) => {
    const list = photos[stageKey] || [];
    const p = list[index];
    if (!p) return;
    await onDeletePhoto(stageKey, p.id);
    const remaining = list.length - 1;
    if (remaining <= 0) setViewer(null);
    else setViewer({ stageKey, index: Math.min(index, remaining - 1) });
  };

  return (
    <div
      style={{
        background: "#fff",
        border: "1px solid #EAE5D8",
        borderRadius: 12,
        padding: "13px 14px",
        display: "flex",
        flexDirection: "column",
        gap: 6
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
          <span style={{ fontSize: 11.5, color: "#8C8577", fontFamily: "'IBM Plex Mono', monospace" }}>
            {fmtDate(c.date)}{c.patientRef ? ` · ${c.patientRef}` : ""}{(c.patientAge || (c.patientSex && c.patientSex !== "N/A")) ? ` · ${c.patientAge || ""}${c.patientAge && c.patientSex !== "N/A" ? " " : ""}${c.patientSex !== "N/A" ? c.patientSex : ""}` : ""}
          </span>
          <span style={{ fontSize: 15.5, fontWeight: 700, color: "#1E2624", fontFamily: "'Source Serif 4', serif" }}>
            {c.procedure || "Untitled procedure"}
          </span>
        </div>
        {!confirmingDelete ? (
          <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
            <button onClick={onEdit} style={iconBtn} aria-label="Edit case"><Pencil size={15} /></button>
            <button onClick={onDelete} style={iconBtn} aria-label="Delete case"><Trash2 size={15} /></button>
          </div>
        ) : (
          <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
            <button onClick={onConfirmDelete} style={{ ...iconBtn, color: "#C1502E" }} aria-label="Confirm delete"><Check size={15} /></button>
            <button onClick={onCancelDelete} style={iconBtn} aria-label="Cancel delete"><X size={15} /></button>
          </div>
        )}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginTop: 2 }}>
        <Pill color={CATEGORY_COLORS[c.category]}>{c.category}</Pill>
        {c.laterality !== "N/A" && <Pill color="#8C8577">{c.laterality}</Pill>}
        {c.urgency && c.urgency !== "Elective" && <Pill color={URGENCY_COLORS[c.urgency] || "#8C8577"}>{c.urgency}</Pill>}
        {c.role && c.role !== "Primary Surgeon" && <Pill color="#A4907C">{c.role}</Pill>}
        {c.anaesthesiaType && c.anaesthesiaType !== "N/A" && <Pill color="#4A6FA5">{c.anaesthesiaType}</Pill>}
      </div>
      {(c.surgeon || c.location) && (
        <div style={{ fontSize: 12.5, color: "#5C6B66", display: "flex", gap: 12, flexWrap: "wrap" }}>
          {c.surgeon && <span>👤 {c.surgeon}</span>}
          {c.location && <span><MapPin size={12} style={{ verticalAlign: -1 }} /> {c.location}</span>}
        </div>
      )}
      {c.notes && <div style={{ fontSize: 12.5, color: "#8C8577", fontStyle: "italic" }}>{c.notes}</div>}

      <button
        type="button"
        onClick={() => setPhotosOpen((v) => !v)}
        style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "none", cursor: "pointer", padding: "4px 0 0", color: "#5C6B66", fontSize: 12, fontWeight: 600, alignSelf: "flex-start" }}
      >
        <Images size={13} /> Photos{photoCount > 0 ? ` (${photoCount})` : ""}
        {photosOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
      </button>

      {photosOpen && (
        <div style={{ display: "flex", gap: 8, marginTop: 2 }}>
          {PHOTO_STAGES.map((s) => (
            <PhotoStageGrid
              key={s.key}
              label={s.label}
              photos={photos[s.key] || []}
              photoCache={photoCache}
              caseId={c.id}
              stageKey={s.key}
              saving={savingStage === s.key}
              error={stageErrors[s.key]}
              onPickFiles={(files) => handlePickFiles(s.key, files)}
              onOpen={(idx) => setViewer({ stageKey: s.key, index: idx })}
            />
          ))}
        </div>
      )}

      {viewer && (
        <PhotoViewer
          stageLabel={PHOTO_STAGES.find((s) => s.key === viewer.stageKey)?.label || ""}
          photos={photos[viewer.stageKey] || []}
          photoCache={photoCache}
          caseId={c.id}
          stageKey={viewer.stageKey}
          index={viewer.index}
          onClose={() => setViewer(null)}
          onNavigate={(idx) => setViewer((v) => ({ ...v, index: idx }))}
          onDateChange={(idx, date) => {
            const p = (photos[viewer.stageKey] || [])[idx];
            if (p) onUpdatePhotoDate(viewer.stageKey, p.id, date);
          }}
          onDelete={(idx) => handleViewerDelete(viewer.stageKey, idx)}
        />
      )}
    </div>
  );
}

const iconBtn = {
  background: "#F7F4EE",
  border: "1px solid #EAE5D8",
  borderRadius: 7,
  width: 28,
  height: 28,
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  color: "#5C6B66",
  cursor: "pointer"
};

export default function SurgicalCaseLog() {
  const [cases, setCases] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("capture");
  const [editingId, setEditingId] = useState(null);
  const [deletingId, setDeletingId] = useState(null);
  const [manualOpen, setManualOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [statFilters, setStatFilters] = useState({
    category: "All",
    role: "All",
    urgency: "All",
    laterality: "All",
    anaesthesiaType: "All",
    location: "",
    dateFrom: "",
    dateTo: ""
  });
  const [filtersOpen, setFiltersOpen] = useState(false);
  const [photoCache, setPhotoCache] = useState({});

  const [imgPreview, setImgPreview] = useState(null);
  const [sourceFileName, setSourceFileName] = useState(null);
  const [extracting, setExtracting] = useState(false);
  const [extractingLabel, setExtractingLabel] = useState("Reading the list…");
  const [extractError, setExtractError] = useState(null);
  const [reviewRows, setReviewRows] = useState(null);

  const [dictationSupported] = useState(
    () => typeof window !== "undefined" && !!(window.SpeechRecognition || window.webkitSpeechRecognition)
  );
  const [dictating, setDictating] = useState(false);
  const [liveTranscript, setLiveTranscript] = useState("");
  const [finalTranscript, setFinalTranscript] = useState("");
  const recognitionRef = useRef(null);
  const finalTranscriptRef = useRef("");

  const fileInputCamera = useRef(null);
  const fileInputUpload = useRef(null);
  const fileInputSpreadsheet = useRef(null);
  const fileInputWord = useRef(null);
  const fileInputPdf = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (res && res.value) setCases(JSON.parse(res.value));
      } catch (e) {
        // no saved cases yet
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  const casesRef = useRef(cases);
  useEffect(() => {
    casesRef.current = cases;
  }, [cases]);

  const persist = async (next) => {
    casesRef.current = next;
    setCases(next);
    try {
      await window.storage.set(STORAGE_KEY, JSON.stringify(next), false);
    } catch (e) {
      // storage failed silently; data still lives in this session
    }
  };

  const loadPhotoIfNeeded = async (caseId, stage, photoId) => {
    const key = photoKey(caseId, stage, photoId);
    if (photoCache[key] !== undefined) return;
    try {
      const res = await window.storage.get(key, false);
      const value = res && res.value ? res.value : null;
      setPhotoCache((c) => (c[key] !== undefined ? c : { ...c, [key]: value }));
    } catch (e) {
      setPhotoCache((c) => (c[key] !== undefined ? c : { ...c, [key]: null }));
    }
  };

  const savePhoto = async (caseId, stage, file) => {
    const photoId = uid();
    const key = photoKey(caseId, stage, photoId);
    const compressed = await compressImageFile(file);
    const result = await window.storage.set(key, compressed, false);
    if (!result) throw new Error("Couldn't save that photo. Try again.");
    setPhotoCache((c) => ({ ...c, [key]: compressed }));
    const newEntry = { id: photoId, date: todayStr() };
    const next = casesRef.current.map((cs) =>
      cs.id === caseId
        ? { ...cs, photos: { ...(cs.photos || {}), [stage]: [...((cs.photos && cs.photos[stage]) || []), newEntry] } }
        : cs
    );
    await persist(next);
  };

  const deletePhoto = async (caseId, stage, photoId) => {
    const key = photoKey(caseId, stage, photoId);
    try {
      await window.storage.delete(key, false);
    } catch (e) {
      // key may already be gone; nothing to do
    }
    setPhotoCache((c) => {
      const next = { ...c };
      delete next[key];
      return next;
    });
    const next = casesRef.current.map((cs) => {
      if (cs.id !== caseId) return cs;
      const list = ((cs.photos && cs.photos[stage]) || []).filter((p) => p.id !== photoId);
      return { ...cs, photos: { ...(cs.photos || {}), [stage]: list } };
    });
    await persist(next);
  };

  const updatePhotoDate = async (caseId, stage, photoId, date) => {
    const next = casesRef.current.map((cs) => {
      if (cs.id !== caseId) return cs;
      const list = ((cs.photos && cs.photos[stage]) || []).map((p) => (p.id === photoId ? { ...p, date } : p));
      return { ...cs, photos: { ...(cs.photos || {}), [stage]: list } };
    });
    await persist(next);
  };

  const handleImageFile = async (file) => {
    if (!file) return;
    setExtractError(null);
    setReviewRows(null);
    setSourceFileName(null);
    const url = URL.createObjectURL(file);
    setImgPreview(url);
    setExtractingLabel("Reading the list…");
    setExtracting(true);
    try {
      const base64 = await fileToBase64(file);
      const rows = await extractCasesFromImage(base64, file.type || "image/jpeg");
      setReviewRows(rows.map((r) => ({ ...r, included: true })));
    } catch (e) {
      setExtractError(e.message || "Something went wrong reading that photo.");
    } finally {
      setExtracting(false);
    }
  };

  const handleSpreadsheetFile = async (file) => {
    if (!file) return;
    setExtractError(null);
    setReviewRows(null);
    setImgPreview(null);
    setSourceFileName(file.name);
    setExtractingLabel("Reading the spreadsheet…");
    setExtracting(true);
    try {
      const csv = await fileToCsvText(file);
      if (!csv || !csv.trim()) throw new Error("That spreadsheet looks empty.");
      const rows = await extractCasesFromText(csv, "spreadsheet");
      setReviewRows(rows.map((r) => ({ ...r, included: true })));
    } catch (e) {
      setExtractError(e.message || "Something went wrong reading that spreadsheet.");
    } finally {
      setExtracting(false);
    }
  };

  const handleWordFile = async (file) => {
    if (!file) return;
    setExtractError(null);
    setReviewRows(null);
    setImgPreview(null);
    setSourceFileName(file.name);
    setExtractingLabel("Reading the document…");
    setExtracting(true);
    try {
      const text = await fileToDocxText(file);
      if (!text || !text.trim()) throw new Error("That document looks empty.");
      const rows = await extractCasesFromText(text, "document");
      setReviewRows(rows.map((r) => ({ ...r, included: true })));
    } catch (e) {
      setExtractError(e.message || "Something went wrong reading that document.");
    } finally {
      setExtracting(false);
    }
  };

  const handlePdfFile = async (file) => {
    if (!file) return;
    setExtractError(null);
    setReviewRows(null);
    setImgPreview(null);
    setSourceFileName(file.name);
    setExtractingLabel("Reading the PDF…");
    setExtracting(true);
    try {
      if (file.size > MAX_PDF_BYTES) {
        throw new Error("That PDF is quite large (over 20MB) — try exporting a smaller range of pages.");
      }
      const base64 = await fileToBase64(file);
      const rows = await extractCasesFromPdf(base64);
      setReviewRows(rows.map((r) => ({ ...r, included: true })));
    } catch (e) {
      setExtractError(e.message || "Something went wrong reading that PDF.");
    } finally {
      setExtracting(false);
    }
  };

  const startDictation = () => {
    if (!dictationSupported) return;
    setExtractError(null);
    setReviewRows(null);
    setImgPreview(null);
    setSourceFileName(null);
    setFinalTranscript("");
    setLiveTranscript("");
    finalTranscriptRef.current = "";

    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    const recognition = new SR();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = "en-US";

    recognition.onresult = (event) => {
      let interim = "";
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const res = event.results[i];
        if (res.isFinal) {
          finalTranscriptRef.current += res[0].transcript + " ";
        } else {
          interim += res[0].transcript;
        }
      }
      setLiveTranscript(finalTranscriptRef.current + interim);
    };

    recognition.onerror = (event) => {
      if (event.error === "not-allowed" || event.error === "service-not-allowed") {
        setExtractError("Microphone access was blocked. Check your browser's mic permission and try again.");
      } else if (event.error !== "no-speech" && event.error !== "aborted") {
        setExtractError("Voice dictation stopped unexpectedly. Please try again.");
      }
    };

    recognition.onend = () => {
      setDictating(false);
      setFinalTranscript(finalTranscriptRef.current.trim());
    };

    recognitionRef.current = recognition;
    try {
      recognition.start();
      setDictating(true);
    } catch (e) {
      setExtractError("Couldn't start the microphone. Please try again.");
    }
  };

  const stopDictation = () => {
    recognitionRef.current?.stop();
  };

  const cancelDictation = () => {
    recognitionRef.current?.stop();
    recognitionRef.current = null;
    setDictating(false);
    setFinalTranscript("");
    setLiveTranscript("");
    finalTranscriptRef.current = "";
  };

  const extractFromDictation = async () => {
    const text = finalTranscript.trim();
    setExtractError(null);
    setReviewRows(null);
    setExtractingLabel("Making sense of what you said…");
    setExtracting(true);
    try {
      if (!text) throw new Error("Didn't catch any speech to work with. Try dictating again.");
      const rows = await extractCasesFromText(text, "dictation");
      setReviewRows(rows.map((r) => ({ ...r, included: true })));
      setFinalTranscript("");
      setLiveTranscript("");
    } catch (e) {
      setExtractError(e.message || "Something went wrong understanding that.");
    } finally {
      setExtracting(false);
    }
  };

  const updateReviewRow = (id, patch) => {
    setReviewRows((rows) => rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  };

  const saveReviewed = async () => {
    const toSave = reviewRows.filter((r) => r.included).map(({ included, ...rest }) => rest);
    await persist([...toSave, ...cases]);
    setReviewRows(null);
    setImgPreview(null);
    setTab("log");
  };

  const discardReview = () => {
    setReviewRows(null);
    setImgPreview(null);
    setSourceFileName(null);
    setExtractError(null);
  };

  const addManual = async (form) => {
    await persist([{ ...form }, ...cases]);
    setManualOpen(false);
    setTab("log");
  };

  const saveEdit = async (form) => {
    await persist(cases.map((c) => (c.id === form.id ? form : c)));
    setEditingId(null);
  };

  const doDelete = async (id) => {
    await persist(cases.filter((c) => c.id !== id));
    setDeletingId(null);
  };

  const filteredCases = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return cases;
    return cases.filter((c) => {
      const haystack = [
        c.date, fmtDate(c.date), c.patientRef, c.procedure, c.category,
        c.surgeon, c.location, c.laterality, c.role, c.urgency,
        c.patientAge, c.patientSex, c.anaesthesiaType, c.notes
      ].join(" ").toLowerCase();
      return haystack.includes(q);
    });
  }, [cases, searchQuery]);

  const statsActiveCount = useMemo(() => {
    let n = 0;
    if (statFilters.category !== "All") n++;
    if (statFilters.role !== "All") n++;
    if (statFilters.urgency !== "All") n++;
    if (statFilters.laterality !== "All") n++;
    if (statFilters.anaesthesiaType !== "All") n++;
    if (statFilters.location.trim()) n++;
    if (statFilters.dateFrom) n++;
    if (statFilters.dateTo) n++;
    return n;
  }, [statFilters]);

  const statsScopedCases = useMemo(() => {
    return cases.filter((c) => {
      if (statFilters.category !== "All" && c.category !== statFilters.category) return false;
      if (statFilters.role !== "All" && c.role !== statFilters.role) return false;
      if (statFilters.urgency !== "All" && c.urgency !== statFilters.urgency) return false;
      if (statFilters.laterality !== "All" && c.laterality !== statFilters.laterality) return false;
      if (statFilters.anaesthesiaType !== "All" && c.anaesthesiaType !== statFilters.anaesthesiaType) return false;
      if (statFilters.location.trim()) {
        if (!c.location.toLowerCase().includes(statFilters.location.trim().toLowerCase())) return false;
      }
      if (statFilters.dateFrom) {
        if (!c.date || c.date < statFilters.dateFrom) return false;
      }
      if (statFilters.dateTo) {
        if (!c.date || c.date > statFilters.dateTo) return false;
      }
      return true;
    });
  }, [cases, statFilters]);

  const resetStatFilters = () =>
    setStatFilters({ category: "All", role: "All", urgency: "All", laterality: "All", anaesthesiaType: "All", location: "", dateFrom: "", dateTo: "" });

  const stats = useMemo(() => {
    const byCategory = {};
    const byMonth = {};
    const byLocation = {};
    const bySurgeon = {};
    const byLaterality = {};
    const byRole = {};
    const byUrgency = {};
    const byAnaesthesia = {};
    const bySex = {};
    statsScopedCases.forEach((c) => {
      byCategory[c.category] = (byCategory[c.category] || 0) + 1;
      const mk = monthKey(c.date);
      byMonth[mk] = (byMonth[mk] || 0) + 1;
      const loc = c.location.trim() || "Unspecified";
      byLocation[loc] = (byLocation[loc] || 0) + 1;
      const sg = c.surgeon.trim() || "Unspecified";
      bySurgeon[sg] = (bySurgeon[sg] || 0) + 1;
      byLaterality[c.laterality] = (byLaterality[c.laterality] || 0) + 1;
      byRole[c.role || "Primary Surgeon"] = (byRole[c.role || "Primary Surgeon"] || 0) + 1;
      byUrgency[c.urgency || "Elective"] = (byUrgency[c.urgency || "Elective"] || 0) + 1;
      byAnaesthesia[c.anaesthesiaType || "N/A"] = (byAnaesthesia[c.anaesthesiaType || "N/A"] || 0) + 1;
      bySex[c.patientSex || "N/A"] = (bySex[c.patientSex || "N/A"] || 0) + 1;
    });

    const categoryData = Object.entries(byCategory)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const monthData = Object.entries(byMonth)
      .sort((a, b) => (a[0] > b[0] ? 1 : -1))
      .map(([key, value]) => ({ key, label: monthLabel(key), value }));

    const locationData = Object.entries(byLocation)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const surgeonData = Object.entries(bySurgeon)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8);

    const roleData = Object.entries(byRole)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const urgencyData = Object.entries(byUrgency)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const lateralityData = Object.entries(byLaterality)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const anaesthesiaData = Object.entries(byAnaesthesia)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const sexData = Object.entries(bySex)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const now = new Date();
    const thisMonthKey = now.toISOString().slice(0, 7);
    const thisMonthCount = byMonth[thisMonthKey] || 0;
    const topCategory = categoryData[0]?.name || "—";
    const topLocation = locationData[0]?.name || "—";

    return {
      categoryData, monthData, locationData, surgeonData, roleData, urgencyData, lateralityData,
      anaesthesiaData, sexData,
      thisMonthCount, topCategory, topLocation
    };
  }, [statsScopedCases]);

  const maxLocation = Math.max(1, ...stats.locationData.map((d) => d.value));

  const [exportError, setExportError] = useState(null);
  const [exportBusy, setExportBusy] = useState(null);

  const runExport = async (kind, fn) => {
    setExportError(null);
    setExportBusy(kind);
    try {
      await Promise.resolve(fn());
    } catch (e) {
      setExportError(e.message || "That export didn't work. Please try again.");
    } finally {
      setExportBusy(null);
    }
  };

  return (
    <div style={{
      fontFamily: "'IBM Plex Sans', sans-serif",
      background: "#F7F4EE",
      minHeight: "100%",
      color: "#1E2624",
      padding: "18px 14px 40px",
      maxWidth: 560,
      margin: "0 auto",
      boxSizing: "border-box"
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:wght@600;700&family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        button:focus-visible, input:focus-visible, select:focus-visible {
          outline: 2px solid #2B4C43;
          outline-offset: 2px;
        }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
      `}</style>

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
        <div style={{
          width: 38, height: 38, borderRadius: 10, background: "#2B4C43",
          display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0
        }}>
          <ScanLine color="#F7F4EE" size={20} />
        </div>
        <div>
          <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 700, fontSize: 21, lineHeight: 1.1 }}>SutureLog</div>
          <div style={{ fontSize: 12, color: "#8C8577" }}>{cases.length} case{cases.length === 1 ? "" : "s"} logged</div>
        </div>
      </div>

      <StitchDivider style={{ margin: "12px 0 16px" }} />

      <div style={{ display: "flex", gap: 6, marginBottom: 18, background: "#EFEBDF", padding: 4, borderRadius: 11 }}>
        {[
          { id: "capture", label: "Add", icon: Camera },
          { id: "log", label: "Log", icon: ClipboardList },
          { id: "stats", label: "Stats", icon: BarChart3 },
          { id: "export", label: "Export", icon: FileDown }
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            style={{
              flex: 1,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              gap: 6,
              padding: "9px 0",
              borderRadius: 8,
              border: "none",
              fontSize: 13.5,
              fontWeight: 700,
              cursor: "pointer",
              background: tab === t.id ? "#fff" : "transparent",
              color: tab === t.id ? "#2B4C43" : "#8C8577",
              boxShadow: tab === t.id ? "0 1px 3px rgba(0,0,0,0.08)" : "none"
            }}
          >
            <t.icon size={15} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "capture" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
          {!reviewRows && (
            <div style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 14, padding: 20, textAlign: "center" }}>
              <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 700, fontSize: 17, marginBottom: 4 }}>
                Bring in today's list
              </div>
              <div style={{ fontSize: 13, color: "#8C8577", marginBottom: 16, lineHeight: 1.5 }}>
                Photograph a schedule, or import an Excel/CSV sheet, Word document, or PDF — dates, procedures, and surgeons get filled in for you to check.
              </div>

              {imgPreview && (
                <img src={imgPreview} alt="Selected surgical list" style={{ width: "100%", borderRadius: 10, marginBottom: 14, maxHeight: 220, objectFit: "cover" }} />
              )}

              {sourceFileName && !imgPreview && (
                <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#F7F4EE", border: "1px solid #EAE5D8", borderRadius: 9, padding: "10px 12px", marginBottom: 14, fontSize: 13, color: "#5C6B66", textAlign: "left" }}>
                  <FileSpreadsheet size={16} style={{ flexShrink: 0 }} />
                  <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sourceFileName}</span>
                </div>
              )}

              {extracting && (
                <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, color: "#5C6B66", fontSize: 13.5, padding: "14px 0" }}>
                  <Loader2 size={16} className="spin" style={{ animation: "spin 1s linear infinite" }} />
                  {extractingLabel}
                </div>
              )}

              {extractError && (
                <div style={{ display: "flex", gap: 8, alignItems: "flex-start", background: "#FBEDE7", border: "1px solid #EFC9B8", color: "#8A3A20", borderRadius: 9, padding: "10px 12px", fontSize: 13, textAlign: "left", marginBottom: 12 }}>
                  <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
                  <span>{extractError}</span>
                </div>
              )}

              {!extracting && (
                <>
                  {!dictating && !finalTranscript && (
                    <>
                      <div style={{ display: "flex", gap: 10, justifyContent: "center", flexWrap: "wrap" }}>
                        <button style={btnPrimary} onClick={() => fileInputCamera.current?.click()}>
                          <Camera size={16} /> Take photo
                        </button>
                        <button style={btnGhost} onClick={() => fileInputUpload.current?.click()}>
                          <Upload size={16} /> Upload photo
                        </button>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "16px 0 12px", color: "#B7B1A0", fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
                        <span style={{ flex: 1, height: 1, background: "#EAE5D8" }} />
                        or import a file
                        <span style={{ flex: 1, height: 1, background: "#EAE5D8" }} />
                      </div>
                      <div style={{ display: "flex", gap: 10, justifyContent: "center", flexWrap: "wrap" }}>
                        <button style={btnGhost} onClick={() => fileInputSpreadsheet.current?.click()}>
                          <FileSpreadsheet size={16} /> Excel / CSV
                        </button>
                        <button style={btnGhost} onClick={() => fileInputWord.current?.click()}>
                          <FileText size={16} /> Word doc
                        </button>
                        <button style={btnGhost} onClick={() => fileInputPdf.current?.click()}>
                          <FileIcon size={16} /> PDF
                        </button>
                      </div>
                      {dictationSupported && (
                        <>
                          <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "16px 0 12px", color: "#B7B1A0", fontSize: 11.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5 }}>
                            <span style={{ flex: 1, height: 1, background: "#EAE5D8" }} />
                            or speak it
                            <span style={{ flex: 1, height: 1, background: "#EAE5D8" }} />
                          </div>
                          <div style={{ display: "flex", justifyContent: "center" }}>
                            <button style={btnGhost} onClick={startDictation}>
                              <Mic size={16} /> Dictate a case
                            </button>
                          </div>
                        </>
                      )}
                    </>
                  )}

                  {dictating && (
                    <div style={{ textAlign: "left" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, justifyContent: "center" }}>
                        <span style={{ width: 9, height: 9, borderRadius: "50%", background: "#C1502E", animation: "pulseDot 1.2s ease-in-out infinite" }} />
                        <span style={{ fontSize: 13, fontWeight: 700, color: "#C1502E" }}>Listening…</span>
                      </div>
                      <div style={{ minHeight: 70, maxHeight: 160, overflowY: "auto", background: "#F7F4EE", border: "1px solid #EAE5D8", borderRadius: 9, padding: "10px 12px", fontSize: 13.5, color: liveTranscript ? "#1E2624" : "#8C8577", marginBottom: 12, lineHeight: 1.5 }}>
                        {liveTranscript || "Say something like: \"Laparoscopic appendectomy, patient J.D., general surgery, emergency case, at St. Mary's.\""}
                      </div>
                      <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
                        <button style={btnGhost} onClick={cancelDictation}><X size={16} /> Cancel</button>
                        <button style={btnAccent} onClick={stopDictation}><Square size={14} /> Stop</button>
                      </div>
                    </div>
                  )}

                  {!dictating && finalTranscript && (
                    <div style={{ textAlign: "left" }}>
                      <div style={{ fontSize: 12.5, fontWeight: 600, color: "#5C6B66", marginBottom: 6 }}>
                        Here's what we heard — fix anything before filling in details:
                      </div>
                      <textarea
                        value={finalTranscript}
                        onChange={(e) => setFinalTranscript(e.target.value)}
                        rows={4}
                        style={{ ...inputStyle, width: "100%", resize: "vertical", marginBottom: 12, fontFamily: "'IBM Plex Sans', sans-serif" }}
                      />
                      <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
                        <button style={btnGhost} onClick={cancelDictation}>Discard</button>
                        <button style={btnPrimary} onClick={extractFromDictation}><Check size={16} /> Fill in details</button>
                      </div>
                    </div>
                  )}
                </>
              )}
              <input ref={fileInputCamera} type="file" accept="image/*" capture="environment" style={{ display: "none" }}
                onChange={(e) => handleImageFile(e.target.files?.[0])} />
              <input ref={fileInputUpload} type="file" accept="image/*" style={{ display: "none" }}
                onChange={(e) => handleImageFile(e.target.files?.[0])} />
              <input ref={fileInputSpreadsheet} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
                onChange={(e) => handleSpreadsheetFile(e.target.files?.[0])} />
              <input ref={fileInputWord} type="file" accept=".docx" style={{ display: "none" }}
                onChange={(e) => handleWordFile(e.target.files?.[0])} />
              <input ref={fileInputPdf} type="file" accept=".pdf,application/pdf" style={{ display: "none" }}
                onChange={(e) => handlePdfFile(e.target.files?.[0])} />
            </div>
          )}

          {reviewRows && (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div style={{ fontSize: 13.5, color: "#5C6B66" }}>
                Found <strong>{reviewRows.length}</strong> case{reviewRows.length === 1 ? "" : "s"}. Check them over, then save.
              </div>
              {reviewRows.map((r) => (
                <div key={r.id} style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 12, padding: 12, opacity: r.included ? 1 : 0.5 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5, fontWeight: 600, color: "#5C6B66" }}>
                      <input type="checkbox" checked={r.included} onChange={(e) => updateReviewRow(r.id, { included: e.target.checked })} />
                      Include this case
                    </label>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                    <input style={inputStyle} type="date" value={r.date} onChange={(e) => updateReviewRow(r.id, { date: e.target.value })} />
                    <div style={{ display: "flex", gap: 6 }}>
                      <input style={{ ...inputStyle, flex: 1, minWidth: 0 }} value={r.patientRef} placeholder="Patient ref." onChange={(e) => updateReviewRow(r.id, { patientRef: e.target.value })} />
                      <BarcodeScanButton onScan={(value) => updateReviewRow(r.id, { patientRef: value })} />
                    </div>
                  </div>
                  <input style={{ ...inputStyle, width: "100%", marginBottom: 8 }} value={r.procedure} placeholder="Procedure" onChange={(e) => updateReviewRow(r.id, { procedure: e.target.value })} />
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 8 }}>
                    <input style={inputStyle} value={r.patientAge} placeholder="Age" onChange={(e) => updateReviewRow(r.id, { patientAge: e.target.value })} />
                    <select style={inputStyle} value={r.patientSex} onChange={(e) => updateReviewRow(r.id, { patientSex: e.target.value })}>
                      {SEXES.map((s) => <option key={s} value={s}>{s}</option>)}
                    </select>
                    <select style={inputStyle} value={r.anaesthesiaType} onChange={(e) => updateReviewRow(r.id, { anaesthesiaType: e.target.value })}>
                      {ANAESTHESIA_TYPES.map((a) => <option key={a} value={a}>{a}</option>)}
                    </select>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                    <select style={inputStyle} value={r.category} onChange={(e) => updateReviewRow(r.id, { category: e.target.value })}>
                      {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                    <select style={inputStyle} value={r.laterality} onChange={(e) => updateReviewRow(r.id, { laterality: e.target.value })}>
                      {LATERALITIES.map((l) => <option key={l} value={l}>{l}</option>)}
                    </select>
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                    <input style={inputStyle} value={r.surgeon} placeholder="Surgeon" onChange={(e) => updateReviewRow(r.id, { surgeon: e.target.value })} />
                    <input style={inputStyle} value={r.location} placeholder="Location" onChange={(e) => updateReviewRow(r.id, { location: e.target.value })} />
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <select style={inputStyle} value={r.role} onChange={(e) => updateReviewRow(r.id, { role: e.target.value })}>
                      {ROLES.map((rl) => <option key={rl} value={rl}>{rl}</option>)}
                    </select>
                    <select style={inputStyle} value={r.urgency} onChange={(e) => updateReviewRow(r.id, { urgency: e.target.value })}>
                      {URGENCIES.map((u) => <option key={u} value={u}>{u}</option>)}
                    </select>
                  </div>
                </div>
              ))}
              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 4 }}>
                <button style={btnGhost} onClick={discardReview}>Discard</button>
                <button style={btnPrimary} onClick={saveReviewed}>
                  <Check size={16} /> Save {reviewRows.filter((r) => r.included).length} case{reviewRows.filter((r) => r.included).length === 1 ? "" : "s"}
                </button>
              </div>
            </div>
          )}

          {!reviewRows && !extracting && (
            <div style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 14, padding: 16 }}>
              <button
                onClick={() => setManualOpen((v) => !v)}
                style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", background: "none", border: "none", cursor: "pointer", padding: 0 }}
              >
                <span style={{ fontWeight: 700, fontSize: 14.5, display: "flex", alignItems: "center", gap: 7 }}>
                  <Plus size={16} /> Add a case by hand
                </span>
                {manualOpen ? <ChevronDown size={17} color="#8C8577" /> : <ChevronRight size={17} color="#8C8577" />}
              </button>
              {manualOpen && (
                <div style={{ marginTop: 14 }}>
                  <CaseForm initial={blankCase()} onCancel={() => setManualOpen(false)} onSave={addManual} saveLabel="Add case" />
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {tab === "log" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {cases.length > 0 && (
            <div style={{ position: "relative" }}>
              <Search size={15} color="#8C8577" style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)" }} />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Search procedure, surgeon, site, date, notes…"
                style={{ ...inputStyle, width: "100%", paddingLeft: 32, paddingRight: searchQuery ? 32 : 10 }}
              />
              {searchQuery && (
                <button
                  onClick={() => setSearchQuery("")}
                  aria-label="Clear search"
                  style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", color: "#8C8577", cursor: "pointer", padding: 4, display: "flex" }}
                >
                  <X size={15} />
                </button>
              )}
            </div>
          )}

          {searchQuery && cases.length > 0 && (
            <div style={{ fontSize: 12, color: "#8C8577" }}>
              {filteredCases.length} result{filteredCases.length === 1 ? "" : "s"} for "{searchQuery}"
            </div>
          )}

          {!loaded ? (
            <div style={{ textAlign: "center", color: "#8C8577", fontSize: 13, padding: "30px 0" }}>Loading your log…</div>
          ) : cases.length === 0 ? (
            <div style={{ textAlign: "center", padding: "36px 16px", color: "#8C8577" }}>
              <Stethoscope size={26} style={{ marginBottom: 8, opacity: 0.6 }} />
              <div style={{ fontSize: 14, fontWeight: 600, color: "#5C6B66", marginBottom: 4 }}>No cases yet</div>
              <div style={{ fontSize: 12.5 }}>Snap a photo of today's list to get started, or add one by hand.</div>
            </div>
          ) : filteredCases.length === 0 ? (
            <div style={{ textAlign: "center", padding: "36px 16px", color: "#8C8577" }}>
              <Search size={24} style={{ marginBottom: 8, opacity: 0.6 }} />
              <div style={{ fontSize: 14, fontWeight: 600, color: "#5C6B66", marginBottom: 4 }}>No matches</div>
              <div style={{ fontSize: 12.5 }}>Nothing in the log matches "{searchQuery}".</div>
            </div>
          ) : (
            filteredCases
              .slice()
              .sort((a, b) => (a.date < b.date ? 1 : -1))
              .map((c) =>
                editingId === c.id ? (
                  <div key={c.id} style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 12, padding: 12 }}>
                    <CaseForm initial={c} onCancel={() => setEditingId(null)} onSave={saveEdit} saveLabel="Save changes" />
                  </div>
                ) : (
                  <CaseCard
                    key={c.id}
                    c={c}
                    onEdit={() => setEditingId(c.id)}
                    onDelete={() => setDeletingId(c.id)}
                    confirmingDelete={deletingId === c.id}
                    onConfirmDelete={() => doDelete(c.id)}
                    onCancelDelete={() => setDeletingId(null)}
                    photoCache={photoCache}
                    onLoadPhoto={(stage, photoId) => loadPhotoIfNeeded(c.id, stage, photoId)}
                    onSavePhoto={(stage, file) => savePhoto(c.id, stage, file)}
                    onDeletePhoto={(stage, photoId) => deletePhoto(c.id, stage, photoId)}
                    onUpdatePhotoDate={(stage, photoId, date) => updatePhotoDate(c.id, stage, photoId, date)}
                  />
                )
              )
          )}
        </div>
      )}

      {tab === "stats" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
          {cases.length === 0 ? (
            <div style={{ textAlign: "center", padding: "36px 16px", color: "#8C8577" }}>
              <BarChart3 size={26} style={{ marginBottom: 8, opacity: 0.6 }} />
              <div style={{ fontSize: 14, fontWeight: 600, color: "#5C6B66" }}>Nothing to show yet</div>
              <div style={{ fontSize: 12.5 }}>Log a few cases and your stats will appear here.</div>
            </div>
          ) : (
            <>
              <div style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 14, padding: 14 }}>
                <button
                  onClick={() => setFiltersOpen((v) => !v)}
                  style={{ display: "flex", alignItems: "center", justifyContent: "space-between", width: "100%", background: "none", border: "none", cursor: "pointer", padding: 0 }}
                >
                  <span style={{ fontWeight: 700, fontSize: 14, display: "flex", alignItems: "center", gap: 7 }}>
                    <Search size={15} /> Filter stats
                    {statsActiveCount > 0 && (
                      <span style={{ background: "#2B4C43", color: "#fff", borderRadius: 999, fontSize: 11, fontWeight: 700, padding: "1px 7px" }}>
                        {statsActiveCount}
                      </span>
                    )}
                  </span>
                  {filtersOpen ? <ChevronDown size={17} color="#8C8577" /> : <ChevronRight size={17} color="#8C8577" />}
                </button>
                {filtersOpen && (
                  <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 10 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      <Field label="Specialty">
                        <select style={inputStyle} value={statFilters.category} onChange={(e) => setStatFilters((f) => ({ ...f, category: e.target.value }))}>
                          <option value="All">All specialties</option>
                          {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </Field>
                      <Field label="Your role">
                        <select style={inputStyle} value={statFilters.role} onChange={(e) => setStatFilters((f) => ({ ...f, role: e.target.value }))}>
                          <option value="All">All roles</option>
                          {ROLES.map((r) => <option key={r} value={r}>{r}</option>)}
                        </select>
                      </Field>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      <Field label="Urgency">
                        <select style={inputStyle} value={statFilters.urgency} onChange={(e) => setStatFilters((f) => ({ ...f, urgency: e.target.value }))}>
                          <option value="All">All urgency</option>
                          {URGENCIES.map((u) => <option key={u} value={u}>{u}</option>)}
                        </select>
                      </Field>
                      <Field label="Laterality">
                        <select style={inputStyle} value={statFilters.laterality} onChange={(e) => setStatFilters((f) => ({ ...f, laterality: e.target.value }))}>
                          <option value="All">All</option>
                          {LATERALITIES.map((l) => <option key={l} value={l}>{l}</option>)}
                        </select>
                      </Field>
                    </div>
                    <Field label="Anaesthesia">
                      <select style={inputStyle} value={statFilters.anaesthesiaType} onChange={(e) => setStatFilters((f) => ({ ...f, anaesthesiaType: e.target.value }))}>
                        <option value="All">All anaesthesia types</option>
                        {ANAESTHESIA_TYPES.map((a) => <option key={a} value={a}>{a}</option>)}
                      </select>
                    </Field>
                    <Field label="Location contains">
                      <input style={inputStyle} value={statFilters.location} placeholder="e.g. St. Mary's, Toronto…"
                        onChange={(e) => setStatFilters((f) => ({ ...f, location: e.target.value }))} />
                    </Field>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      <Field label="From date">
                        <input type="date" style={inputStyle} value={statFilters.dateFrom}
                          onChange={(e) => setStatFilters((f) => ({ ...f, dateFrom: e.target.value }))} />
                      </Field>
                      <Field label="To date">
                        <input type="date" style={inputStyle} value={statFilters.dateTo}
                          onChange={(e) => setStatFilters((f) => ({ ...f, dateTo: e.target.value }))} />
                      </Field>
                    </div>
                    {statsActiveCount > 0 && (
                      <button onClick={resetStatFilters} style={{ ...btnGhost, alignSelf: "flex-start", padding: "7px 12px", fontSize: 12.5 }}>
                        Clear filters
                      </button>
                    )}
                  </div>
                )}
              </div>

              {statsScopedCases.length === 0 ? (
                <div style={{ textAlign: "center", padding: "30px 16px", color: "#8C8577" }}>
                  <Search size={22} style={{ marginBottom: 8, opacity: 0.6 }} />
                  <div style={{ fontSize: 14, fontWeight: 600, color: "#5C6B66", marginBottom: 4 }}>No cases match these filters</div>
                  <div style={{ fontSize: 12.5 }}>Try widening the date range or clearing a filter.</div>
                </div>
              ) : (
                <>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                    <SummaryCard label={statsActiveCount > 0 ? "Matching cases" : "Total cases"} value={statsScopedCases.length} />
                    <SummaryCard label="This month" value={stats.thisMonthCount} />
                    <SummaryCard label="Top specialty" value={stats.topCategory} small />
                    <SummaryCard label="Top site" value={stats.topLocation} small />
                  </div>

                  <StatBlock title="By specialty">
                    <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                      <ResponsiveContainer width={140} height={140}>
                        <PieChart>
                          <Pie data={stats.categoryData} dataKey="value" nameKey="name" innerRadius={34} outerRadius={62} paddingAngle={2}>
                            {stats.categoryData.map((d, i) => (
                              <Cell key={i} fill={CATEGORY_COLORS[d.name] || "#8C8577"} />
                            ))}
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 150 }}>
                        {stats.categoryData.map((d) => (
                          <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5 }}>
                            <span style={{ width: 9, height: 9, borderRadius: 3, background: CATEGORY_COLORS[d.name] || "#8C8577", flexShrink: 0 }} />
                            <span style={{ flex: 1, color: "#1E2624" }}>{d.name}</span>
                            <span style={{ color: "#8C8577", fontFamily: "'IBM Plex Mono', monospace" }}>{d.value}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </StatBlock>

                  <StatBlock title="Cases over time">
                    <ResponsiveContainer width="100%" height={170}>
                      <BarChart data={stats.monthData} margin={{ top: 4, right: 4, left: -20, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#EAE5D8" vertical={false} />
                        <XAxis dataKey="label" tick={{ fontSize: 11, fill: "#8C8577" }} axisLine={{ stroke: "#EAE5D8" }} tickLine={false} />
                        <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: "#8C8577" }} axisLine={false} tickLine={false} />
                        <Tooltip />
                        <Bar dataKey="value" fill="#2B4C43" radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </StatBlock>

                  <StatBlock title="By location" icon={MapPin}>
                    <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                      {stats.locationData.map((d) => (
                        <div key={d.name}>
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 3 }}>
                            <span style={{ display: "flex", alignItems: "center", gap: 5 }}><MapPin size={12} color="#8C8577" /> {d.name}</span>
                            <span style={{ color: "#8C8577", fontFamily: "'IBM Plex Mono', monospace" }}>{d.value}</span>
                          </div>
                          <div style={{ background: "#EFEBDF", borderRadius: 6, height: 8, overflow: "hidden" }}>
                            <div style={{ width: `${(d.value / maxLocation) * 100}%`, background: "#C1502E", height: "100%", borderRadius: 6 }} />
                          </div>
                        </div>
                      ))}
                    </div>
                  </StatBlock>

                  {stats.surgeonData.length > 1 && (
                    <StatBlock title="By surgeon">
                      <ResponsiveContainer width="100%" height={Math.max(120, stats.surgeonData.length * 32)}>
                        <BarChart data={stats.surgeonData} layout="vertical" margin={{ top: 0, right: 16, left: 8, bottom: 0 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#EAE5D8" horizontal={false} />
                          <XAxis type="number" allowDecimals={false} tick={{ fontSize: 11, fill: "#8C8577" }} axisLine={false} tickLine={false} />
                          <YAxis type="category" dataKey="name" width={90} tick={{ fontSize: 11.5, fill: "#1E2624" }} axisLine={false} tickLine={false} />
                          <Tooltip />
                          <Bar dataKey="value" fill="#8FA998" radius={[0, 4, 4, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </StatBlock>
                  )}

                  {stats.roleData.length > 1 && (
                    <StatBlock title="By your role">
                      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                        {stats.roleData.map((d) => (
                          <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                            <span style={{ flex: 1 }}>{d.name}</span>
                            <div style={{ flex: 2, background: "#EFEBDF", borderRadius: 6, height: 8, overflow: "hidden" }}>
                              <div style={{ width: `${(d.value / statsScopedCases.length) * 100}%`, background: "#A4907C", height: "100%", borderRadius: 6 }} />
                            </div>
                            <span style={{ color: "#8C8577", fontFamily: "'IBM Plex Mono', monospace", width: 22, textAlign: "right" }}>{d.value}</span>
                          </div>
                        ))}
                      </div>
                    </StatBlock>
                  )}

                  {stats.urgencyData.length > 1 && (
                    <StatBlock title="By urgency">
                      <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                        <ResponsiveContainer width={110} height={110}>
                          <PieChart>
                            <Pie data={stats.urgencyData} dataKey="value" nameKey="name" innerRadius={26} outerRadius={50} paddingAngle={2}>
                              {stats.urgencyData.map((d, i) => (
                                <Cell key={i} fill={URGENCY_COLORS[d.name] || "#8C8577"} />
                              ))}
                            </Pie>
                            <Tooltip />
                          </PieChart>
                        </ResponsiveContainer>
                        <div style={{ display: "flex", flexDirection: "column", gap: 6, flex: 1, minWidth: 130 }}>
                          {stats.urgencyData.map((d) => (
                            <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5 }}>
                              <span style={{ width: 9, height: 9, borderRadius: 3, background: URGENCY_COLORS[d.name] || "#8C8577", flexShrink: 0 }} />
                              <span style={{ flex: 1, color: "#1E2624" }}>{d.name}</span>
                              <span style={{ color: "#8C8577", fontFamily: "'IBM Plex Mono', monospace" }}>{d.value}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    </StatBlock>
                  )}

                  {stats.lateralityData.filter((d) => d.name !== "N/A").length > 1 && (
                    <StatBlock title="By laterality">
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                        {stats.lateralityData.map((d) => (
                          <div key={d.name} style={{ flex: 1, minWidth: 80, background: "#F7F4EE", borderRadius: 10, padding: "10px 8px", textAlign: "center" }}>
                            <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace" }}>{d.value}</div>
                            <div style={{ fontSize: 11.5, color: "#8C8577" }}>{d.name}</div>
                          </div>
                        ))}
                      </div>
                    </StatBlock>
                  )}

                  {stats.anaesthesiaData.filter((d) => d.name !== "N/A").length > 1 && (
                    <StatBlock title="By anaesthesia type">
                      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                        {stats.anaesthesiaData.map((d) => (
                          <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5 }}>
                            <span style={{ flex: 1 }}>{d.name}</span>
                            <div style={{ flex: 2, background: "#EFEBDF", borderRadius: 6, height: 8, overflow: "hidden" }}>
                              <div style={{ width: `${(d.value / statsScopedCases.length) * 100}%`, background: "#4A6FA5", height: "100%", borderRadius: 6 }} />
                            </div>
                            <span style={{ color: "#8C8577", fontFamily: "'IBM Plex Mono', monospace", width: 22, textAlign: "right" }}>{d.value}</span>
                          </div>
                        ))}
                      </div>
                    </StatBlock>
                  )}

                  {stats.sexData.filter((d) => d.name !== "N/A").length > 1 && (
                    <StatBlock title="By patient sex">
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                        {stats.sexData.map((d) => (
                          <div key={d.name} style={{ flex: 1, minWidth: 80, background: "#F7F4EE", borderRadius: 10, padding: "10px 8px", textAlign: "center" }}>
                            <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "'IBM Plex Mono', monospace" }}>{d.value}</div>
                            <div style={{ fontSize: 11.5, color: "#8C8577" }}>{d.name}</div>
                          </div>
                        ))}
                      </div>
                    </StatBlock>
                  )}
                </>
              )}
            </>
          )}
        </div>
      )}

      {tab === "export" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {cases.length === 0 ? (
            <div style={{ textAlign: "center", padding: "36px 16px", color: "#8C8577" }}>
              <FileDown size={26} style={{ marginBottom: 8, opacity: 0.6 }} />
              <div style={{ fontSize: 14, fontWeight: 600, color: "#5C6B66" }}>Nothing to export yet</div>
              <div style={{ fontSize: 12.5 }}>Log a few cases first, then come back here.</div>
            </div>
          ) : (
            <>
              <div style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 14, padding: 14 }}>
                <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 700, fontSize: 15, marginBottom: 4 }}>
                  Export your log
                </div>
                <div style={{ fontSize: 12.5, color: "#8C8577", lineHeight: 1.5 }}>
                  {statsActiveCount > 0
                    ? `Exporting the ${statsScopedCases.length} case${statsScopedCases.length === 1 ? "" : "s"} matching your Stats filters. Clear filters in the Stats tab to export everything.`
                    : `Exporting all ${cases.length} case${cases.length === 1 ? "" : "s"}. Set filters in the Stats tab first if you only want a subset.`}
                  {" "}Every export is organized by specialty, with a full summary up front.
                </div>
              </div>

              {exportError && (
                <div style={{ display: "flex", gap: 8, alignItems: "flex-start", background: "#FBEDE7", border: "1px solid #EFC9B8", color: "#8A3A20", borderRadius: 9, padding: "10px 12px", fontSize: 13 }}>
                  <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
                  <span>{exportError}</span>
                </div>
              )}

              {statsScopedCases.length === 0 ? (
                <div style={{ textAlign: "center", padding: "20px 16px", color: "#8C8577", fontSize: 13 }}>
                  No cases match the current Stats filters, so there's nothing to export right now.
                </div>
              ) : (
                <>
                  <button
                    onClick={() => runExport("excel", () => exportToExcel(statsScopedCases, stats))}
                    disabled={exportBusy !== null}
                    style={{ ...exportCardStyle }}
                  >
                    <FileSpreadsheet size={20} color="#2B4C43" />
                    <div style={{ flex: 1, textAlign: "left" }}>
                      <div style={{ fontWeight: 700, fontSize: 14 }}>Excel workbook (.xlsx)</div>
                      <div style={{ fontSize: 12, color: "#8C8577" }}>Summary sheet, an all-cases sheet, and one sheet per specialty.</div>
                    </div>
                    {exportBusy === "excel" ? <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> : <Download size={16} color="#8C8577" />}
                  </button>

                  <button
                    onClick={() => runExport("word", () => exportToWord(statsScopedCases, stats))}
                    disabled={exportBusy !== null}
                    style={{ ...exportCardStyle }}
                  >
                    <FileText size={20} color="#2B4C43" />
                    <div style={{ flex: 1, textAlign: "left" }}>
                      <div style={{ fontWeight: 700, fontSize: 14 }}>Word document (.doc)</div>
                      <div style={{ fontSize: 12, color: "#8C8577" }}>A formatted report: summary tables, then a section per specialty.</div>
                    </div>
                    {exportBusy === "word" ? <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> : <Download size={16} color="#8C8577" />}
                  </button>

                  <button
                    onClick={() => runExport("pdf", () => openPrintWindow(buildReportHtml(statsScopedCases, stats), { title: "Surgical Case Log", landscape: false }))}
                    disabled={exportBusy !== null}
                    style={{ ...exportCardStyle }}
                  >
                    <FileIcon size={20} color="#2B4C43" />
                    <div style={{ flex: 1, textAlign: "left" }}>
                      <div style={{ fontWeight: 700, fontSize: 14 }}>PDF report</div>
                      <div style={{ fontSize: 12, color: "#8C8577" }}>Opens a print preview — choose "Save as PDF" in the print dialog.</div>
                    </div>
                    {exportBusy === "pdf" ? <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> : <Download size={16} color="#8C8577" />}
                  </button>

                  <button
                    onClick={() => runExport("slides", () => openPrintWindow(buildSlidesHtml(statsScopedCases, stats), { title: "Surgical Case Log — Slides", landscape: true }))}
                    disabled={exportBusy !== null}
                    style={{ ...exportCardStyle }}
                  >
                    <Images size={20} color="#2B4C43" />
                    <div style={{ flex: 1, textAlign: "left" }}>
                      <div style={{ fontWeight: 700, fontSize: 14 }}>Slide-style PDF (for PowerPoint)</div>
                      <div style={{ fontSize: 12, color: "#8C8577" }}>
                        A landscape, one-topic-per-page PDF — a real, editable .pptx isn't available in-browser, but this presents cleanly as-is or drops into PowerPoint as image slides.
                      </div>
                    </div>
                    {exportBusy === "slides" ? <Loader2 size={16} style={{ animation: "spin 1s linear infinite" }} /> : <Download size={16} color="#8C8577" />}
                  </button>
                </>
              )}
            </>
          )}
        </div>
      )}

      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes pulseDot { 0%, 100% { opacity: 1; transform: scale(1); } 50% { opacity: 0.4; transform: scale(0.8); } }
      `}</style>
    </div>
  );
}

function SummaryCard({ label, value, small }) {
  return (
    <div style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 12, padding: "12px 14px" }}>
      <div style={{ fontSize: 11, color: "#8C8577", fontWeight: 600, marginBottom: 4 }}>{label}</div>
      <div style={{
        fontSize: small ? 14.5 : 24,
        fontWeight: 700,
        fontFamily: small ? "'IBM Plex Sans', sans-serif" : "'IBM Plex Mono', monospace",
        color: "#1E2624",
        lineHeight: 1.2,
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap"
      }}>
        {value}
      </div>
    </div>
  );
}

function StatBlock({ title, children }) {
  return (
    <div style={{ background: "#fff", border: "1px solid #EAE5D8", borderRadius: 14, padding: 16 }}>
      <div style={{ fontFamily: "'Source Serif 4', serif", fontWeight: 700, fontSize: 15, marginBottom: 12 }}>{title}</div>
      {children}
    </div>
  );
}
