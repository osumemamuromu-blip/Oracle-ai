import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/* Settings                                                            */
/* ------------------------------------------------------------------ */

// Set to false to switch Drawing, CAD, Video and Advert studios back on.
const PAUSED_PRO_FEATURES = true;

const TEXT_MODEL = "@cf/google/gemma-4-26b-a4b-it";
// Used only when the user sends a photo. Check the Cloudflare model page:
// Llama 3.2 Vision needs a one-time license acceptance.
const VISION_MODEL = "@cf/meta/llama-3.2-11b-vision-instruct";
// Engineering Drawing Studio uses a stronger vision model (falls back to VISION_MODEL if unavailable).
const DRAWING_VISION_MODEL = "@cf/meta/llama-4-scout-17b-16e-instruct";
const IMAGE_MODEL = "@cf/black-forest-labs/flux-1-schnell";
// Drawing uploads allowed per user per day (on top of the normal daily request limit).
const DRAWING_DAILY_LIMIT = 2;

// Daily request limits per plan. Change the numbers here only.
const DAILY_LIMITS = {
  free: 8,
  premium: 100,
  pro: 300
};

// Answer length (max output tokens) per plan, so paid plans get fuller answers.
const CHAT_MAX_TOKENS = { free: 2048, premium: 4096, pro: 6144 };

const PRICES = {
  premium: { monthly: 2500, annual: 25000 },
  pro: { monthly: 7500, annual: 82500 }
};

const taskInstructions = {
  correct:
    "Correct grammar, spelling, punctuation, and sentence errors. Preserve the writer's meaning, facts, and voice. Return the corrected text first, then briefly explain the important corrections made.",
  improve:
    "Improve clarity, grammar, flow, and readability while preserving the writer's meaning, facts, and voice. Return the improved text first, then explain the most important improvements.",
  generate:
    "Write original, clear, useful content based on the user's instructions. Follow the requested format and tone. Explain difficult ideas simply when appropriate.",
  explain:
    "Teach the material step by step using simple language. Explain what it means, why it works, and how to apply it. Give a relevant example where helpful. Do not merely repeat the material.",
  summarize:
    "Summarize the supplied material accurately. Preserve its main ideas and important details. Do not invent information. Explain unfamiliar terms briefly when useful.",
  tone:
    "Rewrite the supplied text in the requested tone while preserving its meaning and important details. Briefly explain significant tone changes.",
  rewrite:
    "Rewrite the supplied text to improve clarity and natural flow while preserving its meaning and facts. Briefly explain important changes.",
  expand:
    "Expand the supplied text with relevant explanations, examples, and details. Do not invent facts or citations. Make the added explanations clear and useful."
};

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization"
};

/* ------------------------------------------------------------------ */
/* Basic helpers                                                       */
/* ------------------------------------------------------------------ */

function respond(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...corsHeaders,
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "X-Content-Type-Options": "nosniff",
      "Referrer-Policy": "no-referrer",
      "X-Frame-Options": "DENY",
      "Strict-Transport-Security": "max-age=31536000; includeSubDomains"
    }
  });
}

function extractText(data) {
  if (typeof data === "string") return data.trim();

  const text =
    data?.response ??
    data?.result ??
    data?.choices?.[0]?.message?.content ??
    data?.output?.[0]?.content?.[0]?.text;

  return typeof text === "string" ? text.trim() : "";
}

function stripCodeFences(text) {
  return text
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "");
}

function parseJsonObject(text, failMessage) {
  const cleaned = stripCodeFences(text);

  try {
    return JSON.parse(cleaned);
  } catch {
    const start = cleaned.indexOf("{");
    const end = cleaned.lastIndexOf("}");

    if (start < 0 || end <= start) {
      throw new Error(failMessage);
    }

    try {
      return JSON.parse(cleaned.slice(start, end + 1));
    } catch {
      throw new Error(failMessage);
    }
  }
}

function getErrorStatus(error) {
  if ([400, 401, 403, 429, 503].includes(error.status)) {
    return error.status;
  }
  return 500;
}

/* ------------------------------------------------------------------ */
/* Quiz parsing                                                        */
/* ------------------------------------------------------------------ */

function parseQuiz(text) {
  const parsed = parseJsonObject(
    text,
    "The AI did not return valid quiz JSON. Please try again."
  );

  const questions = Array.isArray(parsed) ? parsed : parsed.questions;

  if (!Array.isArray(questions) || !questions.length) {
    throw new Error("No quiz questions were returned. Please try again.");
  }

  const validated = questions.map((item, index) => {
    const question = String(item.question || item.prompt || "").trim();
    const options = item.options || item.choices;

    if (!question || !Array.isArray(options) || options.length < 2) {
      throw new Error(
        `Quiz question ${index + 1} is incomplete. Please generate the quiz again.`
      );
    }

    const cleanOptions = options.map(option =>
      String(
        typeof option === "object"
          ? option.text ?? option.label ?? ""
          : option
      ).trim()
    );

    let answer =
      item.answerIndex ??
      item.correctIndex ??
      item.correctAnswerIndex ??
      item.correctAnswer ??
      item.answer;

    if (typeof answer === "string") {
      const value = answer.trim();
      const letter = value.toUpperCase().match(/^[A-D]$/);

      if (letter) {
        answer = letter[0].charCodeAt(0) - 65;
      } else {
        answer = cleanOptions.findIndex(
          option => option.toLowerCase() === value.toLowerCase()
        );
      }
    }

    answer = Number(answer);

    if (
      !Number.isInteger(answer) ||
      answer < 0 ||
      answer >= cleanOptions.length
    ) {
      throw new Error(
        `Quiz question ${index + 1} has an invalid answer. Please try again.`
      );
    }

    return {
      question,
      options: cleanOptions,
      answerIndex: answer,
      explanation: String(item.explanation || item.reason || "").trim()
    };
  });

  return { questions: validated };
}

/* ------------------------------------------------------------------ */
/* Excel building                                                      */
/* ------------------------------------------------------------------ */

// Blocks formulas that can reach outside the workbook.
const UNSAFE_FORMULA = /WEBSERVICE|HYPERLINK|\bCALL\b|EXEC|DDE|\||\bREGISTER\b/i;

const EXCEL_LIMITS = {
  premium: { sheets: 3, rows: 500, preview: 10, maxTokens: 4096 },
  pro: { sheets: 10, rows: 2000, preview: 25, maxTokens: 8192 }
};

const NUMBER_FORMATS = {
  currency: "#,##0.00",
  integer: "#,##0",
  number: "#,##0.00",
  percent: "0.0%",
  date: "yyyy-mm-dd"
};

function buildWorkbook(spec, limits = EXCEL_LIMITS.pro) {
  const sheets = Array.isArray(spec?.sheets) ? spec.sheets.slice(0, limits.sheets) : [];

  if (!sheets.length) {
    throw new Error("No spreadsheet data was returned. Please try again.");
  }

  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  sheets.forEach((sheet, i) => {
    let name =
      String(sheet?.name || `Sheet${i + 1}`)
        .replace(/[\\/?*[\]:]/g, "")
        .slice(0, 31) || `Sheet${i + 1}`;

    if (usedNames.has(name.toLowerCase())) {
      name = `Sheet${i + 1}`;
    }
    usedNames.add(name.toLowerCase());

    const rows = (Array.isArray(sheet?.rows) ? sheet.rows : [])
      .slice(0, limits.rows)
      .map(row =>
        (Array.isArray(row) ? row : [row]).slice(0, 50).map(cell => {
          if (cell === null || cell === undefined) return "";
          if (typeof cell === "number") {
            return Number.isFinite(cell) ? cell : "";
          }
          if (typeof cell === "boolean") return cell;

          const value = String(cell);

          if (value.startsWith("=")) {
            if (UNSAFE_FORMULA.test(value) || value.length > 500) {
              // Store as plain text instead of a live formula.
              return { t: "s", v: value };
            }
            return { t: "n", f: value.slice(1) };
          }

          return value;
        })
      );

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // Auto column widths
    const widths = [];
    rows.forEach(row =>
      row.forEach((cell, j) => {
        const len =
          cell && typeof cell === "object" ? 12 : String(cell).length;
        widths[j] = Math.max(widths[j] || 8, Math.min(len + 2, 50));
      })
    );
    ws["!cols"] = widths.map(wch => ({ wch }));

    // Number formats per column (optional "formats" array from the AI), skipping the header row.
    const formats = Array.isArray(sheet?.formats) ? sheet.formats : [];
    formats.forEach((f, c) => {
      const z = NUMBER_FORMATS[String(f || "").toLowerCase()];
      if (!z) return;
      for (let r = 1; r < rows.length; r++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (cell && (cell.t === "n")) cell.z = z;
      }
    });

    // Filter buttons on the header row when there is a table to filter.
    if (rows.length > 2 && rows[0].length > 1 && ws["!ref"]) {
      ws["!autofilter"] = { ref: ws["!ref"] };
    }

    XLSX.utils.book_append_sheet(wb, ws, name);
  });

  return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
}

/* ------------------------------------------------------------------ */
/* PowerPoint building (no extra libraries)                            */
/* ------------------------------------------------------------------ */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// Minimal ZIP writer (no compression), enough for a valid .pptx
function buildZip(files) {
  const enc = new TextEncoder();
  const chunks = [];
  const central = [];
  let offset = 0;

  const u16 = n => [n & 255, (n >>> 8) & 255];
  const u32 = n => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255];

  for (const [name, content] of files) {
    const nameBytes = enc.encode(name);
    const data = typeof content === "string" ? enc.encode(content) : content;
    const crc = crc32(data);

    const local = new Uint8Array([
      ...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0)
    ]);
    chunks.push(local, nameBytes, data);

    central.push(new Uint8Array([
      ...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21),
      ...u32(crc), ...u32(data.length), ...u32(data.length),
      ...u16(nameBytes.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0),
      ...u32(offset)
    ]), nameBytes);

    offset += local.length + nameBytes.length + data.length;
  }

  const centralSize = central.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array([
    ...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length),
    ...u32(centralSize), ...u32(offset), ...u16(0)
  ]);

  const all = [...chunks, ...central, end];
  const out = new Uint8Array(all.reduce((sum, part) => sum + part.length, 0));
  let pos = 0;
  for (const part of all) {
    out.set(part, pos);
    pos += part.length;
  }
  return out;
}

function bytesToBase64(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const PPT_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

// Slide size: 16:9 widescreen (12192000 x 6858000 EMU)
const SLIDE_W = 12192000;
const SLIDE_H = 6858000;
const THEMES = {
  dark:   { BG: "0B0C0D", ACCENT: "D6F36A", TEXT: "F5F5F2", MUTED: "B8BDB0", PANEL: "1A1C1E" },
  light:  { BG: "FAFAF7", ACCENT: "3F7D12", TEXT: "1A1C1E", MUTED: "5A5F55", PANEL: "ECEEE6" },
  ocean:  { BG: "0A1A2F", ACCENT: "5CC8FF", TEXT: "EAF4FF", MUTED: "9DB8D1", PANEL: "12294A" },
  sunset: { BG: "1F0F14", ACCENT: "FF9F5A", TEXT: "FFF1E8", MUTED: "D1AFA0", PANEL: "33191F" },
  forest: { BG: "0E1F17", ACCENT: "8FE388", TEXT: "EFF8F0", MUTED: "A6C2AE", PANEL: "173326" },
  royal:  { BG: "160F2E", ACCENT: "C79BFF", TEXT: "F3EEFF", MUTED: "B7A8D6", PANEL: "261A4D" }
};
let BG, ACCENT, TEXT, MUTED, PANEL;
function applyTheme(name) {
  const t = THEMES[name] || THEMES.dark;
  ({ BG, ACCENT, TEXT, MUTED, PANEL } = t);
}
applyTheme("dark");

function detectTheme(message) {
  const t = String(message || "").toLowerCase();
  const word = "(?:theme|colou?rs?|colou?r scheme|style|background|mode|look)";
  if (new RegExp(`\\b(?:light|white|bright)\\b.{0,20}${word}|${word}.{0,12}\\b(?:light|white)\\b`).test(t)) return "light";
  if (new RegExp(`\\b(?:blue|ocean|navy|sea)\\b.{0,20}${word}|${word}.{0,12}\\b(?:blue|ocean|navy)\\b`).test(t)) return "ocean";
  if (new RegExp(`\\b(?:orange|sunset|warm|red)\\b.{0,20}${word}|${word}.{0,12}\\b(?:orange|sunset|warm|red)\\b`).test(t)) return "sunset";
  if (new RegExp(`\\b(?:green|forest|nature)\\b.{0,20}${word}|${word}.{0,12}\\b(?:green|forest)\\b`).test(t)) return "forest";
  if (new RegExp(`\\b(?:purple|royal|violet)\\b.{0,20}${word}|${word}.{0,12}\\b(?:purple|royal|violet)\\b`).test(t)) return "royal";
  if (new RegExp(`\\bdark\\b.{0,20}${word}`).test(t)) return "dark";
  return null;
}

function textShape(id, name, x, y, w, h, paragraphs, anchor = "t") {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${anchor}"><a:normAutofit/></a:bodyPr><a:lstStyle/>` +
    paragraphs +
    `</p:txBody></p:sp>`
  );
}

function para(text, { size, color, bold = false, bullet = false, spaceAfter = 0 }) {
  const bu = bullet
    ? `<a:buClr><a:srgbClr val="${ACCENT}"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>`
    : "<a:buNone/>";
  const indent = bullet ? ' marL="342900" indent="-342900"' : ' marL="0" indent="0"';
  return (
    `<a:p><a:pPr${indent}><a:spcAft><a:spcPts val="${spaceAfter}"/></a:spcAft>${bu}</a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="${size}" b="${bold ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `<a:latin typeface="Calibri"/></a:rPr><a:t>${xmlEscape(text)}</a:t></a:r></a:p>`
  );
}

function accentBar(id, x, y, w, h) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Accent"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${ACCENT}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  );
}

function slideXml(shapes) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sld ${PPT_NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${BG}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    shapes +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
  );
}


const SLIDE_LAYOUTS = [
  "plain", "right", "left", "splitright", "splitleft", "banner", "background", "caption",
  "cards", "numbered", "timeline", "statement", "columns"
];
const IMAGE_LAYOUTS = new Set(["right", "left", "splitright", "splitleft", "banner", "background", "caption"]);

// Reads picture-position wishes from the user's own words (applies to the whole deck).
// Content layouts such as cards or timeline are left to the AI so a topic word is never mistaken for a layout.
function detectDeckLayout(message) {
  const t = String(message || "").toLowerCase();
  const pic = "(?:images?|pictures?|photos?|pics?|illustrations?)";
  if (new RegExp(`(?:full[- ]?(?:screen|slide|bleed|page)|background)\\s+${pic}|${pic}\\s+(?:as|for)\\s+(?:the\\s+)?background|background\\s+(?:of\\s+)?${pic}`).test(t)) return "background";
  if (new RegExp(`(?:half|split)[- ](?:screen|slide)?\\s*${pic}[^.]{0,20}left|${pic}\\s+(?:fills?|takes?)\\s+(?:the\\s+)?left\\s+half|split\\s+(?:layout\\s+)?(?:with\\s+)?(?:the\\s+)?${pic}\\s+(?:on\\s+)?(?:the\\s+)?left`).test(t)) return "splitleft";
  if (new RegExp(`(?:half|split)[- ](?:screen|slide)?\\s*${pic}|${pic}\\s+(?:fills?|takes?)\\s+(?:the\\s+)?(?:right\\s+)?half|split[- ]?(?:screen|layout)`).test(t)) return "splitright";
  if (new RegExp(`${pic}\\s+(?:on|at|to)\\s+(?:the\\s+)?left|left[- ](?:side|hand)?\\s*${pic}|${pic}\\s+first`).test(t)) return "left";
  if (new RegExp(`${pic}\\s+(?:on|at)\\s+(?:the\\s+)?(?:top|above)|(?:top|banner|wide|header)\\s+${pic}|banner`).test(t)) return "banner";
  if (new RegExp(`caption|text (?:panel|box) (?:over|on) (?:the )?${pic}`).test(t)) return "caption";
  if (/\b(?:two|2)[- ]columns?\b/.test(t)) return "columns";
  if (new RegExp(`${pic}\\s+(?:on|at|to)\\s+(?:the\\s+)?right|right[- ](?:side|hand)?\\s*${pic}`).test(t)) return "right";
  return null;
}

// Uses the user's layout if they gave one, else the AI's choice, else a sensible alternating default.
function pickLayout(slideLayout, deckLayout, index = 0, hasImages = false) {
  if (deckLayout) return deckLayout;
  const l = String(slideLayout || "").toLowerCase().trim();
  if (SLIDE_LAYOUTS.includes(l)) return l;
  return hasImages ? (index % 2 ? "left" : "right") : "plain";
}

function overlayShape(id) {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Overlay"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SLIDE_W}" cy="${SLIDE_H}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${BG}"><a:alpha val="72000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  );
}

function flatPicture(id, rid, x, y, w, h, img) {
  return pictureShape(id, rid, x, y, w, h, img).replace('prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 3500"/></a:avLst>', 'prst="rect"><a:avLst/>');
}


function rectShape(id, name, x, y, w, h, color, { round = 0, alpha = 100000, geom } = {}) {
  const g = geom || (round ? "roundRect" : "rect");
  const av = round ? `<a:avLst><a:gd name="adj" fmla="val ${round}"/></a:avLst>` : "<a:avLst/>";
  const fill = alpha < 100000
    ? `<a:srgbClr val="${color}"><a:alpha val="${alpha}"/></a:srgbClr>`
    : `<a:srgbClr val="${color}"/>`;
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
    `<a:prstGeom prst="${g}">${av}</a:prstGeom><a:solidFill>${fill}</a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
  );
}

// Plain paragraph without a bullet
function plainPara(text, { size, color, bold = false, align = "l", spaceAfter = 0 }) {
  return (
    `<a:p><a:pPr marL="0" indent="0" algn="${align}"><a:spcAft><a:spcPts val="${spaceAfter}"/></a:spcAft><a:buNone/></a:pPr>` +
    `<a:r><a:rPr lang="en-US" sz="${size}" b="${bold ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `<a:latin typeface="Calibri"/></a:rPr><a:t>${xmlEscape(text)}</a:t></a:r></a:p>`
  );
}

function titleSlide(title, subtitle, img, layout = "right") {
  if (img && layout === "background") {
    return slideXml(
      flatPicture(7, "rId2", 0, 0, SLIDE_W, SLIDE_H, img) +
      overlayShape(8) +
      accentBar(2, 914400, 2300000, 1200000, 64000) +
      textShape(3, "Title", 914400, 2500000, SLIDE_W - 1828800, 1500000, para(title, { size: 4800, color: TEXT, bold: true }), "t") +
      (subtitle ? textShape(4, "Subtitle", 914400, 4100000, SLIDE_W - 1828800, 900000, para(subtitle, { size: 2200, color: MUTED }), "t") : "")
    );
  }
  const left = img && layout === "left";
  const textX = left ? 5900000 + 400000 : 914400;
  const textW = img ? 5200000 : SLIDE_W - 1828800;
  const shapes =
    (img ? (left ? flatPicture(7, "rId2", 0, 0, 5900000, SLIDE_H, img) : flatPicture(7, "rId2", 6700000, 0, SLIDE_W - 6700000, SLIDE_H, img)) : "") +
    accentBar(2, textX, 2300000, 1200000, 64000) +
    textShape(3, "Title", textX, 2500000, textW, 1500000,
      para(title, { size: img ? 4000 : 4800, color: TEXT, bold: true }), "t") +
    (subtitle
      ? textShape(4, "Subtitle", textX, 4100000, textW, 900000,
          para(subtitle, { size: img ? 2000 : 2200, color: MUTED }), "t")
      : "");
  return slideXml(shapes);
}

function contentSlide(title, bullets, notesText, number, total, img, layout = "right") {
  const count = bullets.length;
  const baseSize = count > 7 ? 1800 : count > 5 ? 2000 : 2400;
  const mk = (list, size) => list.map(b => para(b, { size, color: TEXT, bullet: true, spaceAfter: 1000 })).join("");

  const M = 685800;
  const FULL_W = SLIDE_W - M * 2;
  const BODY_Y = 1700000;
  const BODY_H = SLIDE_H - 2600000;
  const footer = textShape(5, "Footer", M, SLIDE_H - 600000, FULL_W, 300000,
    para(`${number} / ${total}`, { size: 1200, color: MUTED }), "ctr");
  const titleBox = textShape(2, "Title", M, 457200, FULL_W, 900000,
    para(title, { size: 3200, color: ACCENT, bold: true }), "ctr");
  const bar = accentBar(3, M, 1400000, 900000, 40000);

  if (layout === "columns") {
    const half = Math.ceil(count / 2);
    const size = Math.min(baseSize, 2200);
    const colW = Math.round((FULL_W - 500000) / 2);
    return slideXml(
      titleBox + bar +
      textShape(4, "Left column", M, BODY_Y, colW, BODY_H, mk(bullets.slice(0, half), size), "t") +
      textShape(6, "Right column", M + colW + 500000, BODY_Y, colW, BODY_H, mk(bullets.slice(half), size), "t") +
      footer
    );
  }

  if (layout === "cards" && count >= 2 && count <= 6) {
    const cols = count <= 3 ? count : count === 4 ? 2 : 3;
    const rows = Math.ceil(count / cols);
    const gap = 300000;
    const cw = Math.floor((FULL_W - gap * (cols - 1)) / cols);
    const ch = Math.floor((BODY_H - gap * (rows - 1)) / rows);
    const size = cols === 3 ? 1800 : 2000;
    let out = titleBox + bar;
    bullets.forEach((b, i) => {
      const cx = M + (i % cols) * (cw + gap);
      const cy = BODY_Y + Math.floor(i / cols) * (ch + gap);
      out += rectShape(10 + i * 3, `Card ${i + 1}`, cx, cy, cw, ch, PANEL, { round: 6000 });
      out += rectShape(11 + i * 3, `Card bar ${i + 1}`, cx + 200000, cy + 200000, 520000, 50000, ACCENT);
      out += textShape(12 + i * 3, `Card text ${i + 1}`, cx + 200000, cy + 380000, cw - 400000, ch - 560000,
        plainPara(b, { size, color: TEXT }), "t");
    });
    return slideXml(out + footer);
  }

  if (layout === "numbered" && count >= 2 && count <= 6) {
    const rowH = Math.min(820000, Math.floor(BODY_H / count));
    const dot = Math.min(520000, rowH - 120000);
    let out = titleBox + bar;
    bullets.forEach((b, i) => {
      const y = BODY_Y + i * rowH;
      out += rectShape(10 + i * 3, `Step ${i + 1}`, M, y, dot, dot, ACCENT, { geom: "ellipse" });
      out += textShape(11 + i * 3, `Step number ${i + 1}`, M, y, dot, dot,
        plainPara(String(i + 1), { size: 2000, color: BG, bold: true, align: "ctr" }), "ctr");
      out += textShape(12 + i * 3, `Step text ${i + 1}`, M + dot + 300000, y, FULL_W - dot - 300000, dot,
        plainPara(b, { size: count > 4 ? 1800 : 2000, color: TEXT }), "ctr");
    });
    return slideXml(out + footer);
  }

  if (layout === "timeline" && count >= 3 && count <= 6) {
    const lineY = 2900000;
    const slot = Math.floor(FULL_W / count);
    let out = titleBox + bar + rectShape(9, "Timeline", M, lineY, FULL_W, 38000, MUTED);
    bullets.forEach((b, i) => {
      const x = M + i * slot;
      out += rectShape(10 + i * 3, `Point ${i + 1}`, x, lineY - 110000, 260000, 260000, ACCENT, { geom: "ellipse" });
      out += textShape(11 + i * 3, `Point number ${i + 1}`, x, lineY - 700000, slot - 150000, 400000,
        plainPara(String(i + 1).padStart(2, "0"), { size: 2400, color: ACCENT, bold: true }), "b");
      out += textShape(12 + i * 3, `Point text ${i + 1}`, x, lineY + 420000, slot - 200000, 2300000,
        plainPara(b, { size: count > 4 ? 1500 : 1700, color: TEXT }), "t");
    });
    return slideXml(out + footer);
  }

  if (layout === "statement" && count >= 1) {
    const rest = bullets.slice(1);
    return slideXml(
      titleBox + bar +
      textShape(4, "Statement", M, BODY_Y, FULL_W, rest.length ? 2000000 : BODY_H,
        plainPara(bullets[0], { size: bullets[0].length > 90 ? 3200 : 3800, color: TEXT, bold: true }), "ctr") +
      (rest.length ? rectShape(6, "Divider", M, BODY_Y + 2150000, 600000, 40000, ACCENT) +
        textShape(7, "Support", M, BODY_Y + 2400000, FULL_W, BODY_H - 2400000, mk(rest, 1800), "t") : "") +
      footer
    );
  }

  if (img && (layout === "splitright" || layout === "splitleft")) {
    const imgW = Math.round(SLIDE_W * 0.42);
    const imgOnLeft = layout === "splitleft";
    const textX = imgOnLeft ? imgW + 500000 : M;
    const textW = SLIDE_W - imgW - 500000 - M;
    return slideXml(
      flatPicture(6, "rId2", imgOnLeft ? 0 : SLIDE_W - imgW, 0, imgW, SLIDE_H, img) +
      textShape(2, "Title", textX, 600000, textW, 1100000, para(title, { size: 2800, color: ACCENT, bold: true }), "ctr") +
      accentBar(3, textX, 1800000, 900000, 40000) +
      textShape(4, "Body", textX, 2100000, textW, SLIDE_H - 3000000, mk(bullets, Math.min(baseSize, 2000)), "t") +
      textShape(5, "Footer", textX, SLIDE_H - 600000, textW, 300000, para(`${number} / ${total}`, { size: 1200, color: MUTED }), "ctr")
    );
  }

  if (img && layout === "caption") {
    const panelY = 2700000;
    return slideXml(
      flatPicture(6, "rId2", 0, 0, SLIDE_W, SLIDE_H, img) +
      rectShape(7, "Text panel", M - 250000, panelY, FULL_W + 500000, SLIDE_H - panelY - 350000, BG, { round: 3000, alpha: 88000 }) +
      textShape(2, "Title", M, panelY + 250000, FULL_W, 700000, para(title, { size: 2800, color: ACCENT, bold: true }), "ctr") +
      textShape(4, "Body", M, panelY + 1050000, FULL_W, SLIDE_H - panelY - 1700000, mk(bullets, Math.min(baseSize, 1900)), "t") +
      textShape(5, "Footer", M, SLIDE_H - 600000, FULL_W, 300000, para(`${number} / ${total}`, { size: 1200, color: MUTED }), "ctr")
    );
  }

  if (img && layout === "background") {
    return slideXml(
      flatPicture(6, "rId2", 0, 0, SLIDE_W, SLIDE_H, img) + overlayShape(7) +
      titleBox + bar +
      textShape(4, "Body", M, BODY_Y, FULL_W, BODY_H, mk(bullets, baseSize), "t") +
      footer
    );
  }

  if (img && layout === "banner") {
    const bannerH = 1900000;
    const bodyY = 1600000 + bannerH + 250000;
    return slideXml(
      titleBox + bar +
      pictureShape(6, "rId2", M, 1600000, FULL_W, bannerH, img) +
      textShape(4, "Body", M, bodyY, FULL_W, SLIDE_H - bodyY - 700000, mk(bullets, Math.min(baseSize, 1800)), "t") +
      footer
    );
  }

  if (img) {
    const IMG_W = 4300000;
    const bodyW = FULL_W - IMG_W - 400000;
    const left = layout === "left";
    const imgX = left ? M : SLIDE_W - M - IMG_W;
    const bodyX = left ? M + IMG_W + 400000 : M;
    return slideXml(
      titleBox + bar +
      textShape(4, "Body", bodyX, BODY_Y, bodyW, BODY_H, mk(bullets, Math.max(1400, baseSize - 400)), "t") +
      pictureShape(6, "rId2", imgX, BODY_Y, IMG_W, BODY_H, img) +
      footer
    );
  }

  return slideXml(titleBox + bar + textShape(4, "Body", M, BODY_Y, FULL_W, BODY_H, mk(bullets, baseSize), "t") + footer);
}

function cleanText(value, max) {
  return String(value ?? "").replace(/\s+/g, " ").trim().slice(0, max);
}

/* ---------- Pictures for slides (only when the user asks for them) ---------- */

const PPT_IMAGE_WORDS = /\b(images?|pictures?|photos?|illustrations?|visuals?|graphics?|imagery|pics?)\b/i;
const PPT_NO_IMAGE = /\b(no|without|don'?t (?:add|include|use|want))\s+(?:any\s+)?(images?|pictures?|photos?|illustrations?|visuals?|graphics?|pics?)\b/i;

function wantsSlideImages(message, flag) {
  const text = String(message || "");
  if (PPT_NO_IMAGE.test(text)) return false;
  return flag === true || PPT_IMAGE_WORDS.test(text);
}

function base64ToBytes(b64) {
  const bin = atob(String(b64).replace(/^data:[^,]+,/, "").replace(/\s+/g, ""));
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

// Reads width, height and type from a JPEG or PNG without decoding it.
function imageInfo(bytes) {
  if (bytes.length > 24 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    return { ext: "png", mime: "image/png", w: dv.getUint32(16), h: dv.getUint32(20) };
  }
  if (bytes.length > 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let i = 2;
    while (i + 9 < bytes.length) {
      if (bytes[i] !== 0xff) { i++; continue; }
      const marker = bytes[i + 1];
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { i += 2; continue; }
      const len = (bytes[i + 2] << 8) | bytes[i + 3];
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { ext: "jpg", mime: "image/jpeg", h: (bytes[i + 5] << 8) | bytes[i + 6], w: (bytes[i + 7] << 8) | bytes[i + 8] };
      }
      i += 2 + len;
    }
    return { ext: "jpg", mime: "image/jpeg", w: 1024, h: 1024 };
  }
  return null;
}

let lastImageError = "";

async function generateSlideImage(env, description) {
  if (!env.AI || typeof env.AI.run !== "function") { lastImageError = "AI binding is missing"; return null; }

  const prompt = (
    cleanText(description, 400) +
    ". Clean modern illustration or photo, well lit, no text, no letters, no captions, no watermark, no logos."
  ).slice(0, 900);

  let timeoutId;
  try {
    const result = await Promise.race([
      env.AI.run(IMAGE_MODEL, { prompt, steps: 4 }),
      new Promise((_, reject) => { timeoutId = setTimeout(() => reject(new Error("image timeout")), 45000); })
    ]);

    let bytes = null;
    const raw = typeof result === "string" ? result : result?.image;
    if (typeof raw === "string" && raw.length > 100) bytes = base64ToBytes(raw);
    else if (result instanceof Uint8Array) bytes = result;
    else if (result instanceof ArrayBuffer) bytes = new Uint8Array(result);
    else if (result && typeof result.getReader === "function") bytes = new Uint8Array(await new Response(result).arrayBuffer());

    if (!bytes || !bytes.length) { lastImageError = "The image model returned no picture"; return null; }
    const info = imageInfo(bytes);
    if (!info || !info.w || !info.h) { lastImageError = "The image model returned an unreadable picture"; return null; }
    return { bytes, ...info, alt: cleanText(description, 200) };
  } catch (error) {
    lastImageError = String(error?.message || error).slice(0, 160);
    console.error("Slide image error:", lastImageError);
    return null;
  } finally {
    clearTimeout(timeoutId);
  }
}

// Runs the picture jobs a few at a time so one slow image does not hold up the rest.
async function generateImages(env, jobs, concurrency = 3) {
  const results = new Array(jobs.length).fill(null);
  let next = 0;
  const workers = Array.from({ length: Math.min(concurrency, jobs.length) }, async () => {
    while (next < jobs.length) {
      const idx = next++;
      results[idx] = await generateSlideImage(env, jobs[idx].description);
    }
  });
  await Promise.all(workers);
  return results;
}

// Picture shape that fills its box without stretching (the extra image edge is cropped).
function pictureShape(id, rid, x, y, w, h, img) {
  const imgAspect = img.w / img.h;
  const boxAspect = w / h;
  let l = 0, t = 0, r = 0, b = 0;
  if (imgAspect > boxAspect) {
    const keep = boxAspect / imgAspect;
    l = r = Math.round(((1 - keep) / 2) * 100000);
  } else if (imgAspect < boxAspect) {
    const keep = imgAspect / boxAspect;
    t = b = Math.round(((1 - keep) / 2) * 100000);
  }
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="${xmlEscape(img.alt || "")}"/>` +
    `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rid}"/><a:srcRect l="${l}" t="${t}" r="${r}" b="${b}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 3500"/></a:avLst></a:prstGeom></p:spPr></p:pic>`
  );
}

function normalizeSlides(spec) {
  const rawSlides = Array.isArray(spec?.slides) ? spec.slides.slice(0, 30) : [];
  const slides = rawSlides
    .map(s => ({
      title: cleanText(s?.title, 120),
      bullets: (Array.isArray(s?.bullets) ? s.bullets : [])
        .map(b => cleanText(b, 300))
        .filter(Boolean)
        .slice(0, 10),
      image: cleanText(s?.image, 300),
      layout: cleanText(s?.layout, 20).toLowerCase()
    }))
    .filter(s => s.title || s.bullets.length);

  if (!slides.length) {
    throw new Error("No slides were returned. Please try again.");
  }

  return {
    slides,
    deckTitle: cleanText(spec?.title, 120) || slides[0].title || "Presentation",
    subtitle: cleanText(spec?.subtitle, 200),
    titleImage: cleanText(spec?.titleImage, 300)
  };
}

// images: { title: imgOrNull, slides: [imgOrNull, ...] } matching normalizeSlides() order
function buildPresentation(spec, images = {}, deckLayout = null, themeName = "dark") {
  applyTheme(themeName);
  const hasPics = Boolean(images.title) || (images.slides || []).some(Boolean);
  const { slides, deckTitle, subtitle } = normalizeSlides(spec);
  const total = slides.length + 1;

  const mediaFiles = [];
  const slideImageFor = [];            // index in slideFiles -> media file name or null
  const addMedia = img => {
    if (!img) return null;
    const name = `image${mediaFiles.length + 1}.${img.ext}`;
    mediaFiles.push([`ppt/media/${name}`, img.bytes, img.ext]);
    return name;
  };

  const titleImg = images.title || null;
  const titleLayout = deckLayout === "left" || deckLayout === "splitleft" ? "left" : deckLayout === "background" || deckLayout === "caption" ? "background" : "right";
  const slideFiles = [titleSlide(deckTitle, subtitle, titleImg, titleLayout)];
  slideImageFor.push(addMedia(titleImg));
  slides.forEach((s, i) => {
    const layout = pickLayout(s.layout, deckLayout, i, hasPics);
    const img = IMAGE_LAYOUTS.has(layout) ? ((images.slides && images.slides[i]) || null) : null;
    slideFiles.push(contentSlide(s.title || `Slide ${i + 2}`, s.bullets, "", i + 2, total, img, layout));
    slideImageFor.push(addMedia(img));
  });

  const n = slideFiles.length;
  const files = [];

  files.push(["[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    (mediaFiles.some(m => m[2] === "jpg") ? `<Default Extension="jpg" ContentType="image/jpeg"/>` : "") +
    (mediaFiles.some(m => m[2] === "png") ? `<Default Extension="png" ContentType="image/png"/>` : "") +
    `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    slideFiles.map((_, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    ).join("") +
    `</Types>`]);

  files.push(["_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
    `</Relationships>`]);

  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  files.push(["docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${xmlEscape(deckTitle)}</dc:title><dc:creator>Malie AI</dc:creator>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`]);

  files.push(["docProps/app.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Malie AI</Application><Slides>${n}</Slides></Properties>`]);

  files.push(["ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:presentation ${PPT_NS}>` +
    `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    `<p:sldIdLst>${slideFiles.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("")}</p:sldIdLst>` +
    `<p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}"/><p:notesSz cx="6858000" cy="9144000"/>` +
    `</p:presentation>`]);

  files.push(["ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>` +
    slideFiles.map((_, i) =>
      `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
    ).join("") +
    `<Relationship Id="rId${n + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
    `</Relationships>`]);

  files.push(["ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldMaster ${PPT_NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="${BG}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` +
    `<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>` +
    `<p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
    `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>` +
    `<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>` +
    `</p:sldMaster>`]);

  files.push(["ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>` +
    `</Relationships>`]);

  files.push(["ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<p:sldLayout ${PPT_NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>` +
    `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
    `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`]);

  files.push(["ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>` +
    `</Relationships>`]);

  files.push(["ppt/theme/theme1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Malie"><a:themeElements>` +
    `<a:clrScheme name="Malie"><a:dk1><a:srgbClr val="${BG}"/></a:dk1><a:lt1><a:srgbClr val="${TEXT}"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="1A1C1E"/></a:dk2><a:lt2><a:srgbClr val="${MUTED}"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="${ACCENT}"/></a:accent1><a:accent2><a:srgbClr val="7FBF3F"/></a:accent2>` +
    `<a:accent3><a:srgbClr val="3F7D12"/></a:accent3><a:accent4><a:srgbClr val="4F8A1A"/></a:accent4>` +
    `<a:accent5><a:srgbClr val="9BB84A"/></a:accent5><a:accent6><a:srgbClr val="5C6B3A"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="${ACCENT}"/></a:hlink><a:folHlink><a:srgbClr val="${MUTED}"/></a:folHlink></a:clrScheme>` +
    `<a:fontScheme name="Malie"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
    `<a:fmtScheme name="Malie"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
    `</a:themeElements></a:theme>`]);

  slideFiles.forEach((xml, i) => {
    files.push([`ppt/slides/slide${i + 1}.xml`, xml]);
    files.push([`ppt/slides/_rels/slide${i + 1}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
      (slideImageFor[i]
        ? `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${slideImageFor[i]}"/>`
        : "") +
      `</Relationships>`]);
  });

  mediaFiles.forEach(([name, bytes]) => files.push([name, bytes]));

  return bytesToBase64(buildZip(files));
}

/* ------------------------------------------------------------------ */
/* Reply language                                                      */
/* ------------------------------------------------------------------ */

const REPLY_LANGUAGES = {
  en: "English", fr: "French", es: "Spanish", pt: "Portuguese", de: "German",
  ar: "Arabic", hi: "Hindi", zh: "Simplified Chinese", ru: "Russian",
  sw: "Swahili", yo: "Yoruba", ig: "Igbo", ha: "Hausa", pcm: "Nigerian Pidgin"
};

// Returns one prompt line, or nothing when the user's language is English or unknown.
function languageLine(code, what = "Write your reply") {
  const name = REPLY_LANGUAGES[String(code || "").toLowerCase()];
  if (!name || name === "English") return [];
  return [`${what} in ${name}, unless the user clearly writes in a different language or asks for another one.`];
}

/* ------------------------------------------------------------------ */
/* AI call                                                             */
/* ------------------------------------------------------------------ */

async function callAI(env, prompt, jsonMode = false, imageDataUrl = null, options = {}) {
  if (!env.AI || typeof env.AI.run !== "function") {
    const error = new Error(
      "Cloudflare Workers AI is not connected. Check that your Worker has an AI binding named AI."
    );
    error.status = 500;
    throw error;
  }

  const hasImage = Boolean(imageDataUrl);
  const model = options.model || (hasImage ? VISION_MODEL : TEXT_MODEL);

  const identityRules = [
    "You are Malie AI, the AI assistant created and developed by Campus Creative, founded by Osume Prosper.",
    "IDENTITY RULE: When asked who created you, who made you, who developed you, who owns Malie AI, or who your founder is, answer clearly that Malie AI was created and developed by Campus Creative, founded by Osume Prosper.",
    "Never answer that Google, Gemini, Google AI, Cloudflare, or another technology provider created Malie AI. They may provide underlying technology, but Campus Creative is the company behind the Malie AI product.",
    "If asked who provides your underlying AI technology, explain that Malie AI currently runs on Cloudflare Workers AI and uses the configured Gemma model.",
    "Be truthful, consistent, and confident about your identity. Do not confuse the creator of Malie AI with the provider of its underlying AI model."
  ];

  const tutorRules = [
    "You are Malie AI, a helpful, patient, accurate AI tutor and writing assistant.",
    "Teach rather than merely give answers.",
    "When answering educational questions, explain the reasoning step by step in simple language.",
    "Define unfamiliar terms and use a relevant example when it helps.",
    "For calculations, show the method, the working, and the final answer with units where appropriate.",
    "For science questions, explain why the answer is correct, not just what the answer is.",
    "For writing tasks, follow the requested task and preserve the user's meaning and facts.",
    "Use headings or numbered steps when they improve clarity.",
    "If the request is unclear, ask a brief clarifying question rather than inventing details.",
    "Be honest about uncertainty and do not invent facts, sources, or citations."
  ];

  const systemPrompt = options.system
    ? [...identityRules, options.system].join(" ")
    : (jsonMode
        ? identityRules
        : [...identityRules, ...tutorRules]
      ).join(" ");

  const userContent = hasImage
    ? [
        { type: "text", text: prompt },
        { type: "image_url", image_url: { url: imageDataUrl } }
      ]
    : prompt;

  const payload = {
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userContent }
    ],
    max_tokens: options.maxTokens || (jsonMode ? 4096 : 2048)
  };
  if (typeof options.temperature === "number") payload.temperature = options.temperature;

  // These options are only sent to the Gemma text model.
  if (!hasImage) {
    if (jsonMode) payload.response_format = { type: "json_object" };
    payload.chat_template_kwargs = { enable_thinking: false };
  }

  let result;

  try {
    result = await env.AI.run(model, payload);
  } catch (cause) {
    console.error("Workers AI error:", String(cause));

    const message = String(cause?.message || cause);
    const isLimit = /limit|quota|neuron|429/i.test(message);
    const isBusy = /capacity|busy|overload|503/i.test(message);
    const needsLicense = /agree|license/i.test(message) && hasImage;

    const error = new Error(
      needsLicense
        ? "The image model has not been enabled yet. The site owner must accept its license once."
        : isLimit
        ? "The Cloudflare AI usage limit may have been reached. Please try again after the usage limit resets."
        : isBusy
        ? "Cloudflare AI is busy right now. Please wait a little and try again."
        : "Cloudflare AI could not complete the request. Please try again."
    );

    error.status = isLimit ? 429 : 503;
    throw error;
  }

  const text = extractText(result);

  if (!text) {
    console.error(
      "Unexpected Workers AI response:",
      JSON.stringify(result).slice(0, 1000)
    );
    throw new Error("The AI returned no text. Please try again.");
  }

  return text;
}


/* ------------------------------------------------------------------ */
/* Premium chat models (Malie KrakenBot 3 Efficient / Malie LEO 5)      */
/* Paid plans only. Provider names are never shown to users.            */
/* ------------------------------------------------------------------ */

// Daily message limits PER MODEL for paying users. Change the numbers here only.
// (These messages also count toward the plan's normal daily request limit.)
// Limits are in CREDITS, not messages: the harder the task, the more credits it uses.
//   Quick question = 1 credit, normal = 2, long/heavy = 3-5 (see estimateCreditCost).
const PREMIUM_MODELS = {
  kraken: { label: "Malie KrakenBot 3 Efficient", limits: { premium: 60, pro: 200 } },
  leo:    { label: "Malie LEO 5",                 limits: { premium: 40, pro: 130 } }
};

// Works out how many credits a request uses, from its size and how demanding it is.
function estimateCreditCost(rawMessage) {
  let text = String(rawMessage || "");
  const marker = "New message from the user:";
  const i = text.lastIndexOf(marker);
  if (i !== -1) text = text.slice(i + marker.length); // ignore chat history, judge the new message
  text = text.trim();

  let cost = text.length < 300 ? 1 : text.length < 1500 ? 2 : text.length < 6000 ? 3 : 4;

  const heavy = /\b(essay|report|assignment|project|research|analy[sz]e|analysis|derive|derivation|prove|proof|solve|code|program|script|algorithm|step[- ]by[- ]step|in detail|detailed|thorough|comprehensive|full explanation|write (?:a|an|me)|summari[sz]e|compare|lesson plan|case study)\b/i;
  if (heavy.test(text)) cost += 1;

  return Math.min(5, cost);
}

const PREMIUM_MODEL_SYSTEM = [
  "You are Malie AI, a helpful, patient, accurate AI tutor and writing assistant created and developed by Campus Creative, founded by Osume Prosper.",
  "Never reveal, guess, or discuss which underlying AI model, company, or provider powers you. If asked, say you are a Malie AI model built by Campus Creative.",
  "Teach rather than merely give answers: explain step by step in simple language, define unfamiliar terms, and show working for calculations.",
  "Be honest about uncertainty and never invent facts, sources, or citations. If the request is unclear, ask one brief clarifying question."
].join(" ");

async function reservePremiumModelUsage(env, userId, modelKey, plan, cost = 1) {
  const limit = PREMIUM_MODELS[modelKey].limits[plan];
  const date = getUsageDate();
  await env.DB.prepare(
    `CREATE TABLE IF NOT EXISTS premium_model_usage (
       user_id TEXT NOT NULL, usage_date TEXT NOT NULL, model TEXT NOT NULL,
       used INTEGER NOT NULL DEFAULT 0, PRIMARY KEY (user_id, usage_date, model))`
  ).run();
  const row = await env.DB.prepare(
    `INSERT INTO premium_model_usage (user_id, usage_date, model, used) VALUES (?, ?, ?, ?)
     ON CONFLICT(user_id, usage_date, model) DO UPDATE SET used = used + ? WHERE used + ? <= ?
     RETURNING used`
  ).bind(userId, date, modelKey, cost <= limit ? cost : limit + 1, cost, cost, limit).first();
  if (!row) {
    const error = new Error(
      `You don't have enough daily credits left on ${PREMIUM_MODELS[modelKey].label} for this request (daily allowance: ${limit} credits; bigger tasks use more). ` +
      (plan === "premium"
        ? "Upgrade to Pro for a higher limit, switch to Malie 1.1, or try again tomorrow."
        : "Switch to Malie 1.1 or try again tomorrow.")
    );
    error.status = 429;
    throw error;
  }
  return { used: row.used, limit, remaining: limit - row.used, cost, date };
}

async function refundPremiumModelUsage(env, userId, modelKey, date, cost = 1) {
  try {
    await env.DB.prepare(
      `UPDATE premium_model_usage SET used = CASE WHEN used >= ? THEN used - ? ELSE 0 END
       WHERE user_id = ? AND usage_date = ? AND model = ?`
    ).bind(cost, cost, userId, date, modelKey).run();
  } catch (e) { console.error("Premium model refund failed:", String(e)); }
}

async function callPremiumModel(env, modelKey, prompt, maxTokens) {
  const isKraken = modelKey === "kraken";
  const key = isKraken ? env.NVIDIA_API_KEY : env.DEEPINFRA_API_KEY;
  const url = isKraken
    ? "https://integrate.api.nvidia.com/v1/chat/completions"
    : "https://api.deepinfra.com/v1/openai/chat/completions";
  const model = isKraken
    ? (env.NVIDIA_MODEL || "nvidia/nemotron-3.5-lightning-30b-a3b")
    : env.DEEPINFRA_MODEL;

  const unavailable = (status = 503) => {
    const e = new Error(`${PREMIUM_MODELS[modelKey].label} is unavailable right now. Please try again shortly or switch to Malie 1.1.`);
    e.status = status;
    return e;
  };
  if (!key || !model) { console.error("Premium model not configured (missing API key or model name):", modelKey); throw unavailable(500); }

  const body = {
    model,
    messages: [
      { role: "system", content: PREMIUM_MODEL_SYSTEM },
      { role: "user", content: prompt }
    ],
    max_tokens: maxTokens,
    stream: false
  };
  if (isKraken) body.chat_template_kwargs = { enable_thinking: false };

  let res;
  try {
    res = await fetch(url, {
      method: "POST",
      headers: { "Authorization": `Bearer ${key}`, "Content-Type": "application/json" },
      signal: AbortSignal.timeout(60000),
      body: JSON.stringify(body)
    });
  } catch (err) {
    console.error("Premium model network error:", modelKey, String(err));
    throw unavailable();
  }
  if (!res.ok) {
    console.error("Premium model HTTP error:", modelKey, res.status, (await res.text()).slice(0, 500));
    throw unavailable(res.status === 429 ? 429 : 503);
  }
  const data = await res.json().catch(() => null);
  let text = String(data?.choices?.[0]?.message?.content || "");
  text = text.replace(/<think>[\s\S]*?<\/think>/gi, "").trim();
  if (!text) throw unavailable();
  return text;
}

async function callNvidiaAI(env, prompt, jsonMode = true, opts = {}) {
  if (!env.NVIDIA_API_KEY) {
    const error = new Error("NVIDIA API key is not configured.");
    error.status = 500;
    throw error;
  }

  const model =
    env.NVIDIA_MODEL || "nvidia/nemotron-3.5-lightning-30b-a3b";

  // At most two attempts, and only for temporary upstream failures.
  // Empty reasoning-only responses and invalid output are not retried here.
  const maxAttempts = 2;
  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await nvidiaStreamOnce(env, model, prompt, jsonMode, opts);
    } catch (error) {
      lastError = error;

      if (!error.retryable || attempt >= maxAttempts) {
        break;
      }

      console.warn("Retrying temporary NVIDIA failure:", JSON.stringify({
        attempt,
        message: String(error.message || error),
        status: error.status || null
      }));

      await new Promise(resolve => setTimeout(resolve, 1200));
    }
  }

  throw lastError || new Error("NVIDIA CAD generation failed.");
}

async function nvidiaStreamOnce(env, model, prompt, jsonMode, opts = {}) {
  const fail = (message, status = 502, retryable = false) => {
    const error = new Error(message);
    error.status = status;
    error.retryable = retryable;
    return error;
  };

  let response;

  try {
    response = await fetch(
      "https://integrate.api.nvidia.com/v1/chat/completions",
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${env.NVIDIA_API_KEY}`,
          "Content-Type": "application/json",
          "Accept": "text/event-stream"
        },
        signal: AbortSignal.timeout(opts.timeoutMs || 100000),
        body: JSON.stringify({
          
model,
chat_template_kwargs: {
  enable_thinking: false
},
messages: [

            {
              role: "system",
              content: opts.system ? opts.system : jsonMode
                ? "You are Malie AI's CAD code generator. Return only a valid JSON object matching the requested schema. Do not use Markdown fences or explanations. Prioritize complete, executable code."
                : "You are Malie AI's CAD code generator. Return the requested result directly."
            },
            {
              role: "user",
              content: prompt
            }
          ],
          temperature: 0.2,
          max_tokens: opts.maxTokens || 6000,
...(jsonMode ? { response_format: { type: "json_object" } } : {}),
stream: true
        })
      }
    );
  } catch (cause) {
    const timedOut =
      cause?.name === "TimeoutError" ||
      cause?.name === "AbortError";

    console.error("NVIDIA fetch failed:", JSON.stringify({
      name: cause?.name || null,
      message: String(cause?.message || cause)
    }));

    throw fail(
      timedOut
        ? "NVIDIA AI request timed out. Try a simpler CAD description."
        : "Could not connect to NVIDIA AI. Please try again.",
      502,
      true
    );
  }

  if (!response.ok) {
    const details = await response.text().catch(() => "");

    console.error("NVIDIA API error:", JSON.stringify({
      status: response.status,
      body: details.slice(0, 1000)
    }));

    if (response.status === 429) {
      throw fail(
        "NVIDIA rejected the request because of a rate limit or account limit. Check your NVIDIA API usage and account status.",
        429,
        false
      );
    }

    const retryable =
      response.status === 408 ||
      response.status === 425 ||
      response.status >= 500;

    throw fail(
      `NVIDIA API returned HTTP ${response.status}. Check the Worker logs for details.`,
      502,
      retryable
    );
  }

  if (!response.body) {
    throw fail(
      "NVIDIA returned no response stream.",
      502,
      true
    );
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();

  let buffer = "";
  let result = "";
  let finishReason = null;
  let sawReasoning = false;
  let streamError = null;
  let receivedEvents = 0;

  function handleLine(line) {
    const trimmed = line.trim();

    if (!trimmed.startsWith("data:")) {
      return;
    }

    const payload = trimmed.slice(5).trim();

    if (!payload || payload === "[DONE]") {
      return;
    }

    let json;

    try {
      json = JSON.parse(payload);
    } catch {
      // Ignore incomplete or non-JSON SSE lines.
      return;
    }

    receivedEvents++;

    if (json?.error) {
      const detail =
        json.error?.message ||
        json.error?.detail ||
        "NVIDIA reported a streaming error.";

      streamError = String(detail).slice(0, 250);
      return;
    }

    const choice = json?.choices?.[0];
    const delta = choice?.delta;

    if (typeof delta?.content === "string") {
      result += delta.content;
    }

    if (delta?.reasoning_content || delta?.reasoning) {
      sawReasoning = true;
    }

    if (choice?.finish_reason) {
      finishReason = choice.finish_reason;
    }
  }

  try {
    while (true) {
      const { done, value } = await reader.read();

      if (done) {
        break;
      }

      buffer += decoder.decode(value, { stream: true });

      let newlineIndex;

      while ((newlineIndex = buffer.indexOf("\n")) !== -1) {
        const line = buffer.slice(0, newlineIndex);
        buffer = buffer.slice(newlineIndex + 1);
        handleLine(line);
      }
    }

    buffer += decoder.decode();

    if (buffer.trim()) {
      handleLine(buffer);
    }
  } catch (cause) {
    console.error("NVIDIA stream interrupted:", JSON.stringify({
      name: cause?.name || null,
      message: String(cause?.message || cause),
      receivedChars: result.length,
      receivedEvents
    }));

    const timedOut =
      cause?.name === "TimeoutError" ||
      cause?.name === "AbortError";

    throw fail(
      timedOut
        ? "NVIDIA AI request timed out before generation completed."
        : "The connection to NVIDIA AI was interrupted.",
      502,
      true
    );
  } finally {
    try {
      reader.releaseLock();
    } catch {
      // The stream may already have released its reader.
    }
  }

  console.log("NVIDIA CAD response diagnostic:", JSON.stringify({
    model,
    finishReason,
    contentLength: result.length,
    sawReasoning,
    receivedEvents,
    hasStreamError: Boolean(streamError)
  }));

  if (streamError) {
    // Do not automatically retry arbitrary model or account errors.
    throw fail(
      `NVIDIA streaming error: ${streamError}`,
      502,
      false
    );
  }

  if (!result.trim()) {
    if (sawReasoning) {
      throw fail(
        "NVIDIA produced reasoning but no final code. Try a simpler request or check the selected model's generation settings.",
        502,
        false
      );
    }

    throw fail(
      receivedEvents === 0
        ? "NVIDIA returned an empty stream. Please try again."
        : "NVIDIA returned no final code. Please simplify the request.",
      502,
      false
    );
  }

  if (finishReason === "length") {
    throw fail(
      "NVIDIA stopped because the output reached its token limit. Simplify the model or reduce the requested details.",
      502,
      false
    );
  }

  return result;
}


/* ------------------------------------------------------------------ */
/* Auth, plans, usage                                                  */
/* ------------------------------------------------------------------ */

async function authenticateFirebase(request, env) {
  if (!env.FIREBASE_WEB_API_KEY) {
    const error = new Error("Firebase verification is not configured.");
    error.status = 500;
    throw error;
  }

  const authorization = request.headers.get("Authorization") || "";
  const match = authorization.match(/^Bearer\s+(.+)$/i);

  if (!match) {
    const error = new Error("Please sign in with Google to use Malie AI.");
    error.status = 401;
    throw error;
  }

  let response;

  try {
    response = await fetch(
      "https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=" +
        encodeURIComponent(env.FIREBASE_WEB_API_KEY),
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idToken: match[1] })
      }
    );
  } catch {
    const error = new Error("Could not verify your sign-in. Please try again.");
    error.status = 503;
    throw error;
  }

  if (!response.ok) {
    const error = new Error(
      "Your sign-in session is invalid or expired. Please sign in again."
    );
    error.status = 401;
    throw error;
  }

  const data = await response.json();
  const user = data.users?.[0];

  if (!user?.localId) {
    const error = new Error("Could not verify your sign-in.");
    error.status = 401;
    throw error;
  }

  return user.localId;
}

function getUsageDate() {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Africa/Lagos",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());

  const values = {};

  for (const part of parts) {
    if (part.type !== "literal") {
      values[part.type] = part.value;
    }
  }

  return `${values.year}-${values.month}-${values.day}`;
}

// Returns "free", "premium" or "pro".
// A paid plan counts only if it is active and has not expired.
async function getActivePlan(env, userId) {
  if (!env.DB) {
    const error = new Error(
      "Usage tracking is not connected. Please try again later."
    );
    error.status = 500;
    throw error;
  }

  const subscription = await env.DB.prepare(
    `SELECT plan, status, current_period_end
     FROM subscriptions
     WHERE user_id = ?
     LIMIT 1`
  ).bind(userId).first();

  if (!subscription) return "free";

  const plan = String(subscription.plan || "").toLowerCase();
  const status = String(subscription.status || "").toLowerCase();
  const expiry = subscription.current_period_end
    ? Date.parse(subscription.current_period_end)
    : null;

  const active =
    status === "active" &&
    expiry !== null &&
    Number.isFinite(expiry) &&
    expiry > Date.now();

  if (active && (plan === "premium" || plan === "pro")) {
    return plan;
  }

  return "free";
}

async function reserveDailyUsage(env, userId, knownPlan = null) {
  const plan = knownPlan || (await getActivePlan(env, userId));
  const limit = DAILY_LIMITS[plan] ?? DAILY_LIMITS.free;
  const usageDate = getUsageDate();

  const result = await env.DB.prepare(
    `INSERT INTO daily_usage (user_id, usage_date, used)
     VALUES (?, ?, 1)
     ON CONFLICT(user_id, usage_date)
     DO UPDATE SET used = used + 1
     WHERE used < ?
     RETURNING used`
  ).bind(userId, usageDate, limit).first();

  if (!result) {
    const error = new Error(
      `You've reached your ${limit} daily requests on the ${plan} plan. ${
        plan === "free"
          ? "Upgrade to Premium or Pro for more daily requests."
          : plan === "premium"
          ? "Upgrade to Pro for more daily requests, or try again after your limit resets tomorrow."
          : "Your daily limit will reset tomorrow."
      }`
    );
    error.status = 429;
    throw error;
  }

  return {
    plan,
    used: result.used,
    limit,
    remaining: limit - result.used,
    date: usageDate
  };
}

// Gives back one request when the AI fails, so users aren't charged for errors.
async function refundDailyUsage(env, userId, usageDate) {
  try {
    await env.DB.prepare(
      `UPDATE daily_usage
       SET used = CASE WHEN used > 0 THEN used - 1 ELSE 0 END
       WHERE user_id = ? AND usage_date = ?`
    ).bind(userId, usageDate).run();
  } catch (error) {
    console.error("Usage refund failed:", String(error));
  }
}

/* ------------------------------------------------------------------ */
/* Malie Video Studio (Pro) â€” settings and helpers                     */
/* ------------------------------------------------------------------ */

// >>> DAILY VIDEO ALLOWANCE <<<
// Two limits apply to each Pro user per day. A video is refused if EITHER is used up:
//   1. at most VIDEO_DEFAULT_DAILY_VIDEOS videos, and
//   2. at most VIDEO_DEFAULT_DAILY_SECONDS seconds of video in total.
// Each video is 5 or 10 seconds. Worst-case spend per user per day =
// seconds limit x $0.05 (480p) = 15 x 0.05 = $0.75.
// Change the numbers below, or set Worker variables VIDEO_DAILY_VIDEOS and
// VIDEO_DAILY_SECONDS (Settings > Variables) to override them without editing code.
const VIDEO_DEFAULT_DAILY_VIDEOS = 3;
const VIDEO_DEFAULT_DAILY_SECONDS = 15;

const VIDEO_STUDIO = {
  model: "xai/grok-imagine-video",
  resolution: "480p",
  durations: [5, 10],
  aspectRatios: ["16:9", "9:16", "1:1"],
  promptMin: 10,
  promptMax: 800,
  timeoutMs: 170000
};

function videoError(message, status, refund = true) {
  const error = new Error(message);
  error.status = status;
  error.refund = refund;
  return error;
}

function getVideoDailyLimit(env) {
  const raw = Number.parseInt(env.VIDEO_DAILY_VIDEOS, 10);
  return Number.isInteger(raw) && raw >= 0 && raw <= 50
    ? raw
    : VIDEO_DEFAULT_DAILY_VIDEOS;
}

function getVideoDailySeconds(env) {
  const raw = Number.parseInt(env.VIDEO_DAILY_SECONDS, 10);
  return Number.isInteger(raw) && raw >= 0 && raw <= 600
    ? raw
    : VIDEO_DEFAULT_DAILY_SECONDS;
}

function videoUsageSummary(limit, videos, seconds, secondsLimit, date) {
  return {
    limitVideos: limit,
    usedVideos: videos,
    remainingVideos: Math.max(0, limit - videos),
    limitSeconds: secondsLimit,
    usedSeconds: seconds,
    remainingSeconds: Math.max(0, secondsLimit - seconds),
    date
  };
}

async function getVideoUsage(env, userId, usageDate) {
  const row = await env.DB.prepare(
    `SELECT videos_used, seconds_used FROM video_usage
     WHERE user_id = ? AND usage_date = ?`
  ).bind(userId, usageDate).first();

  return {
    videos: Number(row?.videos_used || 0),
    seconds: Number(row?.seconds_used || 0)
  };
}

// Atomically counts one video (and its seconds) for today, but only if the
// user is still under the daily limit. Two simultaneous requests cannot both
// take the last slot.
async function reserveVideo(env, userId, usageDate, seconds, limit, secondsLimit) {
  const limitError = () =>
    videoError(
      limit <= 0 || secondsLimit <= 0
        ? "Video generation is switched off right now."
        : `You don't have enough of today's video allowance left for this video (daily limit: ${limit} video${limit === 1 ? "" : "s"} and ${secondsLimit} seconds). Try a shorter video, or come back after midnight (Lagos time).`,
      429,
      false
    );

  if (limit <= 0 || secondsLimit <= 0 || seconds > secondsLimit) throw limitError();

  const row = await env.DB.prepare(
    `INSERT INTO video_usage (user_id, usage_date, seconds_used, videos_used)
     VALUES (?, ?, ?, 1)
     ON CONFLICT(user_id, usage_date)
     DO UPDATE SET
       seconds_used = seconds_used + excluded.seconds_used,
       videos_used = videos_used + 1
     WHERE videos_used < ?
       AND seconds_used + excluded.seconds_used <= ?
     RETURNING videos_used, seconds_used`
  ).bind(userId, usageDate, seconds, limit, secondsLimit).first();

  if (!row) throw limitError();

  return { videos: Number(row.videos_used), seconds: Number(row.seconds_used) };
}

async function refundVideo(env, userId, usageDate, seconds) {
  try {
    await env.DB.prepare(
      `UPDATE video_usage
       SET seconds_used = MAX(seconds_used - ?, 0),
           videos_used = MAX(videos_used - 1, 0)
       WHERE user_id = ? AND usage_date = ?`
    ).bind(seconds, userId, usageDate).run();
  } catch (error) {
    console.error("Video usage refund failed:", String(error));
  }
}

// Cloudflare documents the result as { state: "Completed", result: { video: "<url>" } }
// and the raw output schema as { video: "<url>" }. Both shapes are accepted.
function extractVideoUrl(result) {
  const state = typeof result?.state === "string" ? result.state.toLowerCase() : "";

  if (state && !["completed", "complete", "succeeded", "success"].includes(state)) {
    if (/fail|error|cancel|reject/.test(state)) {
      throw videoError("The video could not be generated. Please try again.", 502, true);
    }
    // Not finished and no way to check again: do not claim success, do not refund.
    throw videoError(
      "The video provider has not confirmed completion, so no video is available. Please wait a minute before trying again.",
      504,
      false
    );
  }

  const candidate =
    typeof result === "string"
      ? result
      : result?.result?.video ?? result?.video ?? result?.result?.video?.url ?? result?.video?.url;

  if (typeof candidate !== "string" || candidate.length > 2048) {
    console.error("Unexpected video response:", JSON.stringify(result)?.slice(0, 500));
    throw videoError("The provider finished but returned no usable video. Please try again.", 502, true);
  }

  let parsed;
  try {
    parsed = new URL(candidate);
  } catch {
    throw videoError("The provider returned an invalid video link. Please try again.", 502, true);
  }

  if (parsed.protocol !== "https:") {
    throw videoError("The provider returned an unsafe video link. Please try again.", 502, true);
  }

  return parsed.toString();
}

async function runVideoModel(env, { prompt, duration, aspectRatio }) {
  let timeoutId;

  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(
      () =>
        reject(
          videoError(
            "The video is taking longer than expected. It may still finish on the provider's side, so your seconds were not returned. Please wait a few minutes before trying again.",
            504,
            false
          )
        ),
      VIDEO_STUDIO.timeoutMs
    );
  });

  let result;

  try {
    result = await Promise.race([
      env.AI.run(VIDEO_STUDIO.model, {
        prompt,
        aspect_ratio: aspectRatio,
        duration,
        resolution: VIDEO_STUDIO.resolution
      }),
      timeout
    ]);
  } catch (cause) {
    if (cause && cause.refund === false) throw cause;

    const message = String(cause?.message || cause);
    console.error("Video model error:", message);

    if (/moderat|safety|policy|prohibit|not allowed|violat/i.test(message)) {
      throw videoError(
        "That prompt could not be used because it may break content guidelines. Please rephrase it and try again.",
        400,
        true
      );
    }
    if (/invalid|bad request|unprocessable|validation/i.test(message)) {
      throw videoError("The video settings or prompt were not accepted. Please adjust them and try again.", 400, true);
    }
    if (/quota|credit|balance|billing|insufficient|limit|429/i.test(message)) {
      throw videoError("Video generation is busy or temporarily unavailable. Please try again later.", 503, true);
    }
    throw videoError("The video provider could not complete the request. Please try again.", 502, true);
  } finally {
    clearTimeout(timeoutId);
  }

  return extractVideoUrl(result);
}

/* ---------- Drawing upload limit (per user, per day) ---------- */

async function reserveDrawingUpload(env, userId) {
  const date = getUsageDate();
  await env.DB.prepare(
    `CREATE TABLE IF NOT EXISTS drawing_usage (
       user_id TEXT NOT NULL, usage_date TEXT NOT NULL, used INTEGER NOT NULL DEFAULT 0,
       PRIMARY KEY (user_id, usage_date))`
  ).run();

  const row = await env.DB.prepare(
    `INSERT INTO drawing_usage (user_id, usage_date, used) VALUES (?, ?, 1)
     ON CONFLICT(user_id, usage_date) DO UPDATE SET used = used + 1 WHERE used < ?
     RETURNING used`
  ).bind(userId, date, DRAWING_DAILY_LIMIT).first();

  if (!row) {
    const error = new Error(
      `You've used your ${DRAWING_DAILY_LIMIT} drawing uploads for today. The limit resets tomorrow.`
    );
    error.status = 429;
    throw error;
  }
  return { used: row.used, limit: DRAWING_DAILY_LIMIT, remaining: DRAWING_DAILY_LIMIT - row.used, date };
}

async function refundDrawingUpload(env, userId, date) {
  try {
    await env.DB.prepare(
      `UPDATE drawing_usage SET used = CASE WHEN used > 0 THEN used - 1 ELSE 0 END
       WHERE user_id = ? AND usage_date = ?`
    ).bind(userId, date).run();
  } catch (error) {
    console.error("Drawing refund failed:", String(error));
  }
}

/* ------------------------------------------------------------------ */
/* Engineering Drawing Studio — geometry core                          */
/* ------------------------------------------------------------------ */
//
// ONE shared solid model drives every view:
//
//     solid = (union of "add" primitives) minus (union of "cut" primitives)
//
// Primitives are axis-aligned boxes and axis-aligned cylinders.
// Coordinates: x = width (left to right in the FRONT view),
//              y = height (up), z = depth (toward the viewer in the FRONT view).
//
// Nothing here is a per-view approximation. Every view (front, top, right,
// isometric, section) asks the same exact question about the same solid:
//   - "is this point inside the material?"           (membership)
//   - "which stretches of this ray are material?"    (ray intervals)
// Edges are found where the material genuinely changes, and hidden-line
// removal is done by casting rays at the real solid. Cuts are real boolean
// subtractions, never background-coloured shapes drawn on top.

const AXES = ["x", "y", "z"];
const rnd = n => Math.round(n * 100) / 100;
const fmtN = n => {
  const v = Math.round(n * 100) / 100;
  return String(Object.is(v, -0) ? 0 : v);
};
const cleanXml = s => String(s).replace(/[<>&"']/g, "");

function clampNum(value, lo, hi, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? Math.min(hi, Math.max(lo, n)) : fallback;
}

function drawingError(message, status = 400) {
  const error = new Error(message);
  error.status = status;
  return error;
}

/* ---------- Primitive helpers ---------- */

const otherAxes = a => AXES.filter(k => k !== a);
const axI = a => AXES.indexOf(a);

function extentsOf(p) {
  if (p.type === "box") {
    return { x: [p.x, p.x + p.w], y: [p.y, p.y + p.h], z: [p.z, p.z + p.d] };
  }
  if (p.type === "prism") {
    const o = otherAxes(p.axis);
    const e = {};
    e[p.axis] = [p.start, p.start + p.len];
    e[o[0]] = [Math.min(...p.profile.map(q => q[0])), Math.max(...p.profile.map(q => q[0]))];
    e[o[1]] = [Math.min(...p.profile.map(q => q[1])), Math.max(...p.profile.map(q => q[1]))];
    return e;
  }
  const R = p.type === "cone" ? Math.max(p.r1, p.r2) : p.r;
  const e = {};
  for (const a of AXES) {
    const half = a === p.axis ? p.len / 2 : R;
    e[a] = [p.c[a] - half, p.c[a] + half];
  }
  return e;
}

function minDimOf(p) {
  if (p.type === "box") return Math.min(p.w, p.h, p.d);
  if (p.type === "cyl") return Math.min(2 * p.r, p.len);
  if (p.type === "cone") return Math.min(p.len, 2 * Math.max(p.r1, p.r2));
  const e = extentsOf(p);
  return Math.min(...AXES.map(a => e[a][1] - e[a][0]));
}

function pointInPoly(poly, x, y) {
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const a = poly[i], b = poly[j];
    if ((a[1] > y) !== (b[1] > y) && x < ((b[0] - a[0]) * (y - a[1])) / (b[1] - a[1]) + a[0]) inside = !inside;
  }
  return inside;
}

const coneR = (p, x) => p.r1 + ((p.r2 - p.r1) * (x - (p.c[p.axis] - p.len / 2))) / p.len;

function boundsOfParts(parts) {
  const b = { x: [Infinity, -Infinity], y: [Infinity, -Infinity], z: [Infinity, -Infinity] };
  for (const p of parts) {
    const e = extentsOf(p);
    for (const a of AXES) {
      b[a][0] = Math.min(b[a][0], e[a][0]);
      b[a][1] = Math.max(b[a][1], e[a][1]);
    }
  }
  return b;
}

function primContains(p, q, tol = 0) {
  if (p.type === "box") {
    return (
      q[0] >= p.x - tol && q[0] <= p.x + p.w + tol &&
      q[1] >= p.y - tol && q[1] <= p.y + p.h + tol &&
      q[2] >= p.z - tol && q[2] <= p.z + p.d + tol
    );
  }
  if (p.type === "prism") {
    const a = axI(p.axis);
    if (q[a] < p.start - tol || q[a] > p.start + p.len + tol) return false;
    const o = otherAxes(p.axis).map(axI);
    return pointInPoly(p.profile, q[o[0]], q[o[1]]);
  }
  const ai = p.axis === "x" ? 0 : p.axis === "y" ? 1 : 2;
  if (Math.abs(q[ai] - p.c[p.axis]) > p.len / 2 + tol) return false;
  if (p.type === "cone") {
    let rr2 = 0;
    for (let i = 0; i < 3; i++) { if (i === ai) continue; const dv = q[i] - p.c[AXES[i]]; rr2 += dv * dv; }
    const R = coneR(p, q[ai]) + tol;
    return rr2 <= R * R;
  }
  let r2 = 0;
  for (let i = 0; i < 3; i++) {
    if (i === ai) continue;
    const dv = q[i] - p.c[AXES[i]];
    r2 += dv * dv;
  }
  const rr = p.r + tol;
  return r2 <= rr * rr;
}

/* ---------- The solid: membership ---------- */

// clip (optional) = { axis: "x"|"y"|"z", max: number } keeps coordinate <= max.
// It is used for sections: the part in front of the cutting plane is removed.
function makeSolid(parts, clip = null) {
  const adds = parts.filter(p => p.op === "add");
  const cuts = parts.filter(p => p.op === "cut");
  const ci = clip ? AXES.indexOf(clip.axis) : -1;

  return {
    parts, adds, cuts, clip, ci,
    inside(q) {
      if (ci >= 0 && q[ci] > clip.max) return false;
      let ok = false;
      for (const a of adds) if (primContains(a, q)) { ok = true; break; }
      if (!ok) return false;
      for (const c of cuts) if (primContains(c, q)) return false;
      return true;
    }
  };
}

/* ---------- Interval arithmetic along a ray ---------- */

function mergeIv(list) {
  const a = list.filter(i => i[1] > i[0]).sort((x, y) => x[0] - y[0]);
  const out = [];
  for (const i of a) {
    const last = out[out.length - 1];
    if (last && i[0] <= last[1]) last[1] = Math.max(last[1], i[1]);
    else out.push([i[0], i[1]]);
  }
  return out;
}

function subtractIv(a, b) {
  let cur = a.map(i => [i[0], i[1]]);
  for (const c of b) {
    const next = [];
    for (const i of cur) {
      if (c[1] <= i[0] || c[0] >= i[1]) { next.push(i); continue; }
      if (c[0] > i[0]) next.push([i[0], c[0]]);
      if (c[1] < i[1]) next.push([c[1], i[1]]);
    }
    cur = next;
  }
  return cur;
}

function slabIv(o, d, lo, hi, iv) {
  if (Math.abs(d) < 1e-12) return o >= lo && o <= hi ? iv : null;
  let t0 = (lo - o) / d;
  let t1 = (hi - o) / d;
  if (t0 > t1) [t0, t1] = [t1, t0];
  return [Math.max(iv[0], t0), Math.min(iv[1], t1)];
}

// Interval of the ray o + t*d (t in [-INF, INF]) that lies inside one primitive.
function simpleRayIv(p, o, d) {
  let iv = [-1e9, 1e9];

  if (p.type === "box") {
    const lo = [p.x, p.y, p.z];
    const hi = [p.x + p.w, p.y + p.h, p.z + p.d];
    for (let i = 0; i < 3; i++) {
      iv = slabIv(o[i], d[i], lo[i], hi[i], iv);
      if (!iv || iv[1] <= iv[0]) return null;
    }
    return iv;
  }

  const ai = p.axis === "x" ? 0 : p.axis === "y" ? 1 : 2;
  iv = slabIv(o[ai], d[ai], p.c[p.axis] - p.len / 2, p.c[p.axis] + p.len / 2, iv);
  if (!iv || iv[1] <= iv[0]) return null;

  let A = 0, B = 0, C = -p.r * p.r;
  for (let i = 0; i < 3; i++) {
    if (i === ai) continue;
    const oc = o[i] - p.c[AXES[i]];
    A += d[i] * d[i];
    B += 2 * oc * d[i];
    C += oc * oc;
  }

  if (A < 1e-12) return C <= 0 ? iv : null;
  const disc = B * B - 4 * A * C;
  if (disc <= 0) return null;
  const s = Math.sqrt(disc);
  const t0 = (-B - s) / (2 * A);
  const t1 = (-B + s) / (2 * A);
  const r = [Math.max(iv[0], t0), Math.min(iv[1], t1)];
  return r[1] > r[0] ? r : null;
}

function solveQuadLE(A, B, C, lo, hi) {
  const pts = [lo, hi];
  if (Math.abs(A) < 1e-12) {
    if (Math.abs(B) > 1e-12) pts.push(-C / B);
  } else {
    const disc = B * B - 4 * A * C;
    if (disc >= 0) {
      const s = Math.sqrt(disc);
      pts.push((-B - s) / (2 * A), (-B + s) / (2 * A));
    }
  }
  const ts = pts.filter(t => t >= lo && t <= hi).sort((a, b) => a - b);
  const out = [];
  for (let i = 0; i + 1 < ts.length; i++) {
    if (ts[i + 1] - ts[i] < 1e-12) continue;
    const m = (ts[i] + ts[i + 1]) / 2;
    if (A * m * m + B * m + C <= 0) {
      if (out.length && Math.abs(out[out.length - 1][1] - ts[i]) < 1e-12) out[out.length - 1][1] = ts[i + 1];
      else out.push([ts[i], ts[i + 1]]);
    }
  }
  return out;
}

function coneRayIvs(p, o, d) {
  const a = axI(p.axis);
  const lo = p.c[p.axis] - p.len / 2;
  const iv = slabIv(o[a], d[a], lo, lo + p.len, [-1e9, 1e9]);
  if (!iv || iv[1] <= iv[0]) return [];
  let A = 0, B = 0, C = 0;
  for (let i = 0; i < 3; i++) {
    if (i === a) continue;
    const oc = o[i] - p.c[AXES[i]];
    A += d[i] * d[i]; B += 2 * oc * d[i]; C += oc * oc;
  }
  const R0 = coneR(p, o[a]);
  const k = ((p.r2 - p.r1) * d[a]) / p.len;
  return solveQuadLE(A - k * k, B - 2 * R0 * k, C - R0 * R0, iv[0], iv[1]);
}

function prismRayIvs(p, o, d) {
  const a = axI(p.axis);
  const iv = slabIv(o[a], d[a], p.start, p.start + p.len, [-1e9, 1e9]);
  if (!iv || iv[1] <= iv[0]) return [];
  const [i1, i2] = otherAxes(p.axis).map(axI);
  const o1 = o[i1], o2 = o[i2], d1 = d[i1], d2 = d[i2];
  const poly = p.profile;
  if (Math.abs(d1) < 1e-12 && Math.abs(d2) < 1e-12) return pointInPoly(poly, o1, o2) ? [iv] : [];
  const ts = [iv[0], iv[1]];
  for (let i = 0; i < poly.length; i++) {
    const P = poly[i], Q = poly[(i + 1) % poly.length];
    const e0 = Q[0] - P[0], e1 = Q[1] - P[1];
    const den = d1 * e1 - d2 * e0;
    if (Math.abs(den) < 1e-12) continue;
    const w0 = P[0] - o1, w1 = P[1] - o2;
    const t = (w0 * e1 - w1 * e0) / den;
    const s = (w0 * d2 - w1 * d1) / den;
    if (s >= -1e-9 && s <= 1 + 1e-9 && t > iv[0] && t < iv[1]) ts.push(t);
  }
  ts.sort((x, y) => x - y);
  const out = [];
  for (let i = 0; i + 1 < ts.length; i++) {
    if (ts[i + 1] - ts[i] < 1e-12) continue;
    const m = (ts[i] + ts[i + 1]) / 2;
    if (pointInPoly(poly, o1 + d1 * m, o2 + d2 * m)) {
      if (out.length && Math.abs(out[out.length - 1][1] - ts[i]) < 1e-12) out[out.length - 1][1] = ts[i + 1];
      else out.push([ts[i], ts[i + 1]]);
    }
  }
  return out;
}

function primRayIvs(p, o, d) {
  if (p.type === "prism") return prismRayIvs(p, o, d);
  if (p.type === "cone") return coneRayIvs(p, o, d);
  const iv = simpleRayIv(p, o, d);
  return iv ? [iv] : [];
}

// Material intervals along a ray = (union of adds) - (union of cuts), clipped.
function rayMaterial(solid, o, d) {
  const add = [];
  for (const p of solid.adds) add.push(...primRayIvs(p, o, d));
  let mat = mergeIv(add);
  if (!mat.length) return mat;

  const cut = [];
  for (const p of solid.cuts) cut.push(...primRayIvs(p, o, d));
  if (cut.length) mat = subtractIv(mat, mergeIv(cut));

  if (solid.ci >= 0) {
    const i = solid.ci;
    const lim = solid.clip.max;
    const iv = slabIv(o[i], d[i], -1e9, lim, [-1e9, 1e9]);
    if (!iv) return [];
    mat = mat
      .map(m => [Math.max(m[0], iv[0]), Math.min(m[1], iv[1])])
      .filter(m => m[1] > m[0]);
  }
  return mat;
}

/* ---------- Check, repair and normalise what the AI returned ---------- */

const MAX_PARTS = 24;

function boxesTouch(a, b, tol) {
  const ea = extentsOf(a), eb = extentsOf(b);
  return AXES.every(k => ea[k][0] <= eb[k][1] + tol && eb[k][0] <= ea[k][1] + tol);
}

function polyIntersectSeg(a, b, c, d) {
  const o = (p, q, r) => (q[0] - p[0]) * (r[1] - p[1]) - (q[1] - p[1]) * (r[0] - p[0]);
  return o(a, b, c) * o(a, b, d) < 0 && o(c, d, a) * o(c, d, b) < 0;
}

function polySelfIntersects(pts) {
  const n = pts.length;
  for (let i = 0; i < n; i++) {
    for (let j = i + 2; j < n; j++) {
      if (i === 0 && j === n - 1) continue;
      if (polyIntersectSeg(pts[i], pts[(i + 1) % n], pts[j], pts[(j + 1) % n])) return true;
    }
  }
  return false;
}

function sanitizeModel(raw) {
  const warnings = [];
  const list = Array.isArray(raw?.parts) ? raw.parts.slice(0, MAX_PARTS) : [];
  if (Array.isArray(raw?.parts) && raw.parts.length > MAX_PARTS) {
    warnings.push(`Only the first ${MAX_PARTS} features were used.`);
  }

  if (raw?.recognised === false || raw?.recognized === false || raw?.image_kind === "not_a_part") {
    throw drawingError(
      "That image does not appear to show a machine part, engineering drawing or sketch. Please upload a clear picture of the component."
    );
  }

  const confidence = clampNum(raw?.confidence, 0, 1, NaN);
  if (Number.isFinite(confidence) && confidence < 0.2) {
    throw drawingError(
      "I could not identify the part confidently enough to draw it. Please upload a clearer image, ideally with printed dimensions or several views."
    );
  }

  const parts = [];
  let dropped = 0;

  for (const p of list) {
    const type = String(p?.type || "").toLowerCase();
    const op = String(p?.op || "add").toLowerCase() === "cut" ? "cut" : "add";

    if (type === "box") {
      const w = clampNum(p.w, 0, 5000, 0);
      const h = clampNum(p.h, 0, 5000, 0);
      const d = clampNum(p.d, 0, 5000, 0);
      if (w > 0 && h > 0 && d > 0) {
        parts.push({
          type: "box", op,
          x: clampNum(p.x, -5000, 5000, 0),
          y: clampNum(p.y, -5000, 5000, 0),
          z: clampNum(p.z, -5000, 5000, 0),
          w, h, d
        });
      } else dropped++;
    } else if (type === "cyl" || type === "cylinder") {
      const axis = String(p.axis || "z").toLowerCase();
      const r = clampNum(p.r ?? Number(p.diameter ?? p.dia) / 2, 0, 2500, 0);
      const len = clampNum(p.len ?? p.length, 0, 5000, 0);
      if (AXES.includes(axis) && r > 0 && len > 0) {
        parts.push({
          type: "cyl", op, axis, r, len,
          c: {
            x: clampNum(p.cx ?? p.x, -5000, 5000, 0),
            y: clampNum(p.cy ?? p.y, -5000, 5000, 0),
            z: clampNum(p.cz ?? p.z, -5000, 5000, 0)
          }
        });
      } else dropped++;
    } else if (type === "prism" || type === "extrude" || type === "profile") {
      const axis = String(p.axis || "z").toLowerCase();
      const raw = Array.isArray(p.profile) ? p.profile : Array.isArray(p.pts) ? p.pts : Array.isArray(p.points) ? p.points : [];
      let pts = raw
        .map(q => (Array.isArray(q) ? [Number(q[0]), Number(q[1])] : [Number(q?.x ?? q?.a), Number(q?.y ?? q?.b)]))
        .filter(q => Number.isFinite(q[0]) && Number.isFinite(q[1]))
        .map(q => [clampNum(q[0], -5000, 5000, 0), clampNum(q[1], -5000, 5000, 0)]);
      pts = pts.filter((q, i) => i === 0 || Math.hypot(q[0] - pts[i - 1][0], q[1] - pts[i - 1][1]) > 1e-6);
      if (pts.length > 1 && Math.hypot(pts[0][0] - pts[pts.length - 1][0], pts[0][1] - pts[pts.length - 1][1]) < 1e-6) pts.pop();
      const len = clampNum(p.len ?? p.length ?? p.thickness, 0, 5000, 0);
      const start = clampNum(p.start ?? p.s ?? p.from, -5000, 5000, 0);
      const area = pts.reduce((s, q, i) => { const n = pts[(i + 1) % pts.length]; return s + (q[0] * n[1] - n[0] * q[1]); }, 0) / 2;
      if (AXES.includes(axis) && pts.length >= 3 && pts.length <= 48 && len > 0 && Math.abs(area) > 1e-6 && !polySelfIntersects(pts)) {
        parts.push({ type: "prism", op, axis, start, len, profile: area < 0 ? pts.reverse() : pts });
      } else dropped++;
    } else if (type === "cone" || type === "taper" || type === "chamfer") {
      const axis = String(p.axis || "z").toLowerCase();
      const r1 = clampNum(p.r1 ?? p.r_start ?? Number(p.d1) / 2, 0, 2500, NaN);
      const r2 = clampNum(p.r2 ?? p.r_end ?? Number(p.d2) / 2, 0, 2500, NaN);
      const len = clampNum(p.len ?? p.length, 0, 5000, 0);
      if (AXES.includes(axis) && Number.isFinite(r1) && Number.isFinite(r2) && Math.max(r1, r2) > 0 && len > 0) {
        parts.push({
          type: "cone", op, axis, r1, r2, len,
          c: { x: clampNum(p.cx ?? p.x, -5000, 5000, 0), y: clampNum(p.cy ?? p.y, -5000, 5000, 0), z: clampNum(p.cz ?? p.z, -5000, 5000, 0) }
        });
      } else dropped++;
    } else dropped++;
  }

  if (dropped) warnings.push(`${dropped} unreadable feature(s) were ignored.`);

  let adds = parts.filter(p => p.op === "add");
  let cuts = parts.filter(p => p.op === "cut");

  if (!adds.length) {
    throw drawingError(
      "I could not work out the part's shape from that image. Please try a clearer picture, ideally with the dimensions printed on it."
    );
  }

  // The schema example must never come back as if it were the user's part.
  if (
    adds.length === 1 && !cuts.length && adds[0].type === "box" &&
    adds[0].w === 40 && adds[0].h === 20 && adds[0].d === 30 &&
    !adds[0].x && !adds[0].y && !adds[0].z
  ) {
    throw drawingError(
      "The AI returned a placeholder instead of the uploaded part. Please try again or upload a clearer image."
    );
  }

  // Remove duplicate and fully-enclosed additive solids (they only cause
  // overlapping outlines and add nothing to the shape).
  const key = p => JSON.stringify(
    p.type === "box"
      ? [p.type, rnd(p.x), rnd(p.y), rnd(p.z), rnd(p.w), rnd(p.h), rnd(p.d)]
      : p.type === "prism"
      ? [p.type, p.axis, rnd(p.start), rnd(p.len), p.profile.map(q => [rnd(q[0]), rnd(q[1])])]
      : p.type === "cone"
      ? [p.type, p.axis, rnd(p.r1), rnd(p.r2), rnd(p.len), rnd(p.c.x), rnd(p.c.y), rnd(p.c.z)]
      : [p.type, p.axis, rnd(p.r), rnd(p.len), rnd(p.c.x), rnd(p.c.y), rnd(p.c.z)]
  );
  const uniq = (arr) => { const s = new Set(); return arr.filter(p => (s.has(key(p)) ? false : (s.add(key(p)), true))); };
  adds = uniq(adds);
  cuts = uniq(cuts);

  const enclosed = (p, q) => {
    // p fully inside box q ?
    if (q.type !== "box" || p === q) return false;
    const e = extentsOf(p);
    return e.x[0] >= q.x - 1e-6 && e.x[1] <= q.x + q.w + 1e-6 &&
           e.y[0] >= q.y - 1e-6 && e.y[1] <= q.y + q.h + 1e-6 &&
           e.z[0] >= q.z - 1e-6 && e.z[1] <= q.z + q.d + 1e-6;
  };
  const before = adds.length;
  adds = adds.filter((p, i) => !adds.some((q, j) => j !== i && enclosed(p, q) && !(enclosed(q, p) && j > i)));
  if (adds.length < before) warnings.push("Redundant overlapping solids were merged.");

  // Normalise: move the part so its bounding box starts at the origin.
  const bb = boundsOfParts(adds);
  const shift = { x: bb.x[0], y: bb.y[0], z: bb.z[0] };
  const mv = p => {
    if (p.type === "box") return { ...p, x: p.x - shift.x, y: p.y - shift.y, z: p.z - shift.z };
    if (p.type === "prism") {
      const o = otherAxes(p.axis);
      return { ...p, start: p.start - shift[p.axis], profile: p.profile.map(q => [q[0] - shift[o[0]], q[1] - shift[o[1]]]) };
    }
    return { ...p, c: { x: p.c.x - shift.x, y: p.c.y - shift.y, z: p.c.z - shift.z } };
  };
  adds = adds.map(mv);
  cuts = cuts.map(mv);

  const body = makeSolid(adds);

  // A cut that removes nothing is an invented feature: drop it and say so.
  const keptCuts = [];
  for (const c of cuts) {
    const e = extentsOf(c);
    let hit = false;
    const N = 6;
    for (let i = 0; i <= N && !hit; i++)
      for (let j = 0; j <= N && !hit; j++)
        for (let k = 0; k <= N && !hit; k++) {
          const q = [
            e.x[0] + ((e.x[1] - e.x[0]) * i) / N,
            e.y[0] + ((e.y[1] - e.y[0]) * j) / N,
            e.z[0] + ((e.z[1] - e.z[0]) * k) / N
          ];
          if (primContains(c, q) && body.inside(q)) hit = true;
        }
    if (hit) keptCuts.push(c);
  }
  if (keptCuts.length < cuts.length) {
    warnings.push(`${cuts.length - keptCuts.length} cut(s) did not intersect the body and were ignored.`);
  }

  const parts2 = [...adds, ...keptCuts];
  const solid = makeSolid(parts2);
  const B = boundsOfParts(adds);

  // The material must not be eliminated entirely by the cuts.
  let inCount = 0, tot = 0;
  const G = 14;
  for (let i = 0; i < G; i++) for (let j = 0; j < G; j++) for (let k = 0; k < G; k++) {
    tot++;
    if (solid.inside([
      B.x[0] + ((B.x[1] - B.x[0]) * (i + 0.5)) / G,
      B.y[0] + ((B.y[1] - B.y[0]) * (j + 0.5)) / G,
      B.z[0] + ((B.z[1] - B.z[0]) * (k + 0.5)) / G
    ])) inCount++;
  }
  if (inCount / tot < 0.005) {
    throw drawingError(
      "The reconstructed cuts remove essentially all of the material, so the shape was rejected as unreliable. Please try a clearer image."
    );
  }

  // Disconnected solids usually mean a mis-placed feature.
  if (adds.length > 1) {
    const diag = Math.hypot(B.x[1] - B.x[0], B.y[1] - B.y[0], B.z[1] - B.z[0]);
    const seen = new Set([0]);
    const queue = [0];
    while (queue.length) {
      const i = queue.pop();
      adds.forEach((q, j) => { if (!seen.has(j) && boxesTouch(adds[i], q, diag * 1e-4)) { seen.add(j); queue.push(j); } });
    }
    if (seen.size < adds.length) {
      warnings.push("Some solids do not touch the main body; their position is uncertain.");
    }
  }

  const dimsPrinted = String(raw?.dimensions_source || "").toLowerCase() === "printed" && raw?.estimated !== true;
  const notes = Array.isArray(raw?.notes)
    ? raw.notes.map(n => cleanXml(String(n)).slice(0, 140)).filter(Boolean).slice(0, 5)
    : [];

  return {
    title: cleanXml(String(raw?.title || "Machine part")).slice(0, 60) || "Machine part",
    unit: String(raw?.unit || "").toLowerCase() === "in" ? "in" : "mm",
    estimated: !dimsPrinted,
    confidence: Number.isFinite(confidence) ? confidence : null,
    notes,
    warnings,
    parts: parts2
  };
}

/* ------------------------------------------------------------------ */
/* Edge extraction and hidden-line removal                             */
/* ------------------------------------------------------------------ */
//
// Candidate edge curves come from the geometry itself (box edges, plane/plane
// lines, cylinder rims, plane/cylinder lines and cylinder/cylinder
// intersection curves). A candidate piece is kept only where the FINAL solid
// really has an edge there, tested by sampling the material in a small ring
// around the point. So cut-away material, buried faces and flush joints never
// produce lines, and every intersection is exact.

const v3 = {
  dot: (a, b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2],
  cross: (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]],
  len: a => Math.hypot(a[0], a[1], a[2]),
  norm: a => { const l = Math.hypot(a[0], a[1], a[2]) || 1; return [a[0] / l, a[1] / l, a[2] / l]; },
  add: (a, b, s = 1) => [a[0] + b[0] * s, a[1] + b[1] * s, a[2] + b[2] * s]
};

const VIEW_DEFS = {
  front: { cam: [0, 0, 1], u: [1, 0, 0], v: [0, 1, 0], label: "FRONT VIEW" },
  top: { cam: [0, 1, 0], u: [1, 0, 0], v: [0, 0, -1], label: "TOP VIEW" },
  side: { cam: [1, 0, 0], u: [0, 0, -1], v: [0, 1, 0], label: "RIGHT SIDE VIEW" },
  iso: {
    cam: v3.norm([1, 1, 1]),
    u: [0.70710678, 0, -0.70710678],
    v: [-0.40824829, 0.81649658, -0.40824829],
    label: "ISOMETRIC VIEW"
  }
};

const TWO_PI = Math.PI * 2;
const q4 = n => Math.round(n * 1e4);

/* ---------- Solid context: tolerances and counters ---------- */

function makeContext(solidParts, clip) {
  const adds = solidParts.filter(p => p.op === "add");
  const B = boundsOfParts(adds);
  const diag = Math.hypot(B.x[1] - B.x[0], B.y[1] - B.y[0], B.z[1] - B.z[0]) || 1;

  let minDim = Infinity;
  for (const p of solidParts) {
    minDim = Math.min(minDim, minDimOf(p));
  }

  return {
    solid: makeSolid(solidParts, clip),
    B,
    diag,
    // ring radius for edge tests: small against the smallest feature
    eps: Math.max(diag * 1e-4, Math.min(diag * 0.0025, minDim * 0.05)),
    calls: 0,
    maxCalls: 900000
  };
}

function insideC(ctx, q) {
  if (++ctx.calls > ctx.maxCalls) {
    throw drawingError(
      "This part is too complex to draw reliably. Try an image of a simpler part or fewer features.",
      400
    );
  }
  return ctx.solid.inside(q);
}

/* ---------- Is there an edge here? (ring test) ---------- */

const RING_K = 16;
const RING_ANG = Array.from({ length: RING_K }, (_, k) => ((k + 0.5) * TWO_PI) / RING_K);

function ringBasis(tau) {
  const a = [Math.abs(tau[0]), Math.abs(tau[1]), Math.abs(tau[2])];
  const h = a[0] <= a[1] && a[0] <= a[2] ? [1, 0, 0] : a[1] <= a[2] ? [0, 1, 0] : [0, 0, 1];
  const d = v3.dot(h, tau);
  const e1 = v3.norm([h[0] - d * tau[0], h[1] - d * tau[1], h[2] - d * tau[2]]);
  return [e1, v3.cross(tau, e1)];
}

function ringOccupancy(ctx, P, tau, eps) {
  const [e1, e2] = ringBasis(tau);
  const bits = [];
  let n = 0;
  for (const phi of RING_ANG) {
    const c = Math.cos(phi) * eps;
    const s = Math.sin(phi) * eps;
    const q = [
      P[0] + e1[0] * c + e2[0] * s,
      P[1] + e1[1] * c + e2[1] * s,
      P[2] + e1[2] * c + e2[2] * s
    ];
    const b = insideC(ctx, q);
    bits.push(b);
    if (b) n++;
  }
  let flips = 0;
  for (let i = 0; i < RING_K; i++) if (bits[i] !== bits[(i + 1) % RING_K]) flips++;
  return { n, runs: flips / 2 };
}

// mode "edge": a crease (turn of more than ~30 degrees) or a concave corner.
// mode "boundary": any point on the surface (used for silhouettes).
function isEdgePoint(ctx, P, tau, mode) {
  // the test must hold at three ring sizes, so a line stops exactly where the
  // edge stops (a coarse ring would extend it past tangent joins)
  for (const k of [1, 1 / 12, 1 / 150]) {
    const { n, runs } = ringOccupancy(ctx, P, tau, ctx.eps * k);
    if (n === 0 || n === RING_K) return false;
    if (mode === "boundary") continue;
    if (!(runs >= 2 || n <= 6 || n >= 10)) return false;
  }
  return true;
}

/* ---------- Find the parameter intervals where a test is true ---------- */

function intervalsWhere(fn, t0, t1, breaks, M, iters = 13) {
  const span = t1 - t0;
  const base = [t0, t1];
  for (const b of breaks) if (b > t0 + 1e-9 && b < t1 - 1e-9) base.push(b);
  base.sort((a, b) => a - b);

  const ts = [base[0]];
  for (let i = 1; i < base.length; i++) {
    const gap = base[i] - base[i - 1];
    if (gap < 1e-9) continue;
    const k = Math.max(1, Math.ceil((gap / span) * M));
    for (let j = 1; j <= k; j++) ts.push(base[i - 1] + (gap * j) / k);
  }

  const mids = [];
  const st = [];
  for (let i = 0; i + 1 < ts.length; i++) {
    mids.push((ts[i] + ts[i + 1]) / 2);
    st.push(fn(mids[i]));
  }

  const out = [];
  let start = st[0] ? ts[0] : null;

  for (let i = 0; i + 1 < st.length; i++) {
    if (st[i] === st[i + 1]) continue;
    let lo = mids[i], hi = mids[i + 1];
    for (let it = 0; it < iters; it++) {
      const m = (lo + hi) / 2;
      if (fn(m) === st[i]) lo = m; else hi = m;
    }
    const tt = (lo + hi) / 2;
    if (st[i]) { out.push([start, tt]); start = null; } else start = tt;
  }
  if (start !== null) out.push([start, ts[ts.length - 1]]);
  return out.filter(i => i[1] - i[0] > 1e-9);
}

/* ---------- Candidate curves ---------- */

function uniqSorted(vals, tol) {
  const s = [...vals].sort((a, b) => a - b);
  const out = [];
  for (const v of s) if (!out.length || v - out[out.length - 1] > tol) out.push(v);
  return out;
}

function planeSet(parts, clip, tol) {
  const P = { x: [], y: [], z: [] };
  for (const p of parts) {
    if (p.type === "box") {
      P.x.push(p.x, p.x + p.w); P.y.push(p.y, p.y + p.h); P.z.push(p.z, p.z + p.d);
    } else if (p.type === "prism") {
      const o = otherAxes(p.axis);
      P[p.axis].push(p.start, p.start + p.len);
      p.profile.forEach((A, i) => {
        const B = p.profile[(i + 1) % p.profile.length];
        if (Math.abs(A[0] - B[0]) < 1e-9) P[o[0]].push(A[0]);
        if (Math.abs(A[1] - B[1]) < 1e-9) P[o[1]].push(A[1]);
      });
    } else {
      P[p.axis].push(p.c[p.axis] - p.len / 2, p.c[p.axis] + p.len / 2);
    }
  }
  if (clip) P[clip.axis].push(clip.max);
  for (const a of AXES) P[a] = uniqSorted(P[a], tol);
  return P;
}

function lineCandidate(axis, o1, o2, t0, t1, breaks, owners) {
  const others = AXES.filter(a => a !== axis);
  const dir = [0, 0, 0];
  dir[AXES.indexOf(axis)] = 1;
  return {
    type: "line", axis, c1: o1, c2: o2, t0, t1, breaks, owners, tau: () => dir,
    key: `${axis}|${q4(o1)}|${q4(o2)}`,
    at: t => {
      const p = [0, 0, 0];
      p[AXES.indexOf(axis)] = t;
      p[AXES.indexOf(others[0])] = o1;
      p[AXES.indexOf(others[1])] = o2;
      return p;
    }
  };
}

function circleCandidate(cyl, idx, v, thetaBreaks) {
  const a = AXES.indexOf(cyl.axis);
  const o = AXES.filter(k => k !== cyl.axis).map(k => AXES.indexOf(k));
  const c = [cyl.c.x, cyl.c.y, cyl.c.z];
  return {
    type: "circle", t0: 0, t1: TWO_PI, breaks: thetaBreaks, M: 96, owners: [idx],
    key: `c|${cyl.axis}|${q4(v)}|${q4(c[o[0]])}|${q4(c[o[1]])}|${q4(cyl.r)}`,
    at: th => {
      const p = [0, 0, 0];
      p[a] = v;
      p[o[0]] = c[o[0]] + cyl.r * Math.cos(th);
      p[o[1]] = c[o[1]] + cyl.r * Math.sin(th);
      return p;
    },
    tau: th => {
      const t = [0, 0, 0];
      t[o[0]] = -Math.sin(th);
      t[o[1]] = Math.cos(th);
      return t;
    }
  };
}

function curveCandidate(at, t0, t1, breaks, M, owners) {
  const h = (t1 - t0) * 1e-4;
  return {
    type: "curve", t0, t1, breaks, M, owners, key: null, at,
    tau: t => {
      const A = at(t + h), B = at(t - h);
      return A && B ? v3.norm([A[0] - B[0], A[1] - B[1], A[2] - B[2]]) : null;
    }
  };
}

function circleBreaks(cy, planes, o) {
  const br = [];
  planes[o[0]].forEach(k => { const c = (k - cy.c[o[0]]) / cy.r; if (Math.abs(c) <= 1) { const t = Math.acos(c); br.push(t, TWO_PI - t); } });
  planes[o[1]].forEach(k => {
    const s = (k - cy.c[o[1]]) / cy.r;
    if (Math.abs(s) <= 1) { const t = Math.asin(s); br.push((t + TWO_PI) % TWO_PI, (Math.PI - t + TWO_PI) % TWO_PI); }
  });
  return br;
}

function linearCandidate(A, B, breaks, owners) {
  const c = curveCandidate(
    t => [A[0] + (B[0] - A[0]) * t, A[1] + (B[1] - A[1]) * t, A[2] + (B[2] - A[2]) * t],
    0, 1, breaks, 10, owners
  );
  c.linear = true;
  return c;
}

// slanted plane (unit normal n, n.x = d) meeting a cylinder
function planeCylCandidates(n, d, cy, j, out, tol) {
  const a = AXES.indexOf(cy.axis);
  const ec = [0, 0, 0]; ec[a] = 1;
  const c = [cy.c.x, cy.c.y, cy.c.z];
  const nd = v3.dot(n, ec);
  const lo = cy.c[cy.axis] - cy.len / 2, hi = cy.c[cy.axis] + cy.len / 2;
  if (Math.abs(nd) < 1e-9) {
    const dl = v3.dot(n, c) - d;
    if (Math.abs(dl) >= cy.r - tol) return;
    const h = Math.sqrt(cy.r * cy.r - dl * dl);
    const perp = v3.norm(v3.cross(ec, n));
    const o = AXES.filter(k => k !== cy.axis);
    for (const s of [1, -1]) {
      const pt = [c[0] - dl * n[0] + s * h * perp[0], c[1] - dl * n[1] + s * h * perp[1], c[2] - dl * n[2] + s * h * perp[2]];
      out.push(lineCandidate(cy.axis, pt[AXES.indexOf(o[0])], pt[AXES.indexOf(o[1])], lo, hi, [], [j]));
    }
  } else {
    const o = AXES.filter(k => k !== cy.axis).map(k => AXES.indexOf(k));
    const at = th => {
      const p = c.slice();
      p[o[0]] += cy.r * Math.cos(th);
      p[o[1]] += cy.r * Math.sin(th);
      const t = (d - v3.dot(n, p)) / nd;
      if (Math.abs(t) > cy.len / 2 + tol) return null;
      p[a] += t;
      return p;
    };
    out.push(curveCandidate(at, 0, TWO_PI, [], 144, [j]));
  }
}

function prismCandidates(ctx, pr, idx, planes, cands, tol) {
  const a = AXES.indexOf(pr.axis);
  const o = otherAxes(pr.axis).map(axI);
  const a0 = pr.start, a1 = pr.start + pr.len;
  const prof = pr.profile;
  const pt3 = (P, lv) => { const q = [0, 0, 0]; q[a] = lv; q[o[0]] = P[0]; q[o[1]] = P[1]; return q; };
  const levels = [a0, ...planes[pr.axis].filter(v => v > a0 + tol && v < a1 - tol), a1];
  const brAx = planes[pr.axis].filter(v => v > a0 && v < a1);
  const parts = ctx.solid.parts;

  const coord = (i, v) => (i === 0 ? [v, null] : [null, v]);
  prof.forEach((A, i) => {
    const B = prof[(i + 1) % prof.length];
    // axial line at the vertex (keeps the line's two other coordinates in axis order)
    cands.push(lineCandidate(pr.axis, A[0], A[1], a0, a1, brAx, [idx]));

    // outline edge at every level
    const brE = [];
    for (const [k, ax] of [[0, o[0]], [1, o[1]]]) {
      const dv = B[k] - A[k];
      if (Math.abs(dv) < 1e-12) continue;
      for (const v of planes[AXES[ax]]) { const t = (v - A[k]) / dv; if (t > 0 && t < 1) brE.push(t); }
    }
    if (Math.abs(B[0] - A[0]) > 1e-9 && Math.abs(B[1] - A[1]) > 1e-9) {
      for (const lv of levels) cands.push(linearCandidate(pt3(A, lv), pt3(B, lv), brE, [idx]));
    }

    // slanted face against axis-parallel planes: axial lines
    for (const [k, ax] of [[0, o[0]], [1, o[1]]]) {
      const dv = B[k] - A[k];
      if (Math.abs(dv) < 1e-12) continue;
      for (const v of planes[AXES[ax]]) {
        const t = (v - A[k]) / dv;
        if (t <= 1e-9 || t >= 1 - 1e-9) continue;
        const P = [A[0] + (B[0] - A[0]) * t, A[1] + (B[1] - A[1]) * t];
        cands.push(lineCandidate(pr.axis, P[0], P[1], a0, a1, brAx, [idx]));
      }
    }

    // slanted face against cylinders
    const ex = B[0] - A[0], ey = B[1] - A[1];
    const L = Math.hypot(ex, ey);
    if (L < 1e-9 || Math.abs(ex) < 1e-9 || Math.abs(ey) < 1e-9) {
      // axis-aligned face: its plane is already in the plane set; same-axis cylinders meet it in axial lines
    }
    const n = [0, 0, 0];
    n[o[0]] = ey / L; n[o[1]] = -ex / L;
    const dd = n[o[0]] * A[0] + n[o[1]] * A[1];
    parts.forEach((cy, j) => {
      if (cy.type !== "cyl") return;
      if (cy.axis === pr.axis) {
        // axial lines where the face crosses the cylinder circle
        const cc = [cy.c[AXES[o[0]]], cy.c[AXES[o[1]]]];
        const fx = A[0] - cc[0], fy = A[1] - cc[1];
        const qa = ex * ex + ey * ey, qb = 2 * (fx * ex + fy * ey), qc = fx * fx + fy * fy - cy.r * cy.r;
        const disc = qb * qb - 4 * qa * qc;
        if (qa < 1e-12 || disc <= 0) return;
        for (const s of [-1, 1]) {
          const t = (-qb + s * Math.sqrt(disc)) / (2 * qa);
          if (t > 1e-9 && t < 1 - 1e-9) cands.push(lineCandidate(pr.axis, A[0] + ex * t, A[1] + ey * t, a0, a1, brAx, [idx, j]));
        }
      } else if (Math.abs(ex) > 1e-9 && Math.abs(ey) > 1e-9) {
        planeCylCandidates(n, dd, cy, j, cands, tol);
      }
    });
  });
  void coord;
}

function genCandidates(ctx) {
  const parts = ctx.solid.parts;
  const clip = ctx.solid.clip;
  const tol = ctx.diag * 1e-6;
  const planes = planeSet(parts, clip, tol);
  const cands = [];

  const ext = parts.map(extentsOf);
  const near = (r, v) => v >= r[0] - tol * 10 && v <= r[1] + tol * 10;

  // (1) Lines where two perpendicular planes meet (includes all box edges).
  for (const [k1, k2] of [["x", "y"], ["x", "z"], ["y", "z"]]) {
    const k3 = AXES.find(a => a !== k1 && a !== k2);
    for (const v1 of planes[k1]) {
      for (const v2 of planes[k2]) {
        const rel = [];
        ext.forEach((e, i) => { if (near(e[k1], v1) && near(e[k2], v2)) rel.push(i); });
        if (!rel.length) continue;
        let lo = Infinity, hi = -Infinity;
        const br = [];
        for (const i of rel) {
          lo = Math.min(lo, ext[i][k3][0]); hi = Math.max(hi, ext[i][k3][1]);
          br.push(ext[i][k3][0], ext[i][k3][1]);
        }
        if (clip && clip.axis === k3) hi = Math.min(hi, clip.max);
        if (hi - lo < tol) continue;
        // lines are built with the two remaining axes in AXES order
        const first = AXES.indexOf(k1) < AXES.indexOf(k2) ? v1 : v2;
        const second = AXES.indexOf(k1) < AXES.indexOf(k2) ? v2 : v1;
        cands.push(lineCandidate(k3, first, second, lo, hi, br, []));
      }
    }
  }

  const cyls = parts.map((p, i) => ({ p, i })).filter(o => o.p.type === "cyl");

  for (const { p: cy, i } of cyls) {
    const a = cy.axis;
    const o = AXES.filter(k => k !== a);
    const lo = cy.c[a] - cy.len / 2;
    const hi = cy.c[a] + cy.len / 2;

    // (2) Circles: wherever a plane crosses the cylinder at right angles
    //     (own end faces, and openings in faces of other solids).
    for (const v of planes[a]) {
      if (v < lo - tol || v > hi + tol) continue;
      cands.push(circleCandidate(cy, i, v, circleBreaks(cy, planes, o)));
    }

    // (3) Straight lines where a plane parallel to the axis meets the cylinder.
    const brA = [];
    ext.forEach(e => brA.push(e[a][0], e[a][1]));
    for (const k of o) {
      const m = o.find(x => x !== k);
      for (const v of planes[k]) {
        const dlt = v - cy.c[k];
        if (Math.abs(dlt) >= cy.r - tol) continue;
        const h = Math.sqrt(cy.r * cy.r - dlt * dlt);
        for (const sgn of [1, -1]) {
          const pos = {};
          pos[k] = v;
          pos[m] = cy.c[m] + sgn * h;
          cands.push(lineCandidate(a, pos[o[0]], pos[o[1]], lo, hi, brA, [i]));
        }
      }
    }
  }

  // (5) Extruded outlines (prisms): vertex lines, outline edges at every
  //     level, slanted faces against planes and cylinders.
  ctx.solid.parts.forEach((pr, idx) => {
    if (pr.type !== "prism") return;
    prismCandidates(ctx, pr, idx, planes, cands, tol);
  });

  // (6) Cones: rim circles at every plane crossing the axis.
  ctx.solid.parts.forEach((co, idx) => {
    if (co.type !== "cone") return;
    const a = co.axis;
    const o = AXES.filter(k => k !== a);
    const lo = co.c[a] - co.len / 2, hi = co.c[a] + co.len / 2;
    for (const v of planes[a]) {
      if (v < lo - tol || v > hi + tol) continue;
      const r = coneR(co, Math.min(hi, Math.max(lo, v)));
      if (r < 1e-9) continue;
      const cy = { axis: a, c: co.c, r };
      cands.push(circleCandidate(cy, idx, v, circleBreaks(cy, planes, o)));
    }
  });

  // (4) Cylinder / cylinder intersections.
  for (let ii = 0; ii < cyls.length; ii++) {
    for (let jj = ii + 1; jj < cyls.length; jj++) {
      const A = cyls[ii].p, B = cyls[jj].p;
      if (A.axis === B.axis) {
        const a = A.axis;
        const o = AXES.filter(k => k !== a);
        const dx = B.c[o[0]] - A.c[o[0]];
        const dy = B.c[o[1]] - A.c[o[1]];
        const D = Math.hypot(dx, dy);
        if (D < 1e-9 || D >= A.r + B.r || D <= Math.abs(A.r - B.r)) continue;
        const xa = (D * D + A.r * A.r - B.r * B.r) / (2 * D);
        const h = Math.sqrt(Math.max(0, A.r * A.r - xa * xa));
        const ex = dx / D, ey = dy / D;
        const lo = Math.min(A.c[a] - A.len / 2, B.c[a] - B.len / 2);
        const hi = Math.max(A.c[a] + A.len / 2, B.c[a] + B.len / 2);
        const br = [A.c[a] - A.len / 2, A.c[a] + A.len / 2, B.c[a] - B.len / 2, B.c[a] + B.len / 2];
        for (const sgn of [1, -1]) {
          cands.push(lineCandidate(
            a,
            A.c[o[0]] + ex * xa - sgn * ey * h,
            A.c[o[1]] + ey * xa + sgn * ex * h,
            lo, hi, br, [cyls[ii].i, cyls[jj].i]
          ));
        }
      } else {
        const a = AXES.indexOf(A.axis);
        const b = AXES.indexOf(B.axis);
        const m = 3 - a - b;
        const cA = [A.c.x, A.c.y, A.c.z];
        const cB = [B.c.x, B.c.y, B.c.z];
        const br = [];
        for (const s of [1, -1]) {
          const sn = (cB[m] + s * B.r - cA[m]) / A.r;
          if (Math.abs(sn) <= 1) { const t = Math.asin(sn); br.push((t + TWO_PI) % TWO_PI, (Math.PI - t + TWO_PI) % TWO_PI); }
        }
        for (const sgn of [1, -1]) {
          const at = th => {
            const p = [0, 0, 0];
            p[b] = cA[b] + A.r * Math.cos(th);
            p[m] = cA[m] + A.r * Math.sin(th);
            const rad = Math.max(0, B.r * B.r - (p[m] - cB[m]) * (p[m] - cB[m]));
            p[a] = cB[a] + sgn * Math.sqrt(rad);
            return p;
          };
          cands.push(curveCandidate(at, 0, TWO_PI, br, 180, [cyls[ii].i, cyls[jj].i]));
        }
      }
    }
  }

  return cands;
}

/* ---------- Keep only the genuine edge pieces ---------- */

function acceptCandidates(ctx, cands, mode) {
  const lines = new Map();
  const curves = [];

  for (const c of cands) {
    const M = c.M || 24;
    const iv = intervalsWhere(
      t => {
        const P = c.at(t);
        return P ? isEdgePoint(ctx, P, c.tau(t) || [1, 0, 0], mode) : false;
      },
      c.t0, c.t1, c.breaks, M
    );
    if (!iv.length) continue;

    if (c.type === "line") {
      const g = lines.get(c.key);
      if (g) { g.iv.push(...iv); c.owners.forEach(o => g.owners.add(o)); }
      else lines.set(c.key, { c, iv: [...iv], owners: new Set(c.owners) });
    } else {
      curves.push({ c, iv, owners: new Set(c.owners) });
    }
  }
  return { lines, curves };
}

// Edges that do not depend on the view (computed once per solid).
function baseEdges(ctx) {
  return acceptCandidates(ctx, genCandidates(ctx), "edge");
}

// Silhouette generators of cylinders seen from the side (view dependent).
function silhouetteEdges(ctx, V) {
  const cands = [];
  ctx.solid.parts.forEach((cy, i) => {
    const br = [];
    const addBr = ax => ctx.solid.parts.forEach(p => { const e = extentsOf(p)[ax]; br.push(e[0], e[1]); });

    if (cy.type === "cyl") {
      const ea = [0, 0, 0];
      ea[AXES.indexOf(cy.axis)] = 1;
      const n = v3.cross(ea, V.cam);
      if (v3.len(n) < 1e-6) return;
      const nn = v3.norm(n);
      const o = AXES.filter(k => k !== cy.axis);
      const lo = cy.c[cy.axis] - cy.len / 2;
      const hi = cy.c[cy.axis] + cy.len / 2;
      addBr(cy.axis);
      for (const s of [1, -1]) {
        const pos = o.map(k => cy.c[k] + s * cy.r * nn[AXES.indexOf(k)]);
        cands.push(lineCandidate(cy.axis, pos[0], pos[1], lo, hi, br, [i]));
      }
    } else if (cy.type === "cone") {
      const a = AXES.indexOf(cy.axis);
      const ea = [0, 0, 0]; ea[a] = 1;
      const m = v3.cross(v3.cross(ea, V.cam), ea); // camera direction perpendicular to the axis
      const mm = v3.len(m);
      if (mm < 1e-6) return;
      const slope = (cy.r2 - cy.r1) / cy.len;
      const cs = (slope * V.cam[a]) / mm;
      if (Math.abs(cs) > 1) return;
      const phi = Math.atan2(m[AXES.indexOf(otherAxes(cy.axis)[1])], m[AXES.indexOf(otherAxes(cy.axis)[0])]);
      const lo = cy.c[cy.axis] - cy.len / 2, hi = cy.c[cy.axis] + cy.len / 2;
      const o = otherAxes(cy.axis).map(axI);
      for (const s of [1, -1]) {
        const th = phi + s * Math.acos(cs);
        const mk = x => {
          const R = coneR(cy, x);
          const p = [cy.c.x, cy.c.y, cy.c.z];
          p[a] = x;
          p[o[0]] += R * Math.cos(th);
          p[o[1]] += R * Math.sin(th);
          return p;
        };
        cands.push(linearCandidate(mk(lo), mk(hi), [], [i]));
      }
    } else if (cy.type === "prism") {
      const o = otherAxes(cy.axis);
      const a0 = cy.start, a1 = cy.start + cy.len;
      const prof = cy.profile;
      addBr(cy.axis);
      const faceDot = (A, B) => {
        const ex = B[0] - A[0], ey = B[1] - A[1], L = Math.hypot(ex, ey) || 1;
        return (ey / L) * V.cam[axI(o[0])] + (-ex / L) * V.cam[axI(o[1])];
      };
      prof.forEach((P, k) => {
        const prev = prof[(k + prof.length - 1) % prof.length], next = prof[(k + 1) % prof.length];
        const d1 = faceDot(prev, P), d2 = faceDot(P, next);
        if (d1 * d2 <= 1e-12) cands.push(lineCandidate(cy.axis, P[0], P[1], a0, a1, br, [i]));
      });
    }
  });
  return acceptCandidates(ctx, cands, "boundary");
}

function mergeIntervals(iv) {
  const a = iv.map(i => [i[0], i[1]]).sort((x, y) => x[0] - y[0]);
  const out = [];
  for (const i of a) {
    const last = out[out.length - 1];
    if (last && i[0] <= last[1] + 1e-7) last[1] = Math.max(last[1], i[1]);
    else out.push(i);
  }
  return out;
}

/* ---------- Visibility ---------- */

function makeVisibility(ctx, V, solidForRays) {
  const dly = ctx.diag * 2e-4;
  const tmin = dly * 6;
  const offs = [[1, 1], [1, -1], [-1, 1], [-1, -1]].map(([a, b]) => [
    (V.u[0] * a + V.v[0] * b) * dly,
    (V.u[1] * a + V.v[1] * b) * dly,
    (V.u[2] * a + V.v[2] * b) * dly
  ]);

  // true when the point can be seen from the viewer
  return P => {
    for (const d of offs) {
      const o = [P[0] + d[0], P[1] + d[1], P[2] + d[2]];
      let blocked = false;
      for (const iv of rayMaterial(solidForRays, o, V.cam)) {
        if (iv[1] > tmin && iv[1] - Math.max(iv[0], tmin) > tmin) { blocked = true; break; }
      }
      if (!blocked) return true;
    }
    return false;
  };
}

function tessellate(c, a, b) {
  let n = 1;
  if (c.type === "circle") n = Math.max(2, Math.ceil(((b - a) / TWO_PI) * 120));
  else if (c.type === "curve" && !c.linear) n = Math.max(2, Math.ceil(((b - a) / (c.t1 - c.t0)) * 160));
  const pts = [];
  for (let i = 0; i <= n; i++) pts.push(c.at(a + ((b - a) * i) / n));
  return pts;
}

// Splits every accepted edge into visible and hidden pieces and projects
// them to the drawing plane.
function projectEdges(ctx, V, edges, isVisible) {
  const vis = [];
  const hid = [];
  const owners = new Set();
  const proj = P => [v3.dot(P, V.u), v3.dot(P, V.v)];
  const minLen = ctx.diag * 2e-3;
  const seenV = new Set();
  const seenH = new Set();

  const emit = (c, a, b, visible) => {
    const pts = tessellate(c, a, b).map(proj);
    let L = 0;
    for (let i = 1; i < pts.length; i++) L += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]);
    if (L < minLen) return;
    const k = pl => pl.map(q => `${Math.round(q[0] * 200)},${Math.round(q[1] * 200)}`);
    const kk = k([pts[0], pts[pts.length >> 1], pts[pts.length - 1]]);
    const key = (kk[0] < kk[2] ? kk : [kk[2], kk[1], kk[0]]).join("|");
    const seenSet = visible ? seenV : seenH;
    if (seenSet.has(key)) return;
    seenSet.add(key);
    (visible ? vis : hid).push(pts);
  };

  const process = (c, ivs, own) => {
    for (const [a, b] of ivs) {
      const M = c.type === "line" ? 20 : 60;
      const br = c.type === "line" ? c.breaks : [];
      const pieces = intervalsWhere(t => isVisible(c.at(t)), a, b, br, M, 11);
      // visible pieces
      for (const [x, y] of pieces) emit(c, x, y, true);
      // hidden pieces = complement
      let cur = a;
      for (const [x, y] of pieces) {
        if (x > cur + 1e-9) emit(c, cur, x, false);
        cur = y;
      }
      if (b > cur + 1e-9) emit(c, cur, b, false);
    }
    if (own) own.forEach(o => owners.add(o));
  };

  for (const g of edges.lines.values()) process(g.c, mergeIntervals(g.iv), g.owners);
  for (const g of edges.curves) process(g.c, mergeIntervals(g.iv), g.owners);

  return { visible: vis, hidden: hid, owners };
}

function combineEdges(a, b) {
  const lines = new Map();
  for (const src of [a, b]) {
    for (const [k, g] of src.lines) {
      const t = lines.get(k);
      if (t) { t.iv.push(...g.iv); g.owners.forEach(o => t.owners.add(o)); }
      else lines.set(k, { c: g.c, iv: [...g.iv], owners: new Set(g.owners) });
    }
  }
  return { lines, curves: [...a.curves, ...b.curves] };
}

/* ---------- Centre lines (2D, in drawing coordinates) ---------- */

function centreLines(ctx, V, owners) {
  const out = [];
  const seen = new Set();
  const proj = P => [v3.dot(P, V.u), v3.dot(P, V.v)];

  ctx.solid.parts.forEach((cy, i) => {
    if (cy.type !== "cyl" || !owners.has(i)) return;
    const ea = [0, 0, 0];
    ea[AXES.indexOf(cy.axis)] = 1;
    const along = Math.abs(v3.dot(ea, V.cam));
    const c = [cy.c.x, cy.c.y, cy.c.z];
    const ext = Math.max(cy.r * 0.15, ctx.diag * 0.012);

    if (along > 0.999) {
      // axis points at the viewer: a cross through the circle
      const p = proj(c);
      const key = `x|${q4(p[0])}|${q4(p[1])}|${q4(cy.r)}`;
      if (seen.has(key)) return;
      seen.add(key);
      const R = cy.r + ext;
      out.push([[p[0] - R, p[1]], [p[0] + R, p[1]]], [[p[0], p[1] - R], [p[0], p[1] + R]]);
    } else if (along < 1e-3) {
      const lo = cy.c[cy.axis] - cy.len / 2 - ext;
      const hi = cy.c[cy.axis] + cy.len / 2 + ext;
      const a = c.slice(); a[AXES.indexOf(cy.axis)] = lo;
      const b = c.slice(); b[AXES.indexOf(cy.axis)] = hi;
      const pa = proj(a), pb = proj(b);
      const key = `l|${q4(pa[0])}|${q4(pa[1])}|${q4(pb[0])}|${q4(pb[1])}`;
      if (seen.has(key)) return;
      seen.add(key);
      out.push([pa, pb]);
    }
  });
  return out;
}

/* ---------- Public: all edges of one view ---------- */

function computeView(ctx, base, V, withHidden) {
  const sil = silhouetteEdges(ctx, V);
  const all = combineEdges(base, sil);
  const isVisible = makeVisibility(ctx, V, ctx.solid);
  const e = projectEdges(ctx, V, all, isVisible);
  return {
    visible: e.visible,
    hidden: withHidden ? e.hidden : [],
    centres: centreLines(ctx, V, e.owners),
    owners: e.owners
  };
}

/* ------------------------------------------------------------------ */
/* Drawing renderer: line weights, dimensions, hatching, sheet layout  */
/* ------------------------------------------------------------------ */

const DRAWING_THEMES = {
  // clean white drawing sheet (default; also exports well)
  paper: { bg: "#ffffff", ink: "#111111" },
  // the original dark look, kept for the existing UI if it asks for it
  dark: { bg: "#0b0d10", ink: "#d3ec8a" }
};

const LW = { visible: 1.7, hidden: 0.9, centre: 0.7, dim: 0.7, hatch: 0.6, frame: 1.4, section: 1.9 };
const FONT = 'font-family="Arial, Helvetica, sans-serif"';

/* ---------- Choose the section plane ---------- */

function chooseSection(model) {
  const parts = model.parts;
  const B = boundsOfParts(parts.filter(p => p.op === "add"));
  const cyls = parts.filter(p => p.type === "cyl");
  let best = null;

  for (const cy of cyls) {
    for (const k of AXES) {
      if (k === cy.axis) continue;
      const val = cy.c[k];
      if (val <= B[k][0] + 1e-6 || val >= B[k][1] - 1e-6) continue;
      let score = 0;
      for (const o of cyls) {
        if (o.axis === k) continue;
        const tol = Math.max(o.r * 0.05, 1e-6);
        if (Math.abs(o.c[k] - val) <= tol) score += (o.op === "cut" ? 2 : 1) * 2 * o.r;
      }
      score += k === "z" ? 0.002 : k === "x" ? 0.001 : 0;
      if (!best || score > best.score) best = { axis: k, value: val, score };
    }
  }
  if (best) return { axis: best.axis, value: best.value };
  return { axis: "z", value: (B.z[0] + B.z[1]) / 2 };
}

const SECTION_VIEW = {
  z: { cam: [0, 0, 1], u: [1, 0, 0], v: [0, 1, 0] },
  x: { cam: [1, 0, 0], u: [0, 0, -1], v: [0, 1, 0] },
  y: { cam: [0, 1, 0], u: [1, 0, 0], v: [0, 0, -1] }
};

/* ---------- Hatching of the cut face ---------- */

function hatchSegments(fullSolid, V, planeVal, spacing, b2, diag) {
  const out = [];
  const s2 = Math.SQRT1_2;
  const d2 = [s2, s2];
  const n2 = [-s2, s2];
  const ks = [[b2.u0, b2.v0], [b2.u1, b2.v0], [b2.u0, b2.v1], [b2.u1, b2.v1]].map(p => n2[0] * p[0] + n2[1] * p[1]);
  const kmin = Math.floor(Math.min(...ks) / spacing) * spacing;
  const kmax = Math.max(...ks);
  const w = planeVal - diag * 1e-6;
  const R = diag * 4;

  const to3 = (u, v) => [
    V.u[0] * u + V.v[0] * v + V.cam[0] * w,
    V.u[1] * u + V.v[1] * v + V.cam[1] * w,
    V.u[2] * u + V.v[2] * v + V.cam[2] * w
  ];
  const dir3 = [
    V.u[0] * d2[0] + V.v[0] * d2[1],
    V.u[1] * d2[0] + V.v[1] * d2[1],
    V.u[2] * d2[0] + V.v[2] * d2[1]
  ];

  for (let k = kmin; k <= kmax + 1e-9; k += spacing) {
    const bu = n2[0] * k, bv = n2[1] * k;
    const o = to3(bu - d2[0] * R, bv - d2[1] * R);
    for (const iv of rayMaterial(fullSolid, o, dir3)) {
      const t0 = iv[0] - R, t1 = iv[1] - R;
      if (t1 - t0 < spacing * 0.15) continue;
      out.push([[bu + d2[0] * t0, bv + d2[1] * t0], [bu + d2[0] * t1, bv + d2[1] * t1]]);
    }
  }
  return out;
}

/* ---------- Small SVG helpers ---------- */

function boundsOfView(vd) {
  let u0 = Infinity, u1 = -Infinity, v0 = Infinity, v1 = -Infinity;
  for (const pl of [...vd.visible, ...vd.hidden]) {
    for (const p of pl) {
      u0 = Math.min(u0, p[0]); u1 = Math.max(u1, p[0]);
      v0 = Math.min(v0, p[1]); v1 = Math.max(v1, p[1]);
    }
  }
  return Number.isFinite(u0) ? { u0, u1, v0, v1 } : null;
}

function pathData(polys, T) {
  let d = "";
  for (const pl of polys) {
    if (pl.length < 2) continue;
    d += "M" + pl.map(p => `${rnd(T.X(p[0]))} ${rnd(T.Y(p[1]))}`).join("L");
  }
  return d;
}

function pathEl(polys, T, th, w, extra = "") {
  const d = pathData(polys, T);
  return d
    ? `<path d="${d}" fill="none" stroke="${th.ink}" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round" ${extra}/>`
    : "";
}

function txt(th, x, y, s, size = 12, extra = "") {
  return `<text x="${rnd(x)}" y="${rnd(y)}" fill="${th.ink}" ${FONT} font-size="${size}" paint-order="stroke" stroke="${th.bg}" stroke-width="3" stroke-linejoin="round" ${extra}>${cleanXml(s)}</text>`;
}

function ln(th, x1, y1, x2, y2, w = LW.dim, extra = "") {
  return `<line x1="${rnd(x1)}" y1="${rnd(y1)}" x2="${rnd(x2)}" y2="${rnd(y2)}" stroke="${th.ink}" stroke-width="${w}" ${extra}/>`;
}

function arrow(th, x, y, dx, dy) {
  const a = 8, b = 2.8, px = -dy, py = dx;
  return `<polygon points="${rnd(x)},${rnd(y)} ${rnd(x - dx * a + px * b)},${rnd(y - dy * a + py * b)} ${rnd(x - dx * a - px * b)},${rnd(y - dy * a - py * b)}" fill="${th.ink}"/>`;
}

function hDimSvg(th, x1, x2, y, label, refY1, refY2 = refY1) {
  const sgn = y >= refY1 ? 1 : -1;
  const inside = x2 - x1 > 34;
  let o = ln(th, x1, refY1 + sgn * 3, x1, y + sgn * 5) + ln(th, x2, refY2 + sgn * 3, x2, y + sgn * 5);
  if (inside) {
    o += ln(th, x1, y, x2, y) + arrow(th, x1, y, -1, 0) + arrow(th, x2, y, 1, 0);
    o += txt(th, (x1 + x2) / 2, y - 4, label, 12, 'text-anchor="middle"');
  } else {
    o += ln(th, x1 - 14, y, x2 + 14, y) + arrow(th, x1, y, 1, 0) + arrow(th, x2, y, -1, 0);
    o += txt(th, x2 + 18, y + 4, label, 12);
  }
  return o;
}

function vDimSvg(th, y1, y2, x, label, refX1, refX2 = refX1) {
  const sgn = x >= refX1 ? 1 : -1;
  const mid = (y1 + y2) / 2;
  const inside = Math.abs(y2 - y1) > 34;
  const top = Math.min(y1, y2), bot = Math.max(y1, y2);
  let o = ln(th, refX1 + sgn * 3, y1, x + sgn * 5, y1) + ln(th, refX2 + sgn * 3, y2, x + sgn * 5, y2);
  if (inside) {
    o += ln(th, x, top, x, bot) + arrow(th, x, top, 0, -1) + arrow(th, x, bot, 0, 1);
  } else {
    o += ln(th, x, top - 14, x, bot + 14) + arrow(th, x, top, 0, 1) + arrow(th, x, bot, 0, -1);
  }
  o += txt(th, 0, 0, label, 12, `text-anchor="middle" transform="translate(${rnd(x - 5)},${rnd(mid)}) rotate(-90)"`);
  return o;
}

/* ---------- Feature dimensions (diameters) ---------- */

function featureList(ctx, V, owners) {
  const proj = P => [v3.dot(P, V.u), v3.dot(P, V.v)];
  const seen = new Map();
  const list = [];

  ctx.solid.parts.forEach((cy, i) => {
    if (cy.type !== "cyl" || !owners.has(i)) return;
    const ea = [0, 0, 0];
    ea[AXES.indexOf(cy.axis)] = 1;
    const c = [cy.c.x, cy.c.y, cy.c.z];
    const along = Math.abs(v3.dot(ea, V.cam));

    if (along > 0.999) {
      const p = proj(c);
      const key = `o|${q4(p[0])}|${q4(p[1])}|${q4(cy.r)}`;
      if (seen.has(key)) return;
      seen.set(key, true);
      list.push({ kind: "circle", c: p, r: cy.r, op: cy.op });
    } else if (along < 1e-3) {
      const lo = c.slice(); lo[AXES.indexOf(cy.axis)] -= cy.len / 2;
      const hi = c.slice(); hi[AXES.indexOf(cy.axis)] += cy.len / 2;
      const a = proj(lo), b = proj(hi), m = proj(c);
      const horiz = Math.abs(v3.dot(ea, V.u)) > 0.999;
      const key = `s|${horiz}|${q4(m[0])}|${q4(m[1])}|${q4(cy.r)}|${q4(cy.len)}`;
      if (seen.has(key)) return;
      seen.set(key, true);
      list.push({
        kind: "side", horiz, c: m, r: cy.r, op: cy.op,
        e0: horiz ? Math.min(a[0], b[0]) : Math.min(a[1], b[1]),
        e1: horiz ? Math.max(a[0], b[0]) : Math.max(a[1], b[1])
      });
    }
  });
  return list;
}

function drawFeatureDims(th, feats, T, b, est, label) {
  let out = "";
  const tol = 1e-6 * (Math.abs(b.u1 - b.u0) + Math.abs(b.v1 - b.v0) + 1);
  const pre = est ? "~" : "";
  const lead = [[0.7071, -0.7071], [-0.7071, -0.7071], [0.7071, 0.7071], [-0.7071, 0.7071]];
  let li = 0, shown = 0;
  const used = { right: 0, left: 0 };

  // equal circles are dimensioned once with a count
  const groups = new Map();
  for (const f of feats) {
    if (f.kind !== "circle") continue;
    const k = q4(f.r);
    if (!groups.has(k)) groups.set(k, { f, n: 0 });
    groups.get(k).n++;
  }

  for (const { f, n } of groups.values()) {
    if (shown >= 6) break;
    const [dx, dy] = lead[li++ % 4];
    const cx = T.X(f.c[0]), cy = T.Y(f.c[1]), R = f.r * T.s;
    const tail = 22;
    out += ln(th, cx + dx * R, cy + dy * R, cx + dx * (R + tail), cy + dy * (R + tail));
    out += arrow(th, cx + dx * R, cy + dy * R, -dx, -dy);
    out += ln(th, cx + dx * (R + tail), cy + dy * (R + tail), cx + dx * (R + tail) + (dx > 0 ? 14 : -14), cy + dy * (R + tail));
    out += txt(th, cx + dx * (R + tail) + (dx > 0 ? 16 : -16), cy + dy * (R + tail) + 4,
      `${n > 1 ? n + "x " : ""}${pre}Ø${fmtDim(f.r * 2)}`, 12, `text-anchor="${dx > 0 ? "start" : "end"}"`);
    shown++;
  }

  const sides = new Map();
  for (const f of feats) {
    if (f.kind !== "side") continue;
    const k = `${f.horiz}|${q4(f.r)}`;
    if (!sides.has(k)) sides.set(k, f);
  }

  for (const f of sides.values()) {
    if (shown >= 6) break;
    if (f.horiz) {
      const y1 = T.Y(f.c[1] + f.r), y2 = T.Y(f.c[1] - f.r);
      let x, ref;
      if (f.e1 >= b.u1 - tol) { x = T.X(b.u1) + 62 + 30 * used.right++; ref = T.X(f.e1); }
      else if (f.e0 <= b.u0 + tol) { x = T.X(b.u0) - 62 - 30 * used.left++; ref = T.X(f.e0); }
      else continue;
      out += vDimSvg(th, y1, y2, x, `${pre}Ø${fmtDim(f.r * 2)}`, ref);
      shown++;
    }
  }
  return out;
}

const fmtDim = n => {
  const v = Math.round(n * 10) / 10;
  return String(Object.is(v, -0) ? 0 : v);
};

/* ---------- Cutting-plane line ---------- */

function cutLineSvg(th, cl, T, b) {
  // cl = { orient: "v"|"h", pos } in drawing coordinates
  const ext = 18;
  let out = "";
  const dash = 'stroke-dasharray="16 3 3 3"';
  if (cl.orient === "v") {
    const x = T.X(cl.pos), y0 = T.Y(b.v1) - ext, y1 = T.Y(b.v0) + ext;
    out += ln(th, x, y0, x, y1, 1.0, dash);
    out += ln(th, x, y0, x, y0 + 10, 2.4) + ln(th, x, y1 - 10, x, y1, 2.4);
    out += arrow(th, x + 12 * cl.dir, y0 + 5, cl.dir, 0) + arrow(th, x + 12 * cl.dir, y1 - 5, cl.dir, 0);
    out += txt(th, x, y0 - 6, "A", 13, 'text-anchor="middle" font-weight="bold"');
    out += txt(th, x, y1 + 16, "A", 13, 'text-anchor="middle" font-weight="bold"');
  } else {
    const y = T.Y(cl.pos), x0 = T.X(b.u0) - ext, x1 = T.X(b.u1) + ext;
    out += ln(th, x0, y, x1, y, 1.0, dash);
    out += ln(th, x0, y, x0 + 10, y, 2.4) + ln(th, x1 - 10, y, x1, y, 2.4);
    out += arrow(th, x0 + 5, y + 12 * cl.dir, 0, cl.dir) + arrow(th, x1 - 5, y + 12 * cl.dir, 0, cl.dir);
    out += txt(th, x0 - 10, y + 4, "A", 13, 'text-anchor="middle" font-weight="bold"');
    out += txt(th, x1 + 10, y + 4, "A", 13, 'text-anchor="middle" font-weight="bold"');
  }
  return out;
}

/* ---------- Draw one view into a group ---------- */

function drawView(vd, T, th, o) {
  const b = vd.bounds;
  let out = "";

  if (vd.hatch && vd.hatch.length) out += pathEl(vd.hatch, T, th, LW.hatch);
  if (vd.hidden.length) out += pathEl(vd.hidden, T, th, LW.hidden, 'stroke-dasharray="6 3.2" stroke-linecap="butt"');
  if (vd.centres.length) out += pathEl(vd.centres, T, th, LW.centre, 'stroke-dasharray="15 3 2.5 3" stroke-linecap="butt"');
  out += pathEl(vd.visible, T, th, o.section ? LW.section : LW.visible);

  if (o.cutLine) out += cutLineSvg(th, o.cutLine, T, b);

  if (o.dims) {
    const X0 = T.X(b.u0), X1 = T.X(b.u1), Yt = T.Y(b.v1), Yb = T.Y(b.v0);
    const pre = o.estimated ? "~" : "";
    out += hDimSvg(th, X0, X1, Yb + 30, pre + fmtDim(b.u1 - b.u0), Yb);
    out += vDimSvg(th, Yt, Yb, X1 + 30, pre + fmtDim(b.v1 - b.v0), X1);
    out += drawFeatureDims(th, vd.feats || [], T, b, o.estimated);
  }
  return out;
}

/* ---------- Frame and title block for single views ---------- */

function noteLines(model) {
  const n = [];
  if (model.estimated) n.push("DIMENSIONS ESTIMATED FROM IMAGE PROPORTIONS - NOT FOR MANUFACTURE");
  else n.push("DIMENSIONS TAKEN FROM THE UPLOADED IMAGE - VERIFY BEFORE USE");
  for (const w of model.warnings || []) n.push(w);
  for (const w of model.notes || []) n.push(w);
  return n.slice(0, 4);
}

function frameSvg(model, projection, label, body, themeName = "paper") {
  const th = DRAWING_THEMES[themeName] || DRAWING_THEMES.paper;
  const pj = projection === "first" ? "FIRST-ANGLE" : "THIRD-ANGLE";
  const notes = noteLines(model);
  const conf = model.confidence != null ? ` | Confidence ${Math.round(model.confidence * 100)}%` : "";

  let noteSvg = "";
  notes.slice(0, 2).forEach((t, i) => {
    noteSvg += txt(th, 24, 506 + i * 14, t.length > 120 ? t.slice(0, 117) + "..." : t, 10);
  });

  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${DRAW.W} ${DRAW.H}" width="100%" role="img" aria-label="${cleanXml(model.title)} - ${cleanXml(label)}" style="background:${th.bg};border-radius:12px;display:block">` +
    `<rect x="14" y="14" width="${DRAW.W - 28}" height="${DRAW.H - 28}" fill="none" stroke="${th.ink}" stroke-width="${LW.frame}"/>` +
    body +
    txt(th, DRAW.W / 2, 540, label, 14, 'text-anchor="middle" font-weight="bold"') +
    noteSvg +
    ln(th, 14, 562, DRAW.W - 14, 562, 1) +
    txt(th, 24, 584, `${model.title} | Units: ${model.unit} | ${pj} PROJECTION | Fitted scale, not 1:1${conf}`, 10) +
    txt(th, DRAW.W - 24, 584, "AI-generated from image" + (model.estimated ? " (dimensions estimated)" : "") + " - verify before use", 10, 'text-anchor="end"') +
    `</svg>`
  );
}

const DRAW = {
  W: 900,
  H: 600,
  area: { x: 120, y: 60, w: 600, h: 400 }
};

function fitTransform(b, area, scaleOverride) {
  const bw = Math.max(b.u1 - b.u0, 1e-6);
  const bh = Math.max(b.v1 - b.v0, 1e-6);
  const s = scaleOverride || Math.min(area.w / bw, area.h / bh);
  const ox = area.x + (area.w - bw * s) / 2;
  const oy = area.y + (area.h - bh * s) / 2;
  return { s, X: u => ox + (u - b.u0) * s, Y: v => oy + (b.v1 - v) * s, w: bw * s, h: bh * s };
}

function singleViewSvg(model, projection, label, vd, o, themeName) {
  const th = DRAWING_THEMES[themeName] || DRAWING_THEMES.paper;
  if (!vd || !vd.bounds) {
    return frameSvg(model, projection, label,
      txt(th, DRAW.W / 2, DRAW.H / 2, "No reliable geometry could be reconstructed", 14, 'text-anchor="middle"'), themeName);
  }
  const T = fitTransform(vd.bounds, DRAW.area);
  return frameSvg(model, projection, label, drawView(vd, T, th, { ...o, estimated: model.estimated }), themeName);
}

/* ---------- Multi-view sheet ---------- */

function sheetSvg(model, projection, views, sec, themeName) {
  const th = DRAWING_THEMES[themeName] || DRAWING_THEMES.paper;
  const W = 1400, H = 990, M = 30, GAP = 96, TB = 120;
  const first = projection === "first";
  const bF = views.front.bounds, bT = views.top.bounds, bS = views.side.bounds, bI = views.iso.bounds, bX = sec.vd.bounds;
  const dim = b => [b.u1 - b.u0, b.v1 - b.v0];
  const [wF, hF] = dim(bF), [wT, dT] = dim(bT), [wS] = dim(bS), [wX, hX] = dim(bX);

  const availW = W - 2 * M - 4 * GAP + GAP * 0.2;
  const availH = H - 2 * M - TB - 50 - GAP - 90;
  const s = Math.min(availW / (Math.max(wF, wT) + wS + wX), availH / (hF + dT));

  const colF = Math.max(wF, wT) * s, colS = wS * s;
  const colX = Math.max(wX * s, 300);
  const rowTop = (first ? hF : dT) * s;
  const rowBot = (first ? dT : hF) * s;

  // grid origin so the whole thing is centred
  const totalW = colF + colS + colX + 2 * GAP;
  const x0 = (W - totalW) / 2;
  const y0 = M + 50;

  const cols = first
    ? { s: x0, f: x0 + colS + GAP, x: x0 + colS + colF + 2 * GAP }
    : { f: x0, s: x0 + colF + GAP, x: x0 + colF + colS + 2 * GAP };
  const rows = { r0: y0, r1: y0 + rowTop + GAP };

  const place = (b, cx, cy, cw, ch, scale) => {
    const w = (b.u1 - b.u0) * scale, h = (b.v1 - b.v0) * scale;
    const ox = cx + (cw - w) / 2, oy = cy + (ch - h) / 2;
    return { s: scale, X: u => ox + (u - b.u0) * scale, Y: v => oy + (b.v1 - v) * scale };
  };

  let body = "";
  const lab = (t, T, b, below) => {
    const cx = T.X((b.u0 + b.u1) / 2);
    const y = below ? T.Y(b.v0) + 62 : T.Y(b.v1) - 28;
    return txt(th, cx, y, t, 13, 'text-anchor="middle" font-weight="bold"');
  };

  const rowF = first ? rows.r0 : rows.r1;           // front row
  const rowH = first ? rowTop : rowBot;
  const rowT = first ? rows.r1 : rows.r0;           // top-view row
  const rowTh = first ? rowBot : rowTop;

  const TF = place(bF, cols.f, rowF, colF, rowH, s);
  const TT = place(bT, cols.f, rowT, colF, rowTh, s);
  const TS = place(bS, cols.s, rowF, colS, rowH, s);
  const isoS = Math.min(colX / (bI.u1 - bI.u0), (rowTop + 0) / (bI.v1 - bI.v0), s * 1.1) * 0.9;
  const TI = place(bI, cols.x, rows.r0, colX, rowTop, isoS);
  const sX = Math.min(s, colX / wX, rowBot / hX);
  const TX = place(bX, cols.x, rows.r1, colX, rowBot, sX);

  body += drawView(views.front, TF, th, { dims: true, estimated: model.estimated, cutLine: sec.cut && sec.cut.view === "front" ? sec.cut.line : null });
  body += drawView(views.top, TT, th, { dims: true, estimated: model.estimated, cutLine: sec.cut && sec.cut.view === "top" ? sec.cut.line : null });
  body += drawView(views.side, TS, th, { dims: true, estimated: model.estimated });
  body += drawView(views.iso, TI, th, { dims: false });
  const secV = { ...sec.vd, hatch: sec.vd.hatchFn ? sec.vd.hatchFn(sX) : sec.vd.hatch };
  body += drawView(secV, TX, th, { dims: false, section: true });

  body += lab("FRONT VIEW", TF, bF, true);
  body += lab("TOP VIEW", TT, bT, true);
  body += lab("RIGHT SIDE VIEW", TS, bS, true);
  body += lab("ISOMETRIC (NOT TO SCALE)", TI, bI, true);
  body += lab("SECTION A-A", TX, bX, true);

  // title block
  const tbY = H - M - TB;
  const pj = first ? "FIRST-ANGLE" : "THIRD-ANGLE";
  const notes = noteLines(model);
  let tb = `<rect x="${M}" y="${tbY}" width="${W - 2 * M}" height="${TB}" fill="none" stroke="${th.ink}" stroke-width="1.2"/>`;
  tb += ln(th, W - M - 420, tbY, W - M - 420, H - M, 1.2);
  tb += ln(th, W - M - 420, tbY + 40, W - M, tbY + 40, 1);
  tb += txt(th, W - M - 408, tbY + 27, model.title, 18, 'font-weight="bold"');
  tb += txt(th, W - M - 408, tbY + 62, `Units: ${model.unit}   |   ${pj} PROJECTION   |   Scale: fitted (not 1:1)`, 11);
  tb += txt(th, W - M - 408, tbY + 82, "AI-generated from uploaded image" + (model.confidence != null ? `   |   Confidence ${Math.round(model.confidence * 100)}%` : ""), 11);
  tb += txt(th, W - M - 408, tbY + 102, model.estimated ? "Dimensions ESTIMATED (~) - verify before use" : "Dimensions from image - verify before use", 11);
  notes.forEach((t, i) => { tb += txt(th, M + 14, tbY + 24 + i * 18, (i === 0 ? "" : "- ") + (t.length > 110 ? t.slice(0, 107) + "..." : t), 11); });

  // projection symbol
  const sx = W - M - 520, sy = tbY + TB / 2;
  const g = first ? 1 : -1;
  tb += `<g transform="translate(${sx},${sy}) scale(${g},1)" fill="none" stroke="${th.ink}" stroke-width="1.2">` +
    `<path d="M-30 -14 L10 -22 L10 22 L-30 14 Z"/><line x1="-34" y1="0" x2="16" y2="0" stroke-dasharray="8 2 2 2" stroke-width="0.7"/>` +
    `<circle cx="42" cy="0" r="22"/><circle cx="42" cy="0" r="12"/><line x1="14" y1="0" x2="72" y2="0" stroke-dasharray="8 2 2 2" stroke-width="0.7"/><line x1="42" y1="-28" x2="42" y2="28" stroke-dasharray="8 2 2 2" stroke-width="0.7"/></g>`;

  return (
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="100%" role="img" aria-label="${cleanXml(model.title)} - engineering drawing sheet" style="background:${th.bg};border-radius:12px;display:block">` +
    `<rect x="${M - 14}" y="${M - 14}" width="${W - 2 * M + 28}" height="${H - 2 * M + 28}" fill="none" stroke="${th.ink}" stroke-width="${LW.frame}"/>` +
    body + tb + `</svg>`
  );
}

/* ---------- Build every view from the one shared model ---------- */

function buildAllViews(model, projection, themeName = "paper") {
  const ctx = makeContext(model.parts, null);
  const base = baseEdges(ctx);

  const mk = key => {
    const V = VIEW_DEFS[key];
    const r = computeView(ctx, base, V, key !== "iso");
    const vd = { ...r, V };
    vd.bounds = boundsOfView(vd);
    vd.feats = key === "iso" ? [] : featureList(ctx, V, r.owners);
    return vd;
  };

  const views = { front: mk("front"), top: mk("top"), side: mk("side"), iso: mk("iso") };
  for (const k of Object.keys(views)) {
    if (!views[k].bounds) {
      throw drawingError("No reliable geometry could be reconstructed from that image. Please try a clearer picture.");
    }
  }

  /* section */
  const plane = chooseSection(model);
  const SV = { ...SECTION_VIEW[plane.axis], label: "SECTION A-A" };
  const ctxS = makeContext(model.parts, { axis: plane.axis, max: plane.value });
  const baseS = baseEdges(ctxS);
  const rs = computeView(ctxS, baseS, SV, false);
  const secVd = { ...rs, V: SV };
  secVd.bounds = boundsOfView(secVd);

  if (secVd.bounds) {
    const T0 = fitTransform(secVd.bounds, DRAW.area);
    const sectionBounds = secVd.bounds;
    // hatch spacing is set in drawing units from a fixed on-paper spacing
    secVd.hatchFn = scale => hatchSegments(ctx.solid, SV, plane.value, 5.2 / scale, sectionBounds, ctx.diag);
    secVd.hatch = secVd.hatchFn(T0.s);
  }
  secVd.centres = []; // centre lines are omitted in the section for clarity

  // where the cutting plane is shown
  let cut = null;
  const Bm = ctx.B;
  if (plane.axis === "z") cut = { view: "top", line: { orient: "h", pos: -plane.value, dir: 1 } };
  else if (plane.axis === "x") cut = { view: "front", line: { orient: "v", pos: plane.value, dir: 1 } };
  else cut = { view: "front", line: { orient: "h", pos: plane.value, dir: 1 } };
  void Bm;

  const out = {};
  out.front = singleViewSvg(model, projection, VIEW_DEFS.front.label, views.front, { dims: true, cutLine: cut.view === "front" ? cut.line : null }, themeName);
  out.top = singleViewSvg(model, projection, VIEW_DEFS.top.label, views.top, { dims: true, cutLine: cut.view === "top" ? cut.line : null }, themeName);
  out.side = singleViewSvg(model, projection, VIEW_DEFS.side.label, views.side, { dims: true }, themeName);
  out.iso = singleViewSvg(model, projection, "ISOMETRIC VIEW (NOT TO SCALE)", views.iso, { dims: false }, themeName);

  if (secVd.bounds) {
    const label = "SECTION A-A  (cut through " + plane.axis.toUpperCase() + " = " + (model.estimated ? "~" : "") + fmtDim(plane.value) + ")";
    out.section = singleViewSvg(model, projection, label, { ...secVd, feats: [] }, { dims: true, section: true }, themeName);
    out.sheet = sheetSvg(model, projection, views, { vd: { ...secVd, feats: [] }, cut }, themeName);
  } else {
    out.section = out.front;
    out.sheet = null;
  }
  out.aux = null;
  out.sectionPlane = { axis: plane.axis, value: rnd(plane.value) };
  return out;
}

/* ------------------------------------------------------------------ */
/* Engineering Drawing Studio — AI prompt and image validation         */
/* ------------------------------------------------------------------ */

const DRAWING_PROMPT = [
  "You are Malie AI Engineering Drawing Studio, a mechanical engineering drawing interpretation system.",
  "Look carefully at the ATTACHED image. It may be a photograph, hand sketch, orthographic drawing, dimensioned drawing, or isometric drawing of ONE machine part.",
  "Reconstruct the specific object that is actually shown, following its real outline. Never substitute a generic or example object, and never copy placeholder values from this prompt.",
  "If the image is not a machine part, drawing or sketch (or is blank/unreadable), set recognised to false and parts to an empty array.",
  "Step 1: decide image_kind (orthographic, isometric, sketch, photo, other, not_a_part). Step 2: note the views you can see and any printed dimension text. Step 3: trace the part's true outline and model it.",
  "Coordinates: x = width (left to right in the FRONT view), y = height (up), z = depth (toward the viewer in the front view). Use one origin and one unit for every primitive.",
  "SHAPE TYPES (all are solids; op is \"add\" for material, \"cut\" for holes, bores, slots, pockets, grooves; list adds first, then cuts):",
  "1. prism = an OUTLINE EXTRUDED along an axis. Use it for any plate, bracket, lug, gusset, flange, wedge, rounded or angled shape. {\"type\":\"prism\",\"op\":\"add\",\"axis\":\"x|y|z\",\"start\":<where it begins along the axis>,\"len\":<extrusion length>,\"profile\":[[a,b],[a,b],...]}. The profile is the closed outline, in order, as 3 to 40 points. Profile coordinates are [x,y] when axis is z, [x,z] when axis is y, and [y,z] when axis is x. Trace slanted edges with their end points. Approximate rounded corners with 3-6 points per corner.",
  "2. box = {\"type\":\"box\",\"op\":\"add\",\"x\":<min x>,\"y\":<min y>,\"z\":<min z>,\"w\":<size x>,\"h\":<size y>,\"d\":<size z>}. Use only for plain rectangular blocks.",
  "3. cyl = straight cylinder (shaft, boss, bore, drilled hole): {\"type\":\"cyl\",\"op\":\"add|cut\",\"axis\":\"x|y|z\",\"cx\":<centre x>,\"cy\":<centre y>,\"cz\":<centre z>,\"r\":<radius>,\"len\":<length>}.",
  "4. cone = tapered or chamfered round segment (taper, chamfer, countersink): {\"type\":\"cone\",\"op\":\"add|cut\",\"axis\":\"x|y|z\",\"cx\":<centre x>,\"cy\":<centre y>,\"cz\":<centre z>,\"r1\":<radius at the low end>,\"r2\":<radius at the high end>,\"len\":<length>}.",
  "Build a stepped turned part (shaft, bush, pulley) from coaxial cyl and cone segments that touch end to end. Build a flat part by extruding its outline with prism, then add cylinder cuts for holes (make the cut longer than the material).",
  "A hole is a cut, never an additive cylinder. A boss is an add, never a cut. Touching solids must share faces exactly so they join. Do not add hidden or invented features.",
  "Use all views in the image together to build ONE consistent 3D shape. If only one view exists, use it as the outline and infer depth conservatively, saying so in notes.",
  "Dimensions: if numbers are printed and readable, use them and set dimensions_source to \"printed\". Otherwise estimate from proportions, set dimensions_source to \"estimated\", and never pretend estimates are exact.",
  "Describe any ambiguity in notes (max 4 short strings). Set confidence 0 to 1 for how well your model matches the image. Use at most 24 shapes.",
  "Return ONLY valid JSON with double quotes, no Markdown, no text outside the JSON, in this shape:",
  "{\"recognised\":true,\"image_kind\":\"...\",\"title\":\"short part name\",\"unit\":\"mm\",\"dimensions_source\":\"printed|estimated\",\"confidence\":0.0,\"notes\":[\"...\"],\"parts\":[ ...shapes as described above... ]}"
].join("\n");


const DRAWING_REFINE_PROMPT = [
  "You are checking and correcting a 3D reconstruction of the part shown in the ATTACHED image.",
  "Below is a first attempt as JSON. Compare it carefully with the image and fix every mistake:",
  "- the overall outline and proportions (width : height : depth) must match the image;",
  "- slanted, curved, stepped or notched edges must follow the image (use prism outlines, not plain boxes, when the outline is not rectangular);",
  "- every visible hole, bore, slot and groove must be present, in the right place and size, and no feature may be invented;",
  "- additive shapes must touch and join; cuts must pass through the material.",
  "If the first attempt is already correct, return it unchanged. Keep the same schema and the same coordinate rules (x = width, y = height, z = depth).",
  "Lower your confidence if you are unsure, and put any doubts in notes. Return ONLY the corrected JSON, nothing else.",
  "FIRST ATTEMPT:"
].join("\n");

// Checks the uploaded file really is a decodable PNG, JPEG or WebP of usable size.
function validateDrawingImage(image) {
  const bad = msg => { throw drawingError(msg, 400); };
  const m = /^data:image\/(png|jpeg|webp);base64,([A-Za-z0-9+/=\s]+)$/.exec(image);
  if (image.length > 5_000_000 || !m) {
    bad("Please upload a PNG, JPG or WebP image under about 3.5 MB.");
  }

  let bytes;
  try {
    const bin = atob(m[2].replace(/\s/g, ""));
    bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  } catch {
    bad("That image file appears to be corrupted. Please upload it again.");
  }
  if (bytes.length < 1024) bad("That image is too small or empty to read. Please upload a clearer picture.");

  const kind = m[1];
  let w = 0, h = 0;
  const u16 = (i) => (bytes[i] << 8) | bytes[i + 1];

  if (kind === "png") {
    if (!(bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)) bad("That file is not a valid PNG image.");
    w = ((bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19]) >>> 0;
    h = ((bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23]) >>> 0;
  } else if (kind === "jpeg") {
    if (!(bytes[0] === 0xff && bytes[1] === 0xd8)) bad("That file is not a valid JPEG image.");
    let i = 2;
    while (i + 9 < bytes.length) {
      if (bytes[i] !== 0xff) { i++; continue; }
      const mk = bytes[i + 1];
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) {
        h = u16(i + 5); w = u16(i + 7); break;
      }
      i += 2 + u16(i + 2);
    }
  } else {
    const tag = String.fromCharCode(...bytes.slice(0, 4)) + String.fromCharCode(...bytes.slice(8, 12));
    if (tag !== "RIFFWEBP") bad("That file is not a valid WebP image.");
    const t = String.fromCharCode(...bytes.slice(12, 16));
    if (t === "VP8X") { w = 1 + (bytes[24] | (bytes[25] << 8) | (bytes[26] << 16)); h = 1 + (bytes[27] | (bytes[28] << 8) | (bytes[29] << 16)); }
    else if (t === "VP8 ") { w = (bytes[26] | (bytes[27] << 8)) & 0x3fff; h = (bytes[28] | (bytes[29] << 8)) & 0x3fff; }
    else if (t === "VP8L") { const b = bytes[21] | (bytes[22] << 8) | (bytes[23] << 16) | (bytes[24] << 24); w = (b & 0x3fff) + 1; h = ((b >> 14) & 0x3fff) + 1; }
  }

  if (!w || !h) bad("I could not read that image. Please upload a standard PNG, JPG or WebP file.");
  if (w < 64 || h < 64) bad("That image is too small (under 64 px). Please upload a larger, clearer picture.");
  if (w > 12000 || h > 12000) bad("That image is extremely large. Please upload one under about 4000 px wide.");
  return { width: w, height: h, type: kind };
}

/* ------------------------------------------------------------------ */
/* CAD code validation                                                 */
/* ------------------------------------------------------------------ */

const CAD_API = Object.freeze({
  primitives: Object.freeze({
    Box: "Box(x,y,z)",
    Sphere: "Sphere(r)",
    Cylinder: "Cylinder(r,h)",
    Cone: "Cone(r1,r2,h)",
    Circle: "Circle(r,wire?)",
    Polygon: "Polygon(points,wire?)",
    BSpline: "BSpline(points,closed?)",
    Text3D: "Text3D(text,size,height,font?)",
    Sketch: "new Sketch([x,y],plane?)"
  }),

  operations: Object.freeze({
    Translate: "Translate([x,y,z],shape)",
    Rotate: "Rotate([ax,ay,az],degrees,shape)",
    Scale: "Scale(factor,shape)",
    Mirror: "Mirror([x,y,z],shape)",
    Union: "Union(shapes)",
    Difference: "Difference(body,tools)",
    Intersection: "Intersection(shapes)",
    Extrude: "Extrude(face,[dx,dy,dz])",
    Revolve: "Revolve(shape,degrees)",
    Loft: "Loft(wires)",
    Pipe: "Pipe(shape,path)",
    Offset: "Offset(shape,distance)",
    FilletEdges: "FilletEdges(shape,radius,indices)",
    ChamferEdges: "ChamferEdges(shape,distance,indices)"
  }),

  methods: Object.freeze([
    "LineTo", "Fillet", "ArcTo", "BSplineTo", "End", "Face", "Wire",
    "ofType", "parallel", "max", "min", "indices"
  ]),

  measurements: Object.freeze([
    "Edges", "Volume", "SurfaceArea", "CenterOfMass"
  ]),

  math: Object.freeze([
    "PI", "abs", "min", "max", "sqrt", "sin", "cos", "tan",
    "round", "floor", "ceil", "pow"
  ]),

  keywords: Object.freeze([
    "const", "let", "var", "new", "true", "false", "null", "undefined"
  ])
});


function validateCadCode(rawCode) {
  let code = String(rawCode || "").trim();

  code = code
    .replace(/^```(?:javascript|js)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();

  if (!code || code.length > 20000) {
    throw new Error(
      "The generated CAD script was empty or too long. Please try again."
    );
  }

  // Scan the source while preserving string literals safely.
  // Comments are removed, and template strings are not supported.
  let masked = "";
  const strings = [];
  let i = 0;

  while (i < code.length) {
    const ch = code[i];
    const next = code[i + 1];

    if (ch === "`") {
      throw new Error("Template strings are not supported in CAD scripts.");
    }

    if (ch === "/" && next === "/") {
      i += 2;
      while (i < code.length && code[i] !== "\n") i++;
      masked += "\n";
      continue;
    }

    if (ch === "/" && next === "*") {
      i += 2;
      while (
        i < code.length &&
        !(code[i] === "*" && code[i + 1] === "/")
      ) i++;

      if (i >= code.length) {
        throw new Error("The generated CAD script has an unclosed comment.");
      }

      i += 2;
      masked += "  ";
      continue;
    }

    if (ch === "'" || ch === '"') {
      const quote = ch;
      let literal = "";
      i++;

      while (i < code.length && code[i] !== quote) {
        if (code[i] === "\\") {
          throw new Error(
            "Escaped characters are not supported in CAD text."
          );
        }

        if (code.charCodeAt(i) < 32) {
          throw new Error("Invalid character in a CAD text literal.");
        }

        literal += code[i++];
      }

      if (i >= code.length) {
        throw new Error("The generated CAD script has an unclosed string.");
      }

      i++;
      strings.push(literal);
      masked += " STRLITERAL ";
      continue;
    }

    masked += ch;
    i++;
  }

  // Never permit general-purpose JavaScript or browser access.
  const forbidden =
    /\b(?:window|document|globalThis|self|fetch|XMLHttpRequest|WebSocket|import|export|eval|Function|constructor|prototype|__proto__|localStorage|sessionStorage|navigator|location|process|require|Worker|setTimeout|setInterval|postMessage|cookie|indexedDB|await|async|class|while|for|do|try|catch|throw|debugger|this|delete|with|yield|return|function|new\s+Function|Proxy|Reflect|WebAssembly|SharedArrayBuffer|Atomics|arguments|super|return)\b/i;

  const forbiddenMatch = masked.match(forbidden);

if (forbiddenMatch) {
  throw new Error(
    "Unsupported JavaScript token: " +
    forbiddenMatch[0] +
    ". Please regenerate the CAD model."
  );
}

  // Only allow documented CAD operations and their supported methods.
  
  // Derive validator allowlists from the shared CAD API registry.
  const allowedFunctions = new Set([
    ...Object.keys(CAD_API.primitives),
    ...Object.keys(CAD_API.operations),
    ...CAD_API.measurements
  ]);

  const allowedMethods = new Set(CAD_API.methods);

  const allowedKeywords = new Set(CAD_API.keywords);
const allowedOutputVariables = new Set(["finalModel"]);

  const declared = new Set();

for (const match of masked.matchAll(
  /\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)/g
)) {
  const name = match[1];

  if (declared.has(name)) {
    throw new Error(
      "The generated CAD script declares '" +
      name +
      "' more than once. Regenerate the complete script with unique variable names."
    );
  }

  declared.add(name);
    }

  // Restrict property access to approved CAD methods.
  for (const match of masked.matchAll(
    /\.\s*([A-Za-z_$][\w$]*)/g
  )) {
    const property = match[1];

    if (
      !allowedMethods.has(property) &&
      !["PI", "abs", "min", "max", "sqrt", "sin", "cos", "tan",
        "round", "floor", "ceil", "pow"].includes(property)
    ) {
      throw new Error(
        "The CAD script uses an unsupported property: " + property
      );
    }
  }

  // Reject computed property access and indirect function calls.
  if (
    /\b[A-Za-z_$][\w$]*\s*\[\s*[A-Za-z_$]/.test(masked) ||
    /\]\s*\(/.test(masked) ||
    /\)\s*\(/.test(masked) ||
    /=>|\?\./.test(masked)
  ) {
    throw new Error(
      "The generated CAD script uses unsupported JavaScript syntax."
    );
  }

  // Reject object literals, function bodies, and other block syntax.
  if (/[{}]/.test(masked)) {
    throw new Error(
      "Use CAD operations and variable declarations only."
    );
  }

  const identifiers = masked.match(/[A-Za-z_$][\w$]*/g) || [];

  for (const identifier of identifiers) {
    if (
      identifier === "STRLITERAL" ||
      allowedKeywords.has(identifier) ||
      allowedFunctions.has(identifier) ||
      allowedMethods.has(identifier) ||
      declared.has(identifier)
    ) {
      continue;
    }

   // Allow the required final CAD output variable.
if (allowedOutputVariables.has(identifier)) continue;

// Math is allowed only with the approved properties listed above.
if (identifier === "Math") continue;

    throw new Error(
      "The CAD script uses an unsupported identifier: " + identifier
    );
  }

  // Permit only safe literal text for plane and selector names.
  for (const literal of strings) {
    if (
      literal.length > 500 ||
      /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(literal)
    ) {
      throw new Error("The CAD script contains invalid text.");
    }
  }

  if (
    ![...allowedFunctions].some(name =>
      new RegExp("\\b" + name + "\\s*\\(").test(masked)
    )
  ) {
    throw new Error(
      "The CAD script did not contain a supported geometry operation."
    );
  }

  // Check bracket balance.
  let depth = 0;
  const pairs = { ")": "(", "]": "[" };
  const stack = [];

  for (const ch of masked) {
    if (ch === "(" || ch === "[") {
      stack.push(ch);
      depth++;
    } else if (ch === ")" || ch === "]") {
      if (stack.pop() !== pairs[ch]) {
        throw new Error("The generated CAD script has invalid brackets.");
      }
      depth--;
    }

    if (depth < 0) {
      throw new Error("The generated CAD script has invalid brackets.");
    }
  }

  if (depth !== 0) {
    throw new Error("The generated CAD script has invalid brackets.");
  }

  return code;
}


/* ------------------------------------------------------------------ */
/* Code Studio (Pro)                                                   */
/* ------------------------------------------------------------------ */

const CODE_SYSTEM = [
  "You are Malie Code, an expert front-end web developer inside Malie AI's Code Studio.",
  "You build complete, working, responsive, accessible and visually polished websites and small web apps using plain HTML, CSS and JavaScript.",
  "Follow the output format given in the user's message exactly.",
  "Never stop mid-file: always finish every file you start, and keep the code compact enough to fit.",
  "Do not use external images or libraries that need an API key. Use CSS, inline SVG or emoji for visuals.",
  "Design quality: use a clear visual hierarchy, a consistent colour palette with CSS variables, generous spacing, a modern font stack, smooth hover states and subtle transitions.",
  "Always include the viewport meta tag, a descriptive title and meta description, semantic HTML, sufficient colour contrast, visible focus states and a mobile-first responsive layout.",
  "When the user asks for a change, edit the existing code and keep everything else working. Never remove features they did not ask to remove."
].join(" ");


/* Nemotron teams up with the main model in Code Studio: it reviews the draft and fixes real bugs. */
const CODE_REVIEW_SYSTEM = [
  "You are Nemotron, a senior code reviewer working hand in hand with Malie Code inside Malie AI's Code Studio.",
  "You receive the user's request and a draft answer written by Malie Code.",
  "Check the draft for real problems only: JavaScript errors, missing or broken event handlers, unclosed tags, broken links between files, layout that breaks on mobile, missing accessibility basics.",
  "If the draft is already correct, reply with exactly: OK",
  "Otherwise reply with the COMPLETE corrected answer in the exact same format as the draft: one or two short sentences, then every file in full, each introduced by a line `FILE: <name>` followed by a fenced code block.",
  "Keep the design, features and file names. Never write '...' or 'rest unchanged'. Never add commentary after the last code block."
].join(" ");

async function nemotronReviewCode(env, request, draft) {
  try {
    if (!env.NVIDIA_API_KEY) return draft;
    if (!draft || draft.length > 22000) return draft;
    if (!/FILE:\s*[\w\-./]+/i.test(draft) && !/<\/html>/i.test(draft)) return draft;
    const prompt = "USER REQUEST:\n" + String(request || "").slice(0, 3000) + "\n\nDRAFT FROM MALIE CODE:\n" + draft;
    const out = String(await callNvidiaAI(env, prompt, false, { system: CODE_REVIEW_SYSTEM, maxTokens: 8192, timeoutMs: 70000 })).trim();
    if (!out || /^ok[.!\s]*$/i.test(out)) return draft;
    const looksRight = /FILE:\s*[\w\-./]+/i.test(out) && /```/.test(out) && out.length > draft.length * 0.6;
    const htmlOk = !/FILE:\s*index\.html/i.test(draft) || /<\/html>/i.test(out);
    return looksRight && htmlOk ? out : draft;
  } catch (e) {
    console.warn("Nemotron review skipped:", String(e && e.message || e));
    return draft;
  }
}

const CODE_LIMITS = { maxPromptChars: 80000, maxOutputTokens: 8192, maxPublishChars: 400000, maxPublished: 50 };

async function ensurePublishedTable(env) {
  await env.DB.prepare(
    `CREATE TABLE IF NOT EXISTS published_pages (
       id TEXT PRIMARY KEY,
       user_id TEXT NOT NULL,
       title TEXT,
       html TEXT NOT NULL,
       created_at TEXT DEFAULT CURRENT_TIMESTAMP,
       updated_at TEXT DEFAULT CURRENT_TIMESTAMP
     )`
  ).run();
}

function randomPageId() {
  const bytes = crypto.getRandomValues(new Uint8Array(9));
  return Array.from(bytes, b => "abcdefghijklmnopqrstuvwxyz0123456789"[b % 36]).join("");
}

/* ------------------------------------------------------------------ */
/* Worker                                                              */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/* App icons (served by the worker, so no image files are needed)      */
/* ------------------------------------------------------------------ */

const APP_ICONS = {
  "/icons/icon-512.png": "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOmE2ZDA3MTVmLTcyZjktNDVhNS1iNjVlLWQ1Mjg2MDgyMmQxNwAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNo/dLyZPQ1IbQgSvYUq0ShJgAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6MTE4MGE5NmQtMTc2Zi00MTllLWEyZDMtMGE4YTMzY2I5NGZjbHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaJV/5VVEF6yu7Y9Oo88XFpsAAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIMGkQISeFGLvAHGNjT3Fepn8w5gUpSCwvZkLdNvWRnHFpGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2jQcfvii6xJ1PFBjQd6b1cDAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFggLAPToKfH726QRbys5mdIMfXQ2oebTy0Wae2juKRq8v1kbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTphNmQwNzE1Zi03MmY5LTQ1YTUtYjY1ZS1kNTI4NjA4MjJkMTcvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDoyZGJjNWFkZi0xOThiLTQ2Y2EtOTA4Mi03Y2Q5MGE5NjNkODVyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggwaRAhJ4UYu8AcY2NPcV6mfzDmBSlILC9mQt029ZGccWiY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIFFbOLB2KBKRXO0VgH42UQX/yN+15LIgNhWNFwC5EpEtomN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYIAr0A0Nk2kOlD9/Th97b/bllZEq5HQOwM22cRWPHHe3odGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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NQBEQM7AYfHOv+r+FxERmeUszmHgMCBHyzwA2/J/5xcKLzeY3ZnAfYRFRESk41nAGczufqHwclrqe2sAwIb2cIyxxOP/IiIiMvuFGGNtaA9vvj8qADgAZ8OjmiP/mvwnIiLSHQyuWeNjw+P88SzBxYtLxtmDtPxPRESkqzSXA9qDWLy4RHOVXxIAyAwMLMGwLXECUA+AiIhIdzCAw7BtZmBgSfKYZXj8n30xxmMjNwwQERGRWS3CGM+G7Nt83w73AESGVzQf1Pp/ERGR7uJgVK0fnvGfMbjDm6Vf4/8iIiLdxeLA4A4HMkBogYg5c0rA7s2AoPF/ERGR7mKaNX73Zs2PLECm0dgZTBFNABQREelG8URATDGu+SMTAPfDGB9NABQREelWEcb4NmQ/aAaACLdn80lNABQREelO8UTAZs2PJ/wZs7d2ABQREelqJh7oN3tDHACyxrFo+EkRERHpRgagWfOzlnnzCuB20woAERGRrtZcCeB2Y968gs3U6zsByQoAERER6V4OKGbq9Z2sDdx2GFNsPqEeABERke4U13hjijZw29koiua293hERERkJkVRNNdizcHJ7QDafUAiIiIyrSIwYM3BFii3+2hERERkRpWtgSVaASAiIpIKBhwGllhgSbuPRkRERGbUEgum3u6jEBERkZlk6hat/xcREUkbZ7K5Qj/G9KJbAYuIiHS7uNY7N2Cy+aJ6AERERFJGQwAiIiLp4yzq9hcREUkbY9t9BCIiIjLzFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgB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FABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhRQAREREUkgBQEREJIUUAERERFJIAUBERCSFFABERERSSAFAREQkhSzg2n0QIiIiMqOcBUy7j0JERERmlLE4N9B8Rz0BIiIi3S2u9c4NWDAPjXpQREREulWz1puHNAQgIiKSPsaCy7b7KERERGQmuawFHm73YYiIiMiMetg6eLg5CqA5ACIiIt3NgcHBwxYotvtoREREZEYVLZG7q3nxr10BRUREupsFB5G7y1pr17X7aERERGTmWGvX2cg3z+BcufmY5gGIiIh0p2QToHLkm2dsI5t9Eiij/QBERES6nQHKjWz2ScvatRUwj2klgIiISFdzca03j7F2bcUCdWd4fuRJERER6UIOoFnz6/HMf+eWNQcAFABERES6k8MQ13yaS/8s5pHmk5oHICIi0p0MjNR8CxB5/AnnArQXgIiISLeyOBdEHn+K3wEamcwT4JKVABoGEBER6S4OMODKcc2PA4Bl/fohYLlWAoiIiHSl5goAljdrvrWABzQcZmlzBkDUvuMTERGRaRBhwGGWAg3AsyQ3AnDc2/wgTQQUERHpLvEEwJFa7yzNK/7I4z6cC9FEQBERkW5jcS6MPO5rvh8N9wA0ensfxvEsmggoIiLSTeIJgI5nG729DyePJQHA47nnhpyJ7m72EmgegIiISHeIwOBMdDfPPTdEPPfPJd39BsBE3s3aEVBERKSrOEyzxseaGwLFmvMAoqU4FxGnAxEREZn9PJyLIi9a2nw/gtEBwASVyv0Ot5w4HWgYQEREZHaLAONwy4NK5X5a6nvrjH8L1BzchtE8ABERkS4QYQwObgNqtNT9DZb82ci7AudA+wGIiIjMdgbn4to+RmsAiADqXngXzvXTnCU4QwcoIiIiUyte5edcf90L72o+Nty7b8d8oKVSedoZlmo5oIiIyKzWXP7HUiqVp4lr/vCF/dghAAtgnfmFBgBERERmOdOs6bFRNX9sAAgBah5XE0UaBhAREZmd4u7/KOqveVzdfCxs/YCxASAeBiiXn3WG2zQMICIiMisl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",
  "/icons/icon-192.png": "iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOmVlOWIzZWUyLTRhYjEtNDNlZi04NGQxLTgzOWFlNjQ3OGE2ZgAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNoviODnr5YyFoTRT0Kvy9VLgAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6NzFlODQ5NGUtODc0OC00MjkyLTgzYWQtNzM1MGRlN2E4MWQ1bHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaJ5/WXPwwWsLhxIbXUl79iEAAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIBp0y1Edyz/qdsJ6I0L5s1dDnFJnbmXgAwrXrVtXDio5pGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2i5lbXHr1nfey1uyFOIC04FAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFggx+NLts0w7SMjMiMeeUfIiVK3lFTLQIWoGUNrp3a0m9FkbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTplZTliM2VlMi00YWIxLTQzZWYtODRkMS04MzlhZTY0NzhhNmYvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDpmODNmMTAzNS1lOTU0LTQ5MGEtYmRhNi1lZjYxYzQxZjFiNGFyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggGnTLUR3LP+p2wnojQvmzV0OcUmduZeADCtetW1cOKjmiY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIBjU7eGL17bXK72IY/uwjx7yCA6Yg2nKCgsosls0/GFJomN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYICr5jW6zz0FQZXD2+VsMf4VwG2ep5/Cz75OzXrpQinIadGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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",
  "/icons/icon-maskable-512.png": "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOjFkZjIzZDY4LWE3NDYtNGRjMi1iYTc0LTAzMzdkMTYwODViNQAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNomyButR2lk0V7DNY5TIChagAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6Nzc1NjViZGYtZWEzMC00ODI5LTg0YzUtYWY5YTQ3Y2YwMjNhbHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaO9OXTyEyO63gn/EgfnNgwwAAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIGFPnZbPBGCgY0e7Bd9m59Q29ey6ky4611cLYvJN0qjopGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2gtjS2QLOl3wWRC8FEbcexpAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFgghVQUX3X6dAJA9/KDZzbHnzXezOtsm1RH1appx+jDUZZkbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYToxZGYyM2Q2OC1hNzQ2LTRkYzItYmE3NC0wMzM3ZDE2MDg1YjUvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDpiNjczMTg3Mi04NTI2LTRhOTAtOWE5NS0wMzA0ZDczNGNhOGRyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggYU+dls8EYKBjR7sF32bn1Db17LqTLjrXVwti8k3SqOiiY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIGh2O24PS9cqJpy77jIO9v3cvB8eRDaB4oHdlmjHTo+TomN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYIM33ErV+sX5We34D+UlqjCExscSN4RuLfJuaYGMQ9jNbdGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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RZGCIJDjOEqn03KchCQp8Kvq7evTqlWrtXr1auVyeRVLRdVqdVmWlPASynZkNaKrS5ttNlqjR4/WyJEjZdueJMn3q6pUKgrDSI5DDLyd5qAvSalUSp6XGvibQH39/er5+9+1atVq9ff3q1AsqlqpKoxCua6nTCatzs5Ojeru1uabb6bu7m5lMp0D/95XqVSW7/ttGQO2batWK2n//fbV+O0maMnixUqmUoQAhhwBgNhoDvzpdEqe16UgqGnhwoV6/PEn9cTcufrrX5/TK68sVU9vr4rFokK/tt7343hJZbNZjR49WuPHj9Xuu+2mfffdR/vus7e2336CJEfVaknVauNabUJgXc39kEgklM12SYq0dOkSPfnk03r8ibl65pm/aNHLi7Vy5UoVi0XVqxU1Zl7ewHaVyWQ0YsQIbTNma+28807a+0Mf1L777aPddt1VXV3diqK6isXiwAxPe0yVW5alWq2uUaM210EfPkCLF70k27bE+I+hZiVSGXNOwLUhx3FULuZ12BFHaNbMaSoUCm13xLRm4M/I85JaunSJfvvbu3XXb+/Wk089rUKuf80b22tOAViW9Q+D99prBHzfl0J/8O+6R43Wfvvtq5NPOlHHHnuMttxyK9VqZVUqFR6kM8APAiU8T6lUVr29qzVt+nT9+te/0Zw5f9SK119b84aWI2dgP6zv1MraawR831cUrLkm3kumtMvOu+jYY6boYx8/SXvusZekSPl8bvBUTdwFfqDOrpG6/obrdf555yud7VQQ8KRADC0CIObaPQDCMJRt28pkOrXo5b/pxz++WlOn3qrXX3tVkuQmUvI8b+C8amNxn6S3XVjWiAPJsuzBc7K1el1BrSJJ2nbceJ1z9lm68MILtNX7xqhYzA2enzZRc3q6o2OEenpW6brrb9B119+ghS+8IEmy3YSSyeTg13LtX2+lGQaDsWZZCnxftUpFUqhUJqujpxyliy/+Vx180CHy6xWVyuXYB1kYhkqn03r+hRe1734fVrVak+PYRi2IxPAjAGKunQPA9311dHSoVqvr//3fVfre9/5Xq1aukO0mlEql1ll4tjGsfXRZqVQUBXVtO3acvv61r+r8889TFIUql0tynHgPPu9UEARKJpNKJJK6+ZapuvLKb2rhiy9KlqN0JiNJG301f/Pcv+/7qlVKctyEzjjjNF1++aXabvwE5fN9bbFo07IsHfDhf9K8P/1JqUyWdQAYUu0xUqDt+L6vrq5uPbfgeR1x5FH66le+ot6+fmU6uuR53jqrxjeW5qmG5vntTEeXli3/uy688DM6/oSTtHz5q+rsHNk4dWAI3/fV2dmpnp4+nXrqaTrzU2dq4Ut/U6ajS6l0evDrtbGPXAdP0UhKZzvlep5uuvFGHXDAQbr11lvV2TlS0tvP9LSyxmmtDk36wO6SItl2vGMG8UMAoOU0B/+7775bhxxyqB595BFlOrrkum7jfPEQvOhHUSTf95VMJpXp6NK9v/+9Djr4ED38yCPq6uo2IgKa++HJp/6kgw/5iG6//XalOzqVSqUHL58cCs3AyHR0acXKVTr99NP15a98Wel0RrZtx/aoufl9vPeHPjTwJwQAhhYBgJbSHHRu+PnPddLJH1N/f07pbNeQDfxv1DwSzXR0aemy5Tp6yjH63T2/a/sIaO6HWffP0pFHTtbCFxcq2zlCgR8M24Dr+74SiYTS2U599zvf1ZlnnSPHaSw0jONMQON0U6jddnu/bMdlESCGHAGAltEcdG6+5Sadd+55cr2EvERCQTD8A63v+0qnM6rW6jr1lE9o2vRpbRsBjf0wUg8+NFsnnniScvmC0tmOlnh6XfMKjmznCE295Radfe55SiSSkuJ3OsCyLNXrNY0bN04jR41SvV6P/ZoGxAsBgJYQBI1B5/4HZun88z+tRDLVctO7QRDISyRUDwKdfvonNW/en9XR0V6Xb4VhqGy2QwsWPKdTTjlNpVJFyWSqJSKsKYoi1et1ZTtH6Napv9SXvvwVZbNdLfW9siEaAVDXlltuobHbbqvQJwAwtAgADLswDJVMprV02VKdffZ58v1Ajuu25At6Y4FgUn29vfrUWecol8vJ87zYHX2uTxRFjatKSiV96sxztHLF64ML/VpRvV5XOtupH3z/B7rxphtjuUAzCEKl01ltN368pIgAwJAiANASPC+hiz/3eS1b+kpLDzpSIwIyHZ169pn5+tp/XKpUqj0u3wrDxqr0Sy//Tz391JPKdHS1/IAaRZG8RFKf/8KX9NJLLyqdzsRqX0RRKMnSuHFjJYkAwJAiADCsGtf6j9Add9yhu+68U+lsZ8sPOlJju1OZrK695lo9+NADsT8VEASBstkuzZnziK666sdKprMKYrAfwjCUl0iqZ9VKfeWSr8vzPK33tsMtrhkAwFAiADBsoiiS53nK5/t1xZXfkm3HazW3ZdkKfF+XX/HNwQfXxFVzvcVlV3xT9VpVtm3HZhhtxFiH7vrNXZoxY4ay2a7YxFjziH/stttKkqIYzV4g/uL7ioXYa94I5bbbb9dzz/5FyXS8ptKDIFAqk9VDDz6k6TNmKBvT+7kHQaBspkP3zbpPD9x/v1KZjth9HpZlKQpDffd7P1QQBDGbSg81erPRsmxXYYwCGPFHAGDYuK6rSqWka669XpbVuCY6bizLkqJQP73mOoVhqFiNOwMsqzFp/tOf/kxRGLfBsyEIAiXTGT344Gz98bE5ymQ6YhGTlmUpDAN1jxypdCYz8D0Uv68/4okAwLBoHP1nNeePf9Sfnn5aiVRaQdD6L9hvFASBvGRasx+YrWeffUaZmN3PvfFQmoyee+5Zzbr/fnnJ1l6A+VYcx1Hg13XLLbfKtp1Y7AfLshQEgbq7RyqTjufPAOKLAMCwiKJIlmXrrrvuVhj4sX7Wu+d5KpcK+t09f5Bte7EYeJrCMJTjJPS7e+5VqZAfWEQXT0EQyHY8TZs+UytWvqZkMhmLNSVhGCqVSiudySiK4hlfiCcCAENuzeK/Pt3/wGxZMTlaezONQcbSjJn3qV6vxipmHMeR71c1fcZMSVYsBsw3E0WREsmkli1doscee1zJZDoW31dRFCmR8JTJpBUF3AsAQ4cAwJCLokipVFp//etzemnhS/ISyVi8UL+ZMAzleEk988xftHjxYiWT8fh8GjdgSmrx4iWaP+8ZOV4iFtv9VmzblqJIDz74sOIQNI01AKE8z1Mmk1Ec18EgvggADLnGQidHTz71tPx6Va7rDvcmvSfNI7hcX4/mzZ8v143H1HMURXLdpObNn6/+vh4lEolYbPdbaQSMpblzn4zNbExzRiyVTCmO9zBAfBEAGHLNKc4//3newJ/E/0Vvzec0X1I8HkzT3MbmNrfD1HMURbJdTwtf+ptee+01JRLxuE2zZdty3MbLcTvsB8QDAYAhZ9u26vWKXnhxoSQpDFv/BfrtNAeZBQuelxTJslv/Rbwx0ER6bsECSfGIlrfTPJpevXqVlixZIs9LDtxut3VFUSTHtuQ48Z4JQ/wQABhSjWlnV/39/Vq+bLks222bgUeytOSVV1QuF+W0+F0NBx/8Uy5qyZJXFIfz5RvKtm2Ffl0vv7xYkq04fFqWZcfidAXaCwGAIee6jlav7lFPb6+cmJ//b4qiSJbjasWKFcrl8rF4MXccR7lcXitWrJDltEeISWum0F9ZukxSfGY2mPnHUCMAMKQaR56uenp6VCqVGvecj8kL9NtxHEf5fEH9/f1yYxAAruOov79f+XwhFsHyTq1Y8XrjNwyswHoRABhSjcHeUS6XU1CvxvoBOm9k27YqlYpyuZxsp/VPAdiOo/5cTpVKZWA/tO72vhPNr3tPT68kyaIAgPVqn1dfxEq+UJDUPiueG3c2tFSv11UqlQaebdDaLMtWqVRSvVZvPEynPcb/Qfl8Qe0SNcCm0PqvUmhL1Wp1uDdho7MsS1Hgq1avq9UX1TUXLdZrNUWh3zYhtrZKtSKJO+sBb4YAwLCI6wNn3l408LnFYdCxFISh2vUouX2/x4CNgwDAsGjhg+P3Lk6fW5y29R1qzHK08ScIvEcEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADOQO9wYAAFqXZVmy7cZ/wyiSojV/Z9uWIkWKQikMozd/J2hJBAAA4B/YjiVLUqUcKAp9NUZ+W5K11lsFA//vKJl2ZdtSEK4bCWhdBAAAYJBlWbIsqVKsSYo0ZvwI7bHPltp+55Ha8f3dGjkqpTAI5ddDLZjfo1dezumZp17X83/plSJfiVRStmMpDKiAVkcAAAAkSY5rq1r2FQaBDp48VsefvqMOPWqsRm+RlitbfmPCf/DtjzpugkJFKhRrevLhV/XbWxfq3jsXqVKsKd2RVOCHw/jZ4O0QAAAAOY6lcqGi943p1CXf2V8nnLajbMtSsVRXf39NURQ1ZgcG3j6SBv/McSwdfNRYHT5lvD514d91xefnaN4TryqZSUlRpIjJgJbEVQAAYDjbsVQuVnXQkeN012Mn6eTTd1I+V1NfX1VBEMlxLLmuLcexZA/8WvvPJKmQq2lVX1l77Lelbp99gs69eE9Vy3VFkWRZb7MBGBYEAAAYzHFsVYpVnXrerpo641iN3iKtnr5yY5B3rQ0evO2BICjk66r7of77h4fom1cfrHrNl2VZ664dREsgAADAUI5rq1ws66iTd9D/XvcR5fN11aqBXPfdDw3NGYHXe4u64MJJuuTb+6tSqsixGW5aDXsEAAxk25aq5brGTxyt//rJwSqW64qiSLb93g/VLUtyXVsr+ku68At7aMrHJqpcrMhxGHJaCXsDAAwURY1Z+ct/eKA23zyjajXYKIP/IKsRAtVaoMt+cIC2HNPZOB3AqNMy2BUAYBjHtVQtV3XimTvryCnj1dtXeU/T/m/Gti2VS7623aZLn79ib/l1XzYrAlsGAQAAhgkDyXYcnXruTqoFwSYdlB3XVr5U0+QTttNWY0eoWvG5KqBFEAAAYBDbsVSt1LTX/ltpz322ULFUl+1suhHZsqRaLdDmo9M67pTtFQY+awFaBHsBAAxiW5YUhTroiDHKJDwFQ3DLXtuy5EeRDjpiG9m2y4ODWgQBAAAGiaLGQ33eP2m0fIVDck7esi3VaoEmTBypzpFJ1eoBpwFaAAEAAIawLMn3Q2U6EhoztlN1P2zcpGeIPu6I0Um9b5sORcHQfFy8NQIAAAwSBJEynQltsVVGfj0asjv0BUGkjs6Ettwqoygcuo+LN0cAAIBhwiCS74dDPg0fho2Pi9ZAAACAYWxbsofh1ry2La4AaCHsCQAwRBQ1btHb31PV0sU5eQl7YFHg0HzcvtVVLVnUL8sZmo+Lt0YAAIBBHMdSvebr1aUFObatoRqHXddW7+qKelZWZDu2xPg/7AgAADBIY/V9qD89sUKurCE5Eg/DSMmEo+f/slqF/ooS3tCFB94cAQAABgnCUJbt6uEZS7W6vyzPG4phIFIURZp59xJx6N86CAAAMEgUSsm0q5eeW6XHZv9dHdmEAn/TDcpRJKVSrpa+ktesexbL8TzuBNgiCAAAMJBlW/rFVc82LgfchCNB4IfKJj398toFyveVlUg6TP+3CAIAAAwTBpGS6YT+eP8S/eLqZzW6K616feNfnx/4kUaMSGruE6/quh/MUyKVUBhwH4BWQQAAgIGiMFIimdD3L52rp//0mkaOSG7Um/SEYSQvYauQr+mSCx9WvRrIdiyO/lsIAQAABooiyXFtFXI1ffGc2epdXVFHhyd/I8wEBEEkx7GUzXj6xr89qgXzXlcyk1A4BE8exIYjAADAUEEQKZVJ6PlnVunMY+7VytfKGtmdku+H7/ryQN8Plcm48lxHXzx/tn5z03NKd6QVcAvglkMAAIDBgiBUOpvQM0+u0HH73qn7/7BEo0em5XmOfD9UGEZvO20fhtHgAD96ZFqLX8rpE4fdrduu/6vS2RSDf4tyh3sDsHFEUaQwDBWGrf2D1tzGdr4NaPNzDMPWfeTp2tvYrviZ2HBBECmd9bTq9bLOPfZenfnZ3XXRJXtqq6075EehKuWgMSsQStHAdfyWGjcVsh1L6YyrpOsqX6jqZ/83X9+/bK5yfdXG4M+iv5ZFALQJ13Vl2546Oztkbcpret6jMAxk27ZSqdRwb8omk8lkZNu2Ojs7ZNvOcG/OejX3QyaTGe5N2WTi8jMRRaEsy5PrDu/LcRBE8pKOFDm68ar5mvabv+nw48brhNMmauLu3RrZnZRr2bIHnuMbSfIVqpiva+FzvZrxm0W65/a/6aUFq+V6nlLZBIN/i7MSqUz7HooZwLIsBYGvUaNGaY9JkxQEQUs/ZzuKJNdx9eprr2rBggVyHLdtZgOa+2KPSZM0atRo+YE/5I9b3VDN/dDTs1rz5s9vy/0Ql58JRZLjOJo3f756enpaYl84jqVqNVDo12XZrsZuP0LbjOvQxPeP0ohRycZjfeuhFsxfrVeXFfXywj5Vy1VJrtJZd4NOG2D4EQBtofGCF9Srw70h74CtRCqt9rstqKVapaT4fF6WEqmM4rO9Gyp+PxOOl5TjuGqVfWFZku1YCoNI1UrQuIWg3nhEb0my5SVduW7jEj/u8hcfBECbsCxrWJ7v/W41z8+2I9u2W/bc/xu1836I28/EcK8DeCuWbclunPQfnEyJJClqHOmHUdQq3YJ3gAAAAMBA8cljAACw0RAAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAMRAAAAGAgAgAAAAMRAAAAGIgAAADAQAQAAAAGIgAAADAQAQAAgIEIAAAADEQAAABgIAIAAAADEQAAABiIAAAAwEAEAAAABiIAAAAwEAEAAICBCAAAAAxEAAAAYCACAAAAAxEAAAAYiAAAAMBABAAAAAYiAAAAMBABAACAgQgAAAAM9P8BjHcmJP8EwAkAAAAASUVORK5CYII=",
  "/icons/apple-touch-icon.png": "iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAYAAAA9zQYyAAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOjllOGI5OWQ5LTczNmEtNGExNy1iZmIyLWJhMmRmYWRkYWZhZgAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNo8Ydgxz2vnjvGK/WLhF/J7AAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6ZGE5Mzg0MzQtNmZlOS00YWI2LThlMzgtMDhiMjY5ZGUzMmFhbHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaBVCZ+816hgf9AHQq7O7NbEAAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYIJGiQfDtyPq2fEjglJvqRs5A/fEwDTvwzvaXy2MvlcvPpGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2iWcRCdNKt7SQ34ZVz7zcoHAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFgg5a9LFC5HBXQpbkVtUL/zqMv1iB8Hdd6wEuazxDvfNFlkbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTo5ZThiOTlkOS03MzZhLTRhMTctYmZiMi1iYTJkZmFkZGFmYWYvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDo1OTUxZDViNi1mNDQwLTQyMjEtOTVhYS03NjYwNGM4ZTViZWFyY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggkaJB8O3I+rZ8SOCUm+pGzkD98TANO/DO9pfLYy+Vy8+iY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIJDNMSjD9WbmXIV2saJxiQ5EuW/DgOJP43ds/H/StB31omN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYIP60fdr42BRO772iULdRbqnk1NA4cqYl2v58Fqg2q+F/dGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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",
  "/icons/favicon-32.png": "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAWfmNhQlgAABZ+anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFlhqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOjA5OWEwNTMwLWIxZTAtNGU5Yy04ZTUxLTdhMDgwYzM3NDBjMQAAAAOTanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuGp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNoiJpSiQKbm39uibXxLZQItgAAAGxjYm9yo2lkYzpmb3JtYXRpaW1hZ2UvcG5namluc3RhbmNlSUR4LHhtcDppaWQ6ZDg2YTg4YmUtYzYxNy00YTY5LTlkNTAtZjBhYjA1OTcwNTgybHJlbGF0aW9uc2hpcGhwYXJlbnRPZgAAAeJqdW1iAAAAQWp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuYWN0aW9ucy52MgAAAAAYYzJzaBg2EoW8pzDrd4kmyWhuQDQAAAGZY2JvcqJnYWN0aW9uc4KiZmFjdGlvbmtjMnBhLm9wZW5lZGpwYXJhbWV0ZXJzoWtpbmdyZWRpZW50c4GiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYICjfiYUvKw25y9PKph9B1aKvCAZGvhj+wDeDwgu8+tvGpGZhY3Rpb254HWNvbS5hbnRocm9waWMuY2xhdWRlLnByb3ZpZGVkanBhcmFtZXRlcnOheB9jb20uYW50aHJvcGljLm9yaWdpbi1jb25maWRlbmNlZ3Vua25vd25rZGVzY3JpcHRpb254ZkNsYXVkZSBwcm92aWRlZCB0aGlzIGZpbGUgYXQgdGhlIHJlcXVlc3Qgb2YgYSB1c2VyIGFuZCBtYXkgaGF2ZSBjcmVhdGVkIG9yIG1vZGlmaWVkIHRoZSBmaWxlIGNvbnRlbnRzLm1zb2Z0d2FyZUFnZW50oWRuYW1lZkNsYXVkZXJhbGxBY3Rpb25zSW5jbHVkZWT1AAAAyGp1bWIAAABAanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5oYXNoLmRhdGEAAAAAGGMyc2jYmzlTh4DYzss+W+DDFs1OAAAAgGNib3KlY2FsZ2ZzaGEyNTZjcGFkTQAAAAAAAAAAAAAAAABkaGFzaFggWCkgfV/dxPSSUzcIZtg300ghz8ObNozof9+PY/y6uz5kbmFtZW5qdW1iZiBtYW5pZmVzdGpleGNsdXNpb25zgaJlc3RhcnQYIWZsZW5ndGgZFooAAAI+anVtYgAAACdqdW1kYzJjbAARABCAAACqADibcQNjMnBhLmNsYWltLnYyAAAAAg9jYm9ypWNhbGdmc2hhMjU2aXNpZ25hdHVyZXhNc2VsZiNqdW1iZj0vYzJwYS91cm46YzJwYTowOTlhMDUzMC1iMWUwLTRlOWMtOGU1MS03YTA4MGMzNzQwYzEvYzJwYS5zaWduYXR1cmVqaW5zdGFuY2VJRHgseG1wOmlpZDoxMTVjYzRjNC1lYjc3LTQ4ZGEtODA1YS1kMDA5NjdiNTA3MzByY3JlYXRlZF9hc3NlcnRpb25zg6JjdXJseC1zZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmluZ3JlZGllbnQudjNkaGFzaFggKN+JhS8rDbnL08qmH0HVoq8IBka+GP7AN4PCC7z628aiY3VybHgqc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5hY3Rpb25zLnYyZGhhc2hYIE0axZsaJ/2vi+/TjPKzQ4G2MLhXsiMxWtguDrc5erH4omN1cmx4KXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaGFzaC5kYXRhZGhhc2hYIAO1deMYsBae4WRQ6gsWonhK08y+t31cebRNINfZk9O4dGNsYWltX2dlbmVyYXRvcl9pbmZvo2RuYW1lb0FudGhyb3BpYyBGaWxlc2d2ZXJzaW9uZTEuMC4wa3NwZWNWZXJzaW9uZTIuNC4wAAAQOGp1bWIAAAAoanVtZGMyY3MAEQAQgAAAqgA4m3EDYzJwYS5zaWduYXR1cmUAAAAQCGNib3LShFkCEqIBJhghWQIKMIICBjCCAY2gAwIBAgIUQOWgCu7COdC+uIP6BkIFPWdVEwAwCgYIKoZIzj0EAwMwSTEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxLjAsBgNVBAMTJUFudGhyb3BpYyBDb250ZW50IENyZWRlbnRpYWxzIFJvb3QgQ0EwHhcNMjYwODA3MTg0MzU2WhcNMjgwODA2MTk0MzU2WjBEMRcwFQYDVQQKEw5BbnRocm9waWMsIFBCQzEpMCcGA1UEAxMgQW50aHJvcGljIENsYXVkZSBDb250ZW50IFNpZ25pbmcwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAASYegpry1AYBRTVNL1CpTlbROnY3dey+UrsF9C3phYrATN3ZHf93Mo8RQN0KOUuOn19P4oWNFWe5n2/She9N7eTo1gwVjAOBgNVHQ8BAf8EBAMCB4AwFQYDVR0lBA4wDAYKKwYBBAGD6F4CATAMBgNVHRMBAf8EAjAAMB8GA1UdIwQYMBaAFM5R4gSBTmRbI/jjxM+aPpzB11zCMAoGCCqGSM49BAMDA2cAMGQCMDFzHRSeAXrSy1WOzkbhPZ6Km2wGTmZ/2gK18k8BQGXyqz88Rdrz6CTX9flAnYNVxgIwcF9c3fVhqmJKpi+UhasNUMko69cyX6STPfta3Q8EjyzDjzoyrol46FP6VFHhvUcJoWNwYWRZDZ4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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"
};


/* ------------------------------------------------------------------ */
/* Security + plugin helpers                                           */
/* ------------------------------------------------------------------ */

function safeMessage(error, fallback = "The request failed. Please try again.") {
  const st = error && error.status;
  return st && st !== 500 && error.message ? String(error.message).slice(0, 300) : fallback;
}

function isSafePublicUrl(raw) {
  let u;
  try { u = new URL(raw); } catch { return null; }
  const h = u.hostname.toLowerCase();
  if (u.protocol !== "https:" || u.username || u.password) return null;
  if (u.port && u.port !== "443") return null;
  if (
    h === "localhost" || !h.includes(".") ||
    /^\[/.test(h) ||
    /^(0\.|127\.|10\.|192\.168\.|169\.254\.|100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\.|172\.(1[6-9]|2\d|3[01])\.)/.test(h) ||
    /^\d+\.\d+\.\d+\.\d+$/.test(h) && /^(0|127|10|169\.254|192\.168|172\.(1[6-9]|2\d|3[01]))/.test(h) ||
    /\.(local|internal|localhost|lan|home|corp|intranet)$/.test(h)
  ) return null;
  return u;
}

const _rl = new Map();
function rateLimit(key, max, windowMs) {
  const now = Date.now();
  const hits = (_rl.get(key) || []).filter(t => now - t < windowMs);
  if (hits.length >= max) {
    _rl.set(key, hits);
    const e = new Error("Too many requests. Please slow down and try again shortly.");
    e.status = 429;
    throw e;
  }
  hits.push(now);
  _rl.set(key, hits);
  if (_rl.size > 5000) _rl.clear();
}

async function getPluginSecret(env, userId, plugin) {
  if (!env.PLUGIN_SECRET_KEY) { const e = new Error("Plugins are not configured on the server."); e.status = 503; throw e; }
  const row = await env.DB.prepare(
    "SELECT secret, meta FROM user_plugins WHERE user_id = ? AND plugin = ?"
  ).bind(userId, plugin).first().catch(() => null);
  if (!row) { const e = new Error("Connect this plugin first."); e.status = 400; throw e; }
  const packed = Uint8Array.from(atob(row.secret), c => c.charCodeAt(0));
  const keyBytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(env.PLUGIN_SECRET_KEY));
  const aesKey = await crypto.subtle.importKey("raw", keyBytes, "AES-GCM", false, ["decrypt"]);
  const pt = await crypto.subtle.decrypt({ name: "AES-GCM", iv: packed.slice(0, 12) }, aesKey, packed.slice(12));
  let meta = {};
  try { meta = JSON.parse(row.meta || "{}"); } catch {}
  return { secret: new TextDecoder().decode(pt), meta };
}

async function omniFetch(env, userId, path, init) {
  const { secret, meta } = await getPluginSecret(env, userId, "omniroute");
  const base = isSafePublicUrl(meta.baseUrl || "");
  if (!base) { const e = new Error("Your OmniRoute address is not allowed. Reconnect it."); e.status = 400; throw e; }
  const r = await fetch(base.href.replace(/\/+$/, "") + path, {
    ...init,
    redirect: "error",
    headers: { "Authorization": `Bearer ${secret}`, "Content-Type": "application/json" },
    signal: AbortSignal.timeout(60000)
  }).catch(() => null);
  if (!r || !r.ok) { const e = new Error("OmniRoute did not respond successfully."); e.status = 400; throw e; }
  return r.json();
}

async function callOmniRoute(env, userId, model, prompt) {
  const data = await omniFetch(env, userId, "/chat/completions", {
    method: "POST",
    body: JSON.stringify({ model, messages: [{ role: "user", content: prompt }], stream: false })
  });
  return String(data?.choices?.[0]?.message?.content || "").trim() || "No reply came back from the model.";
}

async function githubApi(token, path, init = {}) {
  const r = await fetch("https://api.github.com" + path, {
    ...init,
    headers: {
      "Authorization": `Bearer ${token}`,
      "Accept": "application/vnd.github+json",
      "User-Agent": "Malie-AI",
      "X-GitHub-Api-Version": "2022-11-28",
      ...(init.body ? { "Content-Type": "application/json" } : {})
    },
    signal: AbortSignal.timeout(20000)
  }).catch(() => null);
  if (!r) { const e = new Error("Could not reach GitHub."); e.status = 503; throw e; }
  if (r.status === 401) { const e = new Error("GitHub rejected your token. Reconnect GitHub."); e.status = 400; throw e; }
  if (r.status === 403 || r.status === 404) { const e = new Error("GitHub denied this. Check the repo name and token permissions."); e.status = 400; throw e; }
  if (!r.ok) { const e = new Error("GitHub returned an error."); e.status = 400; throw e; }
  return r.json();
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders });
    }

    const url = new URL(request.url);
    const path = url.pathname;

    if (request.method === "POST") {
      const allowed = String(env.ALLOWED_ORIGINS || "").split(",").map(x => x.trim()).filter(Boolean);
      const origin = request.headers.get("Origin");
      if (allowed.length && origin && !allowed.includes(origin)) {
        return respond({ error: "Origin not allowed." }, 403);
      }
      if (Number(request.headers.get("Content-Length") || 0) > 12 * 1024 * 1024) {
        return respond({ error: "Request is too large." }, 413);
      }
    }

    if (
      PAUSED_PRO_FEATURES &&
      request.method === "POST" &&
      (path === "/api/drawing" ||
        path === "/api/cad/generate" ||
        path.startsWith("/api/video/") ||
        path.startsWith("/api/advert/"))
    ) {
      return respond({ error: "This feature is not available right now." }, 503);
    }


    /* ---------------- Published pages (public, sandboxed) ---------------- */
    if (request.method === "GET" && /^\/p\/[a-z0-9]{6,20}$/.test(path)) {
      if (!env.DB) return new Response("Not available.", { status: 503 });
      try {
        await ensurePublishedTable(env);
        const page = await env.DB.prepare(
          "SELECT html FROM published_pages WHERE id = ?"
        ).bind(path.slice(3)).first();

        if (!page) return new Response("This page does not exist.", { status: 404 });

        return new Response(page.html, {
          headers: {
            "Content-Type": "text/html; charset=utf-8",
            // Runs the page in an isolated origin so it can never touch Malie AI.
            "Content-Security-Policy": "sandbox allow-scripts allow-forms allow-modals allow-popups allow-popups-to-escape-sandbox",
            "X-Content-Type-Options": "nosniff",
            "Referrer-Policy": "no-referrer",
            "Cache-Control": "public, max-age=60"
          }
        });
      } catch (error) {
        console.error("Published page error:", String(error?.message || error));
        return new Response("Could not load this page.", { status: 500 });
      }
    }

    /* ---------------- Code Studio (Pro only) ---------------- */
    if (
      request.method === "POST" &&
      ["/api/code/access", "/api/code/generate", "/api/code/publish", "/api/code/unpublish"].includes(path)
    ) {
      try {
        
const userId = await authenticateFirebase(request, env);
const plan = await getActivePlan(env, userId);

if (plan !== "pro") {
  return respond({
    error: "Code Studio is available on the Pro plan only."
  }, 403);
}

        if (!env.DB || !env.AI) {
          return respond({ error: "Code Studio is temporarily unavailable." }, 503);
        }

        if (path === "/api/code/access") {
          return respond({ ok: true, plan: "pro" });
        }

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        /* ----- Generate ----- */
        if (path === "/api/code/generate") {
          const message = String(body.message || "").trim();

          if (!message) return respond({ error: "Please describe what to build." }, 400);
          if (message.length > CODE_LIMITS.maxPromptChars) {
            return respond({ error: "Your project is too large to send. Remove some code and try again." }, 400);
          }

          const usage = await reserveDailyUsage(env, userId, plan);
          let reply;

          try {
            reply = await callAI(env, message, false, null, {
              system: CODE_SYSTEM,
              maxTokens: CODE_LIMITS.maxOutputTokens
            });
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          if (body.team !== false) {
            reply = await nemotronReviewCode(env, body.request, typeof reply === "string" ? reply : JSON.stringify(reply));
          }

          return respond({ reply, usage });
        }

        /* ----- Publish ----- */
        if (path === "/api/code/publish") {
          const html = String(body.html || "");
          const title = String(body.title || "Untitled").slice(0, 80);

          if (!html.trim()) return respond({ error: "There is no page to publish." }, 400);
          if (html.length > CODE_LIMITS.maxPublishChars) {
            return respond({ error: "This page is too large to publish." }, 413);
          }

          await ensurePublishedTable(env);

          // Re-publishing the same page updates its link instead of making a new one.
          const existingId = String(body.id || "");
          if (/^[a-z0-9]{6,20}$/.test(existingId)) {
            const updated = await env.DB.prepare(
              `UPDATE published_pages
               SET html = ?, title = ?, updated_at = CURRENT_TIMESTAMP
               WHERE id = ? AND user_id = ?`
            ).bind(html, title, existingId, userId).run();

            if (updated?.meta?.changes > 0) {
              return respond({ ok: true, id: existingId, url: `${url.origin}/p/${existingId}` });
            }
          }

          const count = await env.DB.prepare(
            "SELECT COUNT(*) AS count FROM published_pages WHERE user_id = ?"
          ).bind(userId).first();

          if ((count?.count || 0) >= CODE_LIMITS.maxPublished) {
            return respond({
              error: `You can keep up to ${CODE_LIMITS.maxPublished} published pages. Unpublish one first.`
            }, 429);
          }

          const id = randomPageId();
          await env.DB.prepare(
            "INSERT INTO published_pages (id, user_id, title, html) VALUES (?, ?, ?, ?)"
          ).bind(id, userId, title, html).run();

          return respond({ ok: true, id, url: `${url.origin}/p/${id}` });
        }

        /* ----- Unpublish ----- */
        if (path === "/api/code/unpublish") {
          const id = String(body.id || "");
          if (!/^[a-z0-9]{6,20}$/.test(id)) return respond({ error: "Invalid page." }, 400);

          await ensurePublishedTable(env);
          await env.DB.prepare(
            "DELETE FROM published_pages WHERE id = ? AND user_id = ?"
          ).bind(id, userId).run();

          return respond({ ok: true });
        }
      } catch (error) {
        console.error("Code Studio error:", String(error?.message || error));
        return respond(
          { error: error.message || "Code Studio could not process your request." },
          getErrorStatus(error)
        );
      }
    }

    /* ---------------- Health ---------------- */
    if (request.method === "GET" && path === "/api/health") {
      return respond({
        ok: true,
        service: "Malie AI",
        provider: "Cloudflare Workers AI",
        endpoints: [
          "/api/health",
          "/api/chat",
          "/api/write",
          "/api/quiz",
          "/api/excel",
          "/api/pptx",
          "/api/subscription/status",
          "/api/payment/request",
          "/api/payment/submit-reference",
          "/api/drawing",
          "/api/drawing",
          "/api/cad/generate",
          "/api/video/access",
          "/api/video/generate",
          "/api/code/access",
          "/api/code/generate",
          "/api/code/publish",
          "/api/code/unpublish"
        ]
      });
    }

/* ---------------- Knowledge Studio: Premium + Pro ---------------- */
if (path.startsWith("/api/knowledge/")) {
  try {
    const userId = await authenticateFirebase(request, env);
    const plan = await getActivePlan(env, userId);

    if (plan !== "premium" && plan !== "pro") {
      return respond({
        error: "Knowledge Studio is available on Premium and Pro plans. Upgrade to access it."
      }, 403);
    }

    if (!env.DB || !env.AI) {
      return respond({
        error: "Knowledge Studio is temporarily unavailable."
      }, 503);
    }

    const limits = plan === "pro"
      ? { documents: 50, chars: 100000 }
      : { documents: 10, chars: 30000 };

    const readBody = async () => {
      try {
        return await request.json();
      } catch {
        const error = new Error("Invalid JSON request.");
        error.status = 400;
        throw error;
      }
    };

    if (request.method === "POST" &&
        path === "/api/knowledge/upload") {
      const body = await readBody();
      const name = String(body.name || "").trim();
      const content = String(body.content || "").trim();
      const type = String(body.type || "").toLowerCase();

      if (!name || name.length > 180) {
        return respond({ error: "Provide a valid document name." }, 400);
      }

      if (!["pdf", "txt"].includes(type)) {
        return respond({ error: "Only PDF and TXT documents are supported." }, 400);
      }

      if (!content || content.length > limits.chars) {
        return respond({
          error: `Document text must be between 1 and ${limits.chars} characters on your ${plan} plan.`
        }, 400);
      }

      const count = await env.DB.prepare(
        "SELECT COUNT(*) AS count FROM knowledge_documents WHERE user_id = ?"
      ).bind(userId).first();

      if ((count?.count || 0) >= limits.documents) {
        return respond({
          error: `Your ${plan} plan allows up to ${limits.documents} saved documents.`
        }, 429);
      }

      const id = crypto.randomUUID();

      await env.DB.prepare(`
        INSERT INTO knowledge_documents
          (id, user_id, name, file_type, content)
        VALUES (?, ?, ?, ?, ?)
      `).bind(id, userId, name, type, content).run();

      return respond({
        ok: true,
        document: { id, name, type, characters: content.length }
      });
    }

  if (request.method === "POST" &&
        path === "/api/knowledge/list") {
      const result = await env.DB.prepare(`
        SELECT id, name, file_type, created_at,
               LENGTH(content) AS characters
        FROM knowledge_documents
        WHERE user_id = ?
        ORDER BY created_at DESC
      `).bind(userId).all();

      return respond({ ok: true, documents: result.results || [] });
    }

    if (request.method === "POST" &&
        path === "/api/knowledge/delete") {
      const body = await readBody();
      const id = String(body.id || "");

      if (!id) return respond({ error: "Document ID is required." }, 400);

      await env.DB.prepare(`
        DELETE FROM knowledge_documents
        WHERE id = ? AND user_id = ?
      `).bind(id, userId).run();

      return respond({ ok: true });
    }

    if (request.method === "POST" &&
        path === "/api/knowledge/ask") {
      const body = await readBody();
      const question = String(body.question || "").trim();
      const ids = Array.isArray(body.documentIds)
        ? body.documentIds.slice(0, 2).map(String)
        : [];

      if (!question || question.length > 4000) {
        return respond({ error: "Enter a question of up to 4,000 characters." }, 400);
      }

      if (!ids.length) {
        return respond({ error: "Select at least one document." }, 400);
      }

      const placeholders = ids.map(() => "?").join(",");

      const docs = await env.DB.prepare(`
        SELECT id, name, content
        FROM knowledge_documents
        WHERE user_id = ? AND id IN (${placeholders})
      `).bind(userId, ...ids).all();

      const documents = docs.results || [];

      if (!documents.length || documents.length !== new Set(ids).size) {
        return respond({ error: "One or more selected documents were not found." }, 404);
      }

      // Bound the context so requests cannot send arbitrarily large documents.
      const context = documents.map(doc =>
        `DOCUMENT: ${doc.name}\nCONTENT:\n${doc.content.slice(0, 25000)}`
      ).join("\n\n---\n\n").slice(0, plan === "pro" ? 50000 : 30000);

      const usage = await reserveDailyUsage(env, userId, plan);

      try {
        const prompt = `
You are Malie AI Knowledge Studio, a document-grounded study assistant.

Answer using the supplied document content only.
Do not invent facts, quotations, page numbers, or citations.
If the answer is not supported by the supplied content, say so.
When possible, quote a short relevant passage and identify its document name.
The text may come from PDF extraction, so page numbers are not available unless explicitly included.

DOCUMENTS:
${context}

USER QUESTION:
${question}

Give a clear, useful answer and identify which document supports it.
`;

        const answer = await callAI(env, prompt, false);

        if (!answer || !String(answer).trim()) {
          throw new Error("The AI returned an empty answer.");
        }

        return respond({ ok: true, answer, usage });
      } catch (error) {
        await refundDailyUsage(env, userId, usage.date);
        throw error;
      }
    }

    return respond({ error: "Knowledge Studio endpoint not found." }, 404);
  } catch (error) {
    console.error("Knowledge Studio error:", error);
    return respond({
      error: error?.status
        ? error.message
        : "Knowledge Studio could not complete the request."
    }, getErrorStatus(error));
  }
}

    /* ---------------- Subscription status ---------------- */
    if (request.method === "POST" && path === "/api/subscription/status") {
      try {
        const userId = await authenticateFirebase(request, env);

        if (!env.DB) {
          return respond(
            { error: "Subscription service is temporarily unavailable." },
            503
          );
        }

        const subscription = await env.DB.prepare(
          `SELECT plan, status, current_period_end
           FROM subscriptions
           WHERE user_id = ?
           LIMIT 1`
        ).bind(userId).first();

        const expiry = subscription?.current_period_end
          ? Date.parse(subscription.current_period_end)
          : null;

        const active =
          String(subscription?.status || "").toLowerCase() === "active" &&
          expiry !== null &&
          Number.isFinite(expiry) &&
          expiry > Date.now();

        const latest = await env.DB.prepare(
          `SELECT id, plan, amount, billing_period, status, payment_reference
           FROM payment_requests
           WHERE user_id = ?
           ORDER BY id DESC
           LIMIT 1`
        ).bind(userId).first();

        return respond({
          ok: true,
          subscription: {
            active,
            plan: active
              ? String(subscription.plan || "").toLowerCase()
              : "free",
            expiresAt: active ? subscription.current_period_end : null
          },
          latestRequest: latest
            ? {
                requestId: latest.id,
                plan: latest.plan,
                billingPeriod: latest.billing_period,
                amount: latest.amount,
                status: latest.status,
                hasReference: Boolean(latest.payment_reference)
              }
            : null
        });
      } catch (error) {
        console.error("Subscription status error:", error.message);

        if (error.status === 401 || error.status === 503) {
          return respond({ error: error.message }, error.status);
        }

        return respond({ error: "Could not load your plan." }, 500);
      }
    }

    /* ---------------- Create payment request ---------------- */
    if (request.method === "POST" && path === "/api/payment/request") {
      try {
        const userId = await authenticateFirebase(request, env);

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const plan = String(body.plan || "").toLowerCase();
        const billingPeriod = String(
          body.billingPeriod || "monthly"
        ).toLowerCase();

        const amount = PRICES[plan]?.[billingPeriod];

        if (!amount) {
          return respond({ error: "Invalid plan or billing period." }, 400);
        }

        if (!env.DB) {
          return respond(
            { error: "Payment service is temporarily unavailable." },
            503
          );
        }

        const created = await env.DB.prepare(
          `INSERT INTO payment_requests
             (user_id, plan, amount, billing_period, status)
           VALUES (?, ?, ?, ?, 'pending')
           RETURNING id`
        ).bind(userId, plan, amount, billingPeriod).first();

        if (!created?.id) {
          throw new Error("Payment request could not be created.");
        }

        return respond({
          ok: true,
          requestId: created.id,
          plan,
          billingPeriod,
          amount,
          currency: "NGN",
          paymentDetails: {
            bankName: "FirstBank",
            accountName: env.FIRSTBANK_ACCOUNT_NAME || null,
            accountNumber: env.FIRSTBANK_ACCOUNT_NUMBER || null
          },
          message:
            "Payment request created. Make the transfer, then submit your transfer reference."
        });
      } catch (error) {
        console.error("Payment request error:", error.message);

        if (error.status === 401 || error.status === 503) {
          return respond({ error: error.message }, error.status);
        }

        return respond(
          { error: "Could not create your payment request." },
          500
        );
      }
    }

    /* ---------------- Submit transfer reference ---------------- */
    if (request.method === "POST" && path === "/api/payment/submit-reference") {
      try {
        const userId = await authenticateFirebase(request, env);

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const requestId = Number(body.requestId);
        const transferReference = String(body.transferReference || "").trim();

        if (
          !Number.isSafeInteger(requestId) ||
          requestId < 1 ||
          transferReference.length < 4 ||
          transferReference.length > 120
        ) {
          return respond(
            { error: "Please provide a valid request ID and transfer reference." },
            400
          );
        }

        if (!env.DB) {
          return respond(
            { error: "Payment service is temporarily unavailable." },
            503
          );
        }

        const result = await env.DB.prepare(
          `UPDATE payment_requests
           SET payment_reference = ?
           WHERE id = ?
             AND user_id = ?
             AND status = 'pending'`
        ).bind(transferReference, requestId, userId).run();

        if (!result.success) {
          throw new Error("Reference could not be saved.");
        }

        if (result.meta?.changes !== 1) {
          return respond(
            {
              error:
                "Payment request not found, already processed, or reference already used."
            },
            409
          );
        }

        return respond({
          ok: true,
          message:
            "Transfer reference submitted. Your payment must be verified before your subscription is activated."
        });
      } catch (error) {
        console.error("Payment reference error:", error.message);

        if (error.status === 401 || error.status === 503) {
          return respond({ error: error.message }, error.status);
        }

        if (String(error.message || "").includes("UNIQUE constraint failed")) {
          return respond(
            { error: "That transfer reference has already been submitted." },
            409
          );
        }

        return respond(
          { error: "Could not submit your transfer reference." },
          500
        );
      }
    }

    /* ---------------- Admin: pending payments ---------------- */
    if (request.method === "POST" && path === "/api/admin/pending-payments") {
      try {
        const adminUid = env.MALIE_ADMIN_UID;

        if (!adminUid) {
          return respond({ error: "Admin access is not configured." }, 500);
        }

        const authenticatedUid = await authenticateFirebase(request, env);

        if (authenticatedUid !== adminUid) {
          return respond({ error: "Unauthorized." }, 403);
        }

        if (!env.DB) {
          return respond({ error: "Payment service unavailable." }, 503);
        }

        const result = await env.DB.prepare(
          `SELECT id, user_id, plan, amount, billing_period,
                  payment_reference, status, created_at
           FROM payment_requests
           WHERE status = 'pending'
           ORDER BY created_at ASC
           LIMIT 100`
        ).all();

        return respond({
          ok: true,
          payments: (result.results || []).map(payment => ({
            requestId: payment.id,
            userId: payment.user_id,
            plan: payment.plan,
            amount: payment.amount,
            billingPeriod: payment.billing_period,
            hasReference: Boolean(payment.payment_reference),
            transferReference: payment.payment_reference,
            status: payment.status,
            createdAt: payment.created_at
          }))
        });
      } catch (error) {
        console.error("Admin pending-payments error:", error.message);

        return respond(
          { error: error.message || "Could not load pending payments." },
          error.status || 500
        );
      }
    }

    /* ---------------- Admin: approve payment ---------------- */
    if (request.method === "POST" && path === "/api/admin/approve-payment") {
      try {
        const adminUid = env.MALIE_ADMIN_UID;

        if (!adminUid) {
          return respond({ error: "Admin access is not configured." }, 500);
        }

        let authenticatedUid;

        try {
          authenticatedUid = await authenticateFirebase(request, env);
        } catch (error) {
          return respond(
            { error: error.message || "Unauthorized." },
            error.status || 401
          );
        }

        if (authenticatedUid !== adminUid) {
          return respond({ error: "Unauthorized." }, 403);
        }

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const requestId = Number(body.requestId);

        if (!Number.isSafeInteger(requestId) || requestId < 1) {
          return respond({ error: "Invalid payment request ID." }, 400);
        }

        if (!env.DB) {
          return respond({ error: "Payment service unavailable." }, 503);
        }

        const payment = await env.DB.prepare(
          `SELECT id, user_id, plan, amount, billing_period,
                  payment_reference, status
           FROM payment_requests
           WHERE id = ?`
        ).bind(requestId).first();

        if (!payment) {
          return respond({ error: "Payment request not found." }, 404);
        }

        if (payment.status !== "pending") {
          return respond(
            { error: "This payment request has already been processed." },
            409
          );
        }

        if (!payment.payment_reference) {
          return respond(
            { error: "The customer has not submitted a transfer reference." },
            400
          );
        }

        if (
          !PRICES[payment.plan] ||
          PRICES[payment.plan][payment.billing_period] !== payment.amount
        ) {
          return respond(
            { error: "Payment amount or plan details do not match." },
            400
          );
        }

        // Verify the actual bank transfer before calling this route.
        const now = new Date();
        const confirmedAt = now.toISOString();
        const endDate = new Date(now);

        if (payment.billing_period === "annual") {
          endDate.setUTCFullYear(endDate.getUTCFullYear() + 1);
        } else {
          endDate.setUTCMonth(endDate.getUTCMonth() + 1);
        }

        const results = await env.DB.batch([
          env.DB.prepare(
            `UPDATE payment_requests
             SET status = 'confirmed', confirmed_at = ?
             WHERE id = ?
               AND status = 'pending'
               AND payment_reference IS NOT NULL`
          ).bind(confirmedAt, requestId),

          env.DB.prepare(
            `INSERT INTO subscriptions
               (user_id, plan, status, current_period_end,
                payment_reference, updated_at)
             SELECT user_id, plan, 'active', ?, payment_reference, ?
             FROM payment_requests
             WHERE id = ?
               AND status = 'confirmed'
               AND confirmed_at = ?
             ON CONFLICT(user_id) DO UPDATE SET
               plan = excluded.plan,
               status = 'active',
               current_period_end = excluded.current_period_end,
               payment_reference = excluded.payment_reference,
               updated_at = excluded.updated_at`
          ).bind(endDate.toISOString(), confirmedAt, requestId, confirmedAt)
        ]);

        if (!results[0]?.meta?.changes) {
          return respond(
            { error: "Payment request could not be confirmed. Check its status." },
            409
          );
        }

        return respond({
          ok: true,
          message: "Payment confirmed and subscription activated.",
          requestId,
          plan: payment.plan,
          billingPeriod: payment.billing_period,
          expiresAt: endDate.toISOString()
        });
      } catch (error) {
        console.error("Admin payment approval error:", error.message);
        return respond({ error: "Could not approve this payment." }, 500);
      }
    }

    /* ---------------- GitHub tools + OmniRoute models ---------------- */
    if (request.method === "POST" && ["/api/github/action", "/api/omniroute/models"].includes(path)) {
      let body;
      try { body = await request.json(); } catch { return respond({ error: "Please send valid JSON." }, 400); }
      try {
        const userId = await authenticateFirebase(request, env);
        rateLimit("plug:" + userId, 40, 60000);

        if (path === "/api/omniroute/models") {
          const data = await omniFetch(env, userId, "/models", { method: "GET" });
          const models = (Array.isArray(data?.data) ? data.data : [])
            .map(m => String(m?.id || "").slice(0, 120)).filter(Boolean).slice(0, 300);
          return respond({ models });
        }

        const { secret } = await getPluginSecret(env, userId, "github");
        const action = String(body.action || "");
        const repo = String(body.repo || "").trim();
        const repoOk = /^[A-Za-z0-9_.-]{1,100}\/[A-Za-z0-9_.-]{1,100}$/.test(repo);
        const slim = (x) => x;

        if (action === "repos") {
          const d = await githubApi(secret, "/user/repos?per_page=50&sort=updated");
          return respond({ items: d.map(r => ({ name: r.full_name, private: !!r.private, desc: String(r.description || "").slice(0, 200), url: r.html_url, updated: r.updated_at })) });
        }
        if (!repoOk) return respond({ error: "Enter the repo as owner/name." }, 400);

        if (action === "issues") {
          const d = await githubApi(secret, `/repos/${repo}/issues?state=open&per_page=30`);
          return respond({ items: d.filter(i => !i.pull_request).map(i => ({ number: i.number, title: String(i.title).slice(0, 200), user: i.user?.login, url: i.html_url })) });
        }
        if (action === "prs") {
          const d = await githubApi(secret, `/repos/${repo}/pulls?state=open&per_page=30`);
          return respond({ items: d.map(i => ({ number: i.number, title: String(i.title).slice(0, 200), user: i.user?.login, url: i.html_url })) });
        }
        if (action === "branches") {
          const d = await githubApi(secret, `/repos/${repo}/branches?per_page=50`);
          return respond({ items: d.map(b => ({ name: b.name })) });
        }
        if (action === "commits") {
          const d = await githubApi(secret, `/repos/${repo}/commits?per_page=20`);
          return respond({ items: d.map(c => ({ sha: c.sha.slice(0, 7), title: String(c.commit?.message || "").split("\n")[0].slice(0, 160), user: c.commit?.author?.name, url: c.html_url })) });
        }
        if (action === "file") {
          const p = String(body.path || "").replace(/^\/+/, "");
          if (!p || p.length > 300 || p.includes("..")) return respond({ error: "Enter a valid file path." }, 400);
          const d = await githubApi(secret, `/repos/${repo}/contents/${p.split("/").map(encodeURIComponent).join("/")}`);
          if (Array.isArray(d)) return respond({ items: d.map(f => ({ name: f.name, type: f.type, path: f.path })) });
          if (d.encoding !== "base64" || d.size > 200000) return respond({ error: "File is too large to preview." }, 400);
          const text = new TextDecoder().decode(Uint8Array.from(atob(String(d.content).replace(/\n/g, "")), c => c.charCodeAt(0)));
          return respond({ file: { path: d.path, content: text.slice(0, 60000) } });
        }
        if (action === "create_issue") {
          if (body.confirm !== true) return respond({ error: "Please confirm this action." }, 400);
          rateLimit("ghw:" + userId, 10, 60000);
          const title = String(body.title || "").trim().slice(0, 250);
          const text = String(body.body || "").slice(0, 20000);
          if (!title) return respond({ error: "Enter an issue title." }, 400);
          const d = await githubApi(secret, `/repos/${repo}/issues`, { method: "POST", body: JSON.stringify({ title, body: text }) });
          return respond({ ok: true, number: d.number, url: d.html_url });
        }
        return respond({ error: "Unknown action." }, 400);
      } catch (error) {
        console.error("Plugin tool error:", String(error?.message || error));
        return respond({ error: safeMessage(error, "Plugin request failed.") }, getErrorStatus(error));
      }
    }

    /* ---------------- Plugins (connect / disconnect / list) ---------------- */
    if (
      request.method === "POST" &&
      ["/api/plugins/list", "/api/plugins/connect", "/api/plugins/disconnect"].includes(path)
    ) {
      let body;
      try { body = await request.json(); } catch { return respond({ error: "Please send valid JSON." }, 400); }

      try {
        const userId = await authenticateFirebase(request, env);

        await env.DB.prepare(
          `CREATE TABLE IF NOT EXISTS user_plugins (
             user_id TEXT NOT NULL, plugin TEXT NOT NULL, secret TEXT NOT NULL,
             label TEXT, meta TEXT, connected_at TEXT NOT NULL,
             PRIMARY KEY (user_id, plugin))`
        ).run();

        if (path === "/api/plugins/list") {
          const rows = await env.DB.prepare(
            `SELECT plugin, label, connected_at FROM user_plugins WHERE user_id = ?`
          ).bind(userId).all();
          const connected = {};
          for (const r of rows.results || []) connected[r.plugin] = { label: r.label || "", connectedAt: r.connected_at };
          return respond({ connected });
        }

        const plugin = String(body.plugin || "").toLowerCase();

        if (path === "/api/plugins/disconnect") {
          await env.DB.prepare(`DELETE FROM user_plugins WHERE user_id = ? AND plugin = ?`).bind(userId, plugin).run();
          return respond({ ok: true });
        }

        /* ----- connect ----- */
        if (!env.PLUGIN_SECRET_KEY) {
          return respond({ error: "Plugins are not configured yet (missing PLUGIN_SECRET_KEY on the server)." }, 500);
        }

        let secret, label, meta = {};

        if (plugin === "github") {
          const token = String(body.token || "").trim();
          if (token.length < 20 || token.length > 400) return respond({ error: "Please paste a valid GitHub token." }, 400);
          const r = await fetch("https://api.github.com/user", {
            headers: { "Authorization": `Bearer ${token}`, "User-Agent": "Malie-AI", "Accept": "application/vnd.github+json" },
            signal: AbortSignal.timeout(15000)
          }).catch(() => null);
          if (!r || !r.ok) return respond({ error: "GitHub did not accept that token. Check it and its permissions." }, 400);
          const u = await r.json().catch(() => ({}));
          secret = token;
          label = u.login ? "@" + u.login : "GitHub account";
        } else if (plugin === "omniroute") {
          let base = String(body.baseUrl || "").trim().replace(/\/+$/, "");
          const apiKey = String(body.apiKey || "").trim();
          let parsed;
          try { parsed = new URL(base); } catch { return respond({ error: "Enter a valid OmniRoute URL, e.g. https://omniroute.example.com/v1" }, 400); }
          if (!isSafePublicUrl(base)) {
            return respond({ error: "OmniRoute must be a public https:// address (a local address can't be reached from the server)." }, 400);
          }
          if (!apiKey || apiKey.length > 400) return respond({ error: "Please enter your OmniRoute API key." }, 400);
          if (!/\/v1$/.test(base)) base += "/v1";
          const r = await fetch(base + "/models", {
            redirect: "error",
            headers: { "Authorization": `Bearer ${apiKey}` },
            signal: AbortSignal.timeout(15000)
          }).catch(() => null);
          if (!r || !r.ok) return respond({ error: "Could not reach OmniRoute with that URL and key." }, 400);
          secret = apiKey;
          label = parsed.hostname;
          meta = { baseUrl: base };
        } else {
          return respond({ error: "This plugin isn't available to connect yet." }, 400);
        }

        // Encrypt the secret (AES-GCM) before storing it.
        const keyBytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(env.PLUGIN_SECRET_KEY));
        const aesKey = await crypto.subtle.importKey("raw", keyBytes, "AES-GCM", false, ["encrypt"]);
        const iv = crypto.getRandomValues(new Uint8Array(12));
        const ct = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, aesKey, new TextEncoder().encode(secret)));
        const packed = new Uint8Array(iv.length + ct.length);
        packed.set(iv, 0); packed.set(ct, iv.length);
        let bin = ""; for (const b of packed) bin += String.fromCharCode(b);
        const stored = btoa(bin);

        await env.DB.prepare(
          `INSERT INTO user_plugins (user_id, plugin, secret, label, meta, connected_at)
           VALUES (?, ?, ?, ?, ?, ?)
           ON CONFLICT(user_id, plugin) DO UPDATE SET secret = excluded.secret, label = excluded.label,
             meta = excluded.meta, connected_at = excluded.connected_at`
        ).bind(userId, plugin, stored, label, JSON.stringify(meta), new Date().toISOString()).run();

        return respond({ ok: true, label });
      } catch (error) {
        console.error("Plugins error:", String(error?.message || error));
        return respond({ error: error.message || "Plugin request failed." }, getErrorStatus(error));
      }
    }

    /* ---------------- Chat / Write / Quiz / Excel ---------------- */
    if (
      request.method === "POST" &&
      ["/api/chat", "/api/write", "/api/quiz", "/api/excel", "/api/pptx"].includes(path)
    ) {
      let body;

      try {
        body = await request.json();
      } catch {
        return respond({ error: "Please send valid JSON." }, 400);
      }

      try {
        const userId = await authenticateFirebase(request, env);

        /* ----- Chat ----- */
        if (path === "/api/chat") {
          const message = String(body.message || "").trim();

          if (!message) {
            return respond({ error: "Please enter a message." }, 400);
          }

          if (message.length > 30000) {
            return respond(
              { error: "Your message is too long. Please shorten it." },
              400
            );
          }

          const requestedMode = String(body.mode || "").trim().toLowerCase();

          // Look up the plan once; reuse it for the Excel check and usage limit.
          const plan = await getActivePlan(env, userId);

          if (requestedMode === "excel") {
            // Excel Assistant is available to Premium and Pro subscribers.
            if (plan !== "premium" && plan !== "pro") {
              return respond(
                { error: "Please subscribe to Premium or Pro to use this feature." },
                403
              );
            }
          }

          // Optional premium model (KrakenBot / LEO): paid plans only.
          const premiumKey = String(body.model || "").trim().toLowerCase();
          const usePremiumModel = Object.prototype.hasOwnProperty.call(PREMIUM_MODELS, premiumKey);
          if (usePremiumModel && plan !== "premium" && plan !== "pro") {
            return respond(
              { error: "This model is available on Premium and Pro. Please upgrade to use it." },
              403
            );
          }

          const omniModel = String(body.omniModel || "").trim().slice(0, 120);
          if (omniModel) {
            rateLimit("omni:" + userId, 20, 60000);
            const usageO = await reserveDailyUsage(env, userId, plan);
            try {
              const replyO = await callOmniRoute(env, userId, omniModel, message);
              return respond({ reply: replyO, usage: usageO });
            } catch (error) {
              await refundDailyUsage(env, userId, usageO.date);
              throw error;
            }
          }

          const usage = await reserveDailyUsage(env, userId, plan);

          let modelUsage = null;
          if (usePremiumModel) {
            try {
              modelUsage = await reservePremiumModelUsage(env, userId, premiumKey, plan, estimateCreditCost(message));
            } catch (error) {
              await refundDailyUsage(env, userId, usage.date);
              throw error;
            }
          }

          const identityQuestion =
            /\b(?:who\s+(?:created|made|built|developed|owns|founded)\s+(?:you|malie(?:\s+ai)?|this\s+(?:ai|assistant|product))|who\s+(?:is\s+)?(?:the\s+)?(?:creator|founder|owner)\s+of\s+(?:malie(?:\s+ai)?|you|this\s+(?:ai|assistant|product))|who\s+(?:is\s+)?your\s+(?:creator|founder|owner)|who\s+is\s+behind\s+(?:malie(?:\s+ai)?|you)|(?:are\s+you|is\s+malie(?:\s+ai)?)\s+(?:created|made|built|developed)\s+by)\b/i;

          if (identityQuestion.test(message)) {
            if (modelUsage) await refundPremiumModelUsage(env, userId, premiumKey, modelUsage.date, modelUsage.cost);
            return respond({
              reply: usePremiumModel
                ? "I am Malie AI, developed by Campus Creative, founded by Osume Prosper."
                : "I am Malie AI, developed by Campus Creative, founded by Osume Prosper. Google provides the model technology; Campus Creative created the Malie AI product.",
              usage
            });
          }

          let reply;

          try {
            if (usePremiumModel) {
              reply = await callPremiumModel(
                env,
                premiumKey,
                [
                  ...(plan === "pro"
                    ? ["This user is on the Pro plan: give thorough, expert-level answers with worked examples and a short summary."]
                    : ["This user is on the Premium plan: give detailed answers with a clear worked example where useful."]),
                  ...languageLine(body.language),
                  "",
                  "User message:",
                  message
                ].join("\n"),
                CHAT_MAX_TOKENS[plan] || CHAT_MAX_TOKENS.free
              );
              return respond({ reply, usage, modelUsage });
            }
            reply = await callAI(
              env,
              [
                "Respond to the user's message as Malie AI.",
                "If the user is learning something, teach it in clear steps.",
                "For calculations, show the working and explain how the result was obtained.",
                "For school subjects, use simple explanations and examples appropriate to the question.",
                "If the user asks for a direct answer, give it first and then explain it briefly.",
                ...(plan === "pro"
                  ? ["This user is on the Pro plan: give thorough, expert-level answers with worked examples, edge cases, and a short summary at the end."]
                  : plan === "premium"
                    ? ["This user is on the Premium plan: give detailed answers with a clear worked example where useful."]
                    : []),
                ...languageLine(body.language),

                ...(requestedMode === "excel"
                  ? [
                      "You are now Malie AI's Premium Excel Assistant.",
                      "Help with Excel formulas, spreadsheet assignments, data analysis, and troubleshooting spreadsheet errors.",
                      "For formulas, give the exact formula, explain where to enter it, and explain how it works.",
                      "For calculations, show the steps and check the arithmetic where possible.",
                      "For spreadsheet errors, explain the likely cause and provide a specific fix.",
                      "When users paste tables, analyze only the data they provide. Never invent missing data.",
                      "If cell references or column meanings are unclear, ask a brief clarifying question.",
                      "If the user wants a downloadable .xlsx file, tell them to use the Excel File Builder, where they can paste a table or upload a photo and download the finished spreadsheet."
                    ]
                  : []),

                "",
                "User message:",
                message
              ].join("\n"),
              false,
              null,
              { maxTokens: CHAT_MAX_TOKENS[plan] || CHAT_MAX_TOKENS.free }
            );
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            if (modelUsage) await refundPremiumModelUsage(env, userId, premiumKey, modelUsage.date, modelUsage.cost);
            throw error;
          }

          return respond({ reply, usage });
        }

        /* ----- Write ----- */
        if (path === "/api/write") {
          const mode = String(body.mode || "correct").trim();
          const inputText = String(body.text || "").trim();
          const options = body.options || {};

          if (!inputText) {
            return respond({ error: "Please enter some text." }, 400);
          }

          if (inputText.length > 30000) {
            return respond(
              { error: "Your text is too long. Please shorten it and try again." },
              400
            );
          }

          const instruction =
            taskInstructions[mode] || taskInstructions.correct;

          const usage = await reserveDailyUsage(env, userId);

          let result;

          try {
            result = await callAI(
              env,
              [
                "Task: " + instruction,
                "User options: " + JSON.stringify(options),
                "Follow the selected task carefully.",
                "Preserve names, dates, figures, and other supplied facts.",
                "Do not add unsupported information.",
                "",
                "Text to process:",
                inputText
              ].join("\n")
            );
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          return respond({ result, usage });
        }

        /* ----- Quiz ----- */
        if (path === "/api/quiz") {
          const inputText = String(body.text || "").trim();

          const count = Math.max(
            1,
            Math.min(20, Number.parseInt(body.count, 10) || 5)
          );

          const difficulty = String(body.difficulty || "medium").trim();

          if (!inputText) {
            return respond(
              { error: "Please provide notes or text to generate a quiz." },
              400
            );
          }

          if (inputText.length > 30000) {
            return respond(
              { error: "Your source text is too long. Please shorten it and try again." },
              400
            );
          }

          const prompt = [
            "Create exactly " + count + " multiple-choice questions based ONLY on the source material below.",
            "Difficulty: " + difficulty + ".",
            "Each question must have exactly four answer options.",
            "There must be exactly one correct answer per question.",
            "Vary the correct answer position across questions.",
            "Avoid duplicate questions and unsupported facts.",
            "Make the questions test understanding, not only memorisation, when appropriate.",
            "The explanation must teach the learner why the correct option is right.",
            "For calculations, include the key reasoning or working in the explanation.",
            "Keep explanations concise, clear, and suitable for the requested difficulty.",
            "Use answerIndex as a zero-based number: 0 for the first option, 1 for the second, 2 for the third, or 3 for the fourth.",
            "Do not put letters such as A or B in answerIndex.",
            "Return ONLY valid JSON in this exact structure:",
            '{"questions":[{"question":"Question text","options":["Option 1","Option 2","Option 3","Option 4"],"answerIndex":0,"explanation":"Explain why the correct option is right."}]}',
            "",
            "Source material:",
            inputText
          ].join("\n");

          const resultText = await callAI(env, prompt, true);
          const quiz = parseQuiz(resultText);

          if (quiz.questions.length !== count) {
            throw new Error(
              `The AI returned ${quiz.questions.length} questions instead of ${count}. Please generate the quiz again.`
            );
          }

          return respond(quiz);
        }

        /* ----- Excel file builder (Premium / Pro) ----- */
        if (path === "/api/excel") {
          const plan = await getActivePlan(env, userId);

          if (plan !== "premium" && plan !== "pro") {
            return respond(
              { error: "Please subscribe to Premium or Pro to use this feature." },
              403
            );
          }

          const message = String(body.message || "").trim();
          const image = body.image ? String(body.image) : null;

          if (!message && !image) {
            return respond(
              { error: "Send a table, a photo, or instructions." },
              400
            );
          }

          if (message.length > 30000) {
            return respond({ error: "Your message is too long." }, 400);
          }

          if (
            image &&
            (image.length > 5_000_000 ||
              !/^data:image\/(png|jpeg|webp);base64,/.test(image))
          ) {
            return respond(
              { error: "Please upload a PNG, JPG or WebP image under about 3.5 MB." },
              400
            );
          }

          const usage = await reserveDailyUsage(env, userId, plan);

          const prompt = [
            "Convert the user's table/figures (text and/or image) into a spreadsheet.",
            "Copy values exactly. Never invent data. If something is unreadable, leave the cell empty.",
            "Put headers in row 1. Use real numbers (not strings) for numeric values.",
            "Where the user asks for totals, averages or calculations, use Excel formulas as strings starting with '=' (for example \"=SUM(B2:B10)\").",
            "Return ONLY valid JSON in this exact structure:",
            '{"filename":"short_name","sheets":[{"name":"Sheet1","rows":[["Header A","Header B"],["x",1]],"formats":["text","currency"]}]}',
            "\"formats\" is optional: one entry per column, each one of text, integer, number, currency, percent or date. Use percent only when the numbers are fractions such as 0.15.",
            plan === "pro"
              ? "You may use several sheets. When the data has numbers, add a second sheet named Summary with totals, averages, minimum and maximum using formulas that refer to the data sheet."
              : "Use one sheet unless the user clearly asks for more.",
            "",
            "User instructions/data:",
            message || "(see image)"
          ].join("\n");

          let fileBase64;
          let spec;
          const excelLimits = EXCEL_LIMITS[plan] || EXCEL_LIMITS.premium;

          try {
            const text = await callAI(env, prompt, true, image, { maxTokens: excelLimits.maxTokens });
            spec = parseJsonObject(
              text,
              "The AI did not return a valid spreadsheet. Please try again."
            );
            fileBase64 = buildWorkbook(spec, excelLimits);
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          const filename =
            (String(spec.filename || "malie_sheet")
              .replace(/[^\w-]/g, "_")
              .slice(0, 40) || "malie_sheet") + ".xlsx";

          return respond({
            ok: true,
            filename,
            fileBase64,
            preview: spec.sheets?.[0]?.rows?.slice(0, excelLimits.preview) || [],
            usage
          });
        }

        /* ----- PowerPoint builder (Premium / Pro) ----- */
        if (path === "/api/pptx") {
          const plan = await getActivePlan(env, userId);

          if (plan !== "premium" && plan !== "pro") {
            return respond(
              { error: "Please subscribe to Premium or Pro to use this feature." },
              403
            );
          }

          const message = String(body.message || "").trim();

          if (!message) {
            return respond(
              { error: "Tell me the topic, or paste the notes to turn into slides." },
              400
            );
          }

          if (message.length > 30000) {
            return respond({ error: "Your message is too long." }, 400);
          }

          const usage = await reserveDailyUsage(env, userId, plan);
          const wantImages = wantsSlideImages(message, body.images);
          const deckLayout = detectDeckLayout(message);
          const themeName = detectTheme(message) || "dark";
          lastImageError = "";
          const maxImages = plan === "pro" ? 10 : 6;

          const prompt = [
            "Create a clear PowerPoint presentation from the user's request or notes.",
            "Use the user's facts exactly. Never invent statistics, quotes or citations.",
            "If the user gave a slide count, follow it. Otherwise use 6 to 10 slides.",
            "Each slide has a short title and 3 to 6 short bullet points (under 18 words each).",
            "Do not include a title slide in the slides list; the title and subtitle go in the top-level fields.",
            'Each slide has a "layout" chosen from: "plain" (bullets), "cards" (3 to 6 short parallel points shown as cards), "numbered" (3 to 6 ordered steps), "timeline" (3 to 6 points in time order), "statement" (the first bullet is one key sentence shown large, the rest are small supporting bullets), "columns" (bullets in two columns)' +
              (wantImages
                ? ', "right" or "left" (picture beside the bullets), "splitright" or "splitleft" (picture fills almost half the slide), "banner" (wide picture on top), "background" (full-slide picture), "caption" (full-slide picture with a text panel).'
                : ". Do not use picture layouts."),
            "Match each layout to its content and use a good mix across the deck instead of repeating one layout. Follow any layout the user asks for, for all slides or for named slides.",
            deckLayout ? `The user chose the layout "${deckLayout}" for the whole deck. Use it on every slide.` : "",
            ...(wantImages
              ? [
                  `The user wants pictures. Give "titleImage" for the title slide, and an "image" field on up to ${maxImages - 1} slides that use a picture layout (empty string elsewhere).`,
                  "Each image value is a vivid visual description in English, 10 to 25 words, of a concrete scene or object that matches that slide's topic.",
                  "Image descriptions must not contain text, letters, numbers, charts, logos, brand names or real named people."
                ]
              : []),
            ...languageLine(body.language, "Write all slide text"),
            "Return ONLY valid JSON in this exact structure:",
            wantImages
              ? '{"filename":"short_name","title":"Deck title","subtitle":"Short subtitle","titleImage":"Visual description","slides":[{"title":"Slide title","layout":"cards","bullets":["Point one","Point two"],"image":"Visual description"}]}'
              : '{"filename":"short_name","title":"Deck title","subtitle":"Short subtitle","slides":[{"title":"Slide title","bullets":["Point one","Point two"]}]}',
            "",
            "User request/notes:",
            message
          ].join("\n");

          let fileBase64;
          let spec;
          let imageCount = 0;

          try {
            const text = await callAI(env, prompt, true, null, { maxTokens: 4096 });
            spec = parseJsonObject(
              text,
              "The AI did not return a valid presentation. Please try again."
            );

            const images = { title: null, slides: [] };

            if (wantImages) {
              const norm = normalizeSlides(spec);
              const jobs = [];
              // If the AI left a description out, build one from the slide's own title.
              jobs.push({ slot: "title", description: norm.titleImage || `${norm.deckTitle}, ${norm.subtitle || "cover image"}` });
              norm.slides.forEach((s, i) => {
                if (!IMAGE_LAYOUTS.has(pickLayout(s.layout, deckLayout, i, true))) return;
                jobs.push({ slot: i, description: s.image || `${s.title}, related to ${norm.deckTitle}` });
              });

              const picked = jobs.slice(0, maxImages);
              const made = await generateImages(env, picked);

              picked.forEach((job, k) => {
                const img = made[k];
                if (!img) return;
                imageCount++;
                if (job.slot === "title") images.title = img;
                else images.slides[job.slot] = img;
              });
            }

            fileBase64 = buildPresentation(spec, images, deckLayout, themeName);
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            throw error;
          }

          const filename =
            (String(spec.filename || spec.title || "malie_presentation")
              .replace(/[^\w-]/g, "_")
              .slice(0, 40) || "malie_presentation") + ".pptx";

          return respond({
            ok: true,
            filename,
            fileBase64,
            preview: (Array.isArray(spec.slides) ? spec.slides : [])
              .slice(0, 12)
              .map(s => String(s?.title || "").slice(0, 120)),
            title: String(spec.title || "").slice(0, 120),
            images: imageCount,
            imagesRequested: wantImages,
            imageError: imageCount ? "" : lastImageError,
            layout: deckLayout || "",
            theme: themeName,
            usage
          });
        }
      } catch (error) {
        const status = getErrorStatus(error);

        console.error("API error:", { path, status, message: error.message });

        return respond(
          { error: safeMessage(error) },
          status
        );
      }
    }

    /* ---------------- Admin check (any signed-in user) ---------------- */
    if (request.method === "POST" && path === "/api/admin/check") {
      try {
        const uid = await authenticateFirebase(request, env);
        const adminUid = env.MALIE_ADMIN_UID;
        return respond({ ok: true, isAdmin: Boolean(adminUid) && uid === adminUid });
      } catch (error) {
        return respond({ ok: false, isAdmin: false, error: error.message }, error.status || 401);
      }
    }

    /* Advert Studio access check â€” does not start a render */
    if (request.method === "POST" && path === "/api/advert/access") {
      try {
        const userId = await authenticateFirebase(request, env);
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond({
            error: "Malie AI Advert Studio is available on the Pro plan only."
          }, 403);
        }

        return respond({ ok: true, plan: "pro" });
      } catch (e) {
        return respond({
          error: e.message || "Could not verify Advert Studio access."
        }, e.status || 500);
      }
    }

    /* ---------------- Malie AI Advert Studio (Pro only) ---------------- */
    if (
      request.method === "POST" &&
      (path === "/api/advert/start" || path === "/api/advert/status")
    ) {
      try {
        const userId = await authenticateFirebase(request, env);
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond({
            error: "Malie AI Advert Studio is available on the Pro plan only."
          }, 403);
        }

        if (!env.ETERNAL_AI_API_KEY) {
          return respond({ error: "Advert Studio is not configured yet." }, 500);
        }

        if (!env.DB) {
          return respond({ error: "The database is not configured." }, 500);
        }

        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const apiHeaders = {
          "Authorization": `Bearer ${env.ETERNAL_AI_API_KEY}`,
          "Content-Type": "application/json"
        };

        if (path === "/api/advert/status") {
          const projectId = String(body.projectId || "");

          if (!/^[A-Za-z0-9-]{16,80}$/.test(projectId)) {
            return respond({ error: "Invalid video project ID." }, 400);
          }

          const owned = await env.DB.prepare(
            "SELECT project_id, status, video_url FROM advert_projects WHERE project_id = ? AND user_id = ?"
          ).bind(projectId, userId).first();

          if (!owned) {
            return respond({ error: "Video project not found for this account." }, 404);
          }

          if (owned.status === "completed" && owned.video_url) {
            return respond({
              ok: true,
              projectId,
              status: "completed",
              videoUrl: owned.video_url,
              message: "Your advert video is ready. Save it promptly; the provider URL may expire after about 24 hours."
            });
          }

          const statusResponse = await fetch(
            "https://open.eternalai.org/api/image-to-video/" +
              encodeURIComponent(projectId) + "/status",
            {
              headers: { "Authorization": `Bearer ${env.ETERNAL_AI_API_KEY}` },
              signal: AbortSignal.timeout(15000)
            }
          );

          const statusData = await statusResponse.json().catch(() => ({}));
          const result = statusData.result || {};
          console.log("ETERNAL STATUS HTTP:", statusResponse.status);
          console.log("ETERNAL STATUS BODY:", JSON.stringify(statusData));

          if (!statusResponse.ok || (!statusData.result && statusData.status !== true)) {
            return respond({
              error: statusData.error || "Could not retrieve the video status. Please try again."
            }, 502);
          }

          const providerStatus = String(result.status || "processing").toLowerCase();

          // Eternal AI puts the finished video at result.output.video_url
          // (other shapes are checked too).
          const foundUrl =
            result.output?.video_url ||
            result.video_url ||
            (Array.isArray(result.output)
              ? result.output[0]?.video_url || result.output[0]
              : null) ||
            result.output?.url ||
            result.url ||
            null;

          const DONE = ["success", "succeeded", "completed", "complete", "ready", "done", "finished"];
          const FAILED = ["failed", "error", "cancelled", "canceled", "timeout"];

          let status;
          if (typeof foundUrl === "string" && foundUrl) status = "completed";
          else if (DONE.includes(providerStatus)) status = "completed";
          else if (FAILED.includes(providerStatus)) status = "failed";
          else status = providerStatus; // pending / processing / queued

          const videoUrl =
            status === "completed"
              ? (typeof foundUrl === "string" ? foundUrl : null)
              : owned.video_url;

          await env.DB.prepare(
            "UPDATE advert_projects SET status = ?, video_url = ? WHERE project_id = ? AND user_id = ?"
          ).bind(status, videoUrl || null, projectId, userId).run();

          return respond({
            ok: true,
            projectId,
            status,
            videoUrl: videoUrl || null,
            progress: result.progress ?? null,
            message: status === "completed"
              ? "Your advert video is ready. Save it promptly; the provider URL may expire after about 24 hours."
              : status === "failed"
                ? (result.error || statusData.error || "The video generation did not complete.")
                : "Your advert video is still rendering."
          });
        }

        const today = getUsageDate();
        const usage = await env.DB.prepare(
          "SELECT COUNT(*) AS total FROM advert_projects WHERE user_id = ? AND usage_date = ?"
        ).bind(userId, today).first();

        if (Number(usage?.total || 0) >= 3) {
          return respond({
            error: "You have reached the limit of 3 advert video renders per day. Please try again tomorrow."
          }, 429);
        }

        const imageData = String(body.imageData || "");
        const match = imageData.match(
          /^data:(image\/(?:jpeg|png|webp));base64,([A-Za-z0-9+/=\r\n]+)$/
        );

        if (!match) {
          return respond({ error: "Upload a valid JPG, PNG, or WebP product image." }, 400);
        }

        const base64 = match[2].replace(/\s/g, "");

        if (!base64 || base64.length > 20000000) {
          return respond({ error: "Image is too large. Please use an image smaller than 15 MB." }, 413);
        }

        let imageBytes;
        try {
          imageBytes = atob(base64);
        } catch {
          return respond({ error: "The image data is invalid." }, 400);
        }

        if (!imageBytes.length || imageBytes.length > 15 * 1024 * 1024) {
          return respond({ error: "Image must be smaller than 15 MB." }, 413);
        }

        const cleanText = (value, maxLength) => String(value || "")
          .replace(/[<>]/g, "")
          .trim()
          .slice(0, maxLength);

        let rawHeadline = cleanText(body.headline, 400);
        let rawDescription = cleanText(body.description, 400);
        const split = rawHeadline.match(/^\s*(?:headline\s*[:\-]\s*)?(.*?)\s*description\s*[:\-]\s*(.*)$/i);

        if (split) {
          rawHeadline = split[1];
          if (!rawDescription) rawDescription = split[2];
        }

        const stripLabel = (value, label) =>
          value.replace(new RegExp("^\\s*" + label + "\\s*[:\\-]\\s*", "i"), "").trim();

        const headline = stripLabel(rawHeadline, "headline").slice(0, 70) || "A Product Worth Discovering";
        const description = stripLabel(rawDescription, "description").slice(0, 120);
        const callToAction = stripLabel(cleanText(body.callToAction, 40), "call to action") || "Shop Now";

        const videoStyle = ["clean", "fullscreen", "promo", "story"].includes(body.style)
          ? body.style
          : "clean";

        const styleDirection = {
          clean: "Elegant, clean studio commercial. Keep the complete product visible and centred with generous negative space; use a slow cinematic push-in and soft premium lighting.",
          fullscreen: "Full-frame cinematic product commercial with tasteful camera motion and dramatic but realistic lighting. Keep the product recognizable and do not crop important product details.",
          promo: "Energetic premium promotional commercial with a smooth camera move, polished lighting changes, and a refined product reveal.",
          story: "Vertical social-story product commercial, composed for a 9:16 mobile frame. Keep the product fully visible and away from the edges, with smooth cinematic motion."
        }[videoStyle];

        const promptParts = [
          "Create a polished 5-second product advertisement from the supplied image.",
          styleDirection,
          "Preserve the product's exact shape, proportions, colours, packaging, label and existing logo. Do not redesign, replace, or distort the product.",
          "Use realistic studio lighting, natural reflections, subtle atmospheric movement, and smooth camera motion. Professional commercial quality.",
          `Headline concept: ${headline}.`,
          description ? `Supporting message: ${description}.` : "",
          `Call to action concept: ${callToAction}.`,
          "Do not generate new lettering, captions, fake labels, watermarks, or logos; keep any existing product text unchanged."
        ].filter(Boolean).join(" ").slice(0, 2000);

        const renderResponse = await fetch("https://open.eternalai.org/api/image-to-video", {
          method: "POST",
          headers: apiHeaders,
          body: JSON.stringify({
            prompt: promptParts,
            image_url: imageData,
            model_id: "wan-ai/wan2.2-i2v-a14b-lightning",
            duration: "5",
            resolution: "480p",
            aspect_ratio: videoStyle === "story" ? "9:16" : "16:9",
            negative_prompt: "blurry, distorted product, warped packaging, changed label, extra objects, illegible text, low quality, watermark",
            cfg_scale: 0.5
          }),
          signal: AbortSignal.timeout(30000)
        });

        const renderData = await renderResponse.json().catch(() => ({}));
        const requestId = renderData.result?.request_id;

        if (!renderResponse.ok || !requestId) {
          return respond({
            error: renderData.error || "Eternal AI could not start the video generation. Please try again.",
            provider_status: renderResponse.status
          }, renderResponse.status === 402 ? 402 : 502);
        }

        await env.DB.prepare(
          "INSERT INTO advert_projects (project_id, user_id, usage_date, status) VALUES (?, ?, ?, ?)"
        ).bind(requestId, userId, today, "pending").run();

        return respond({
          ok: true,
          projectId: requestId,
          status: "pending",
          message: "Your advert video has been submitted to Eternal AI. It may take a few minutes to finish.",
          durationSeconds: 5,
          resolution: "480p"
        });

      } catch (error) {
        console.error("Advert Studio error:", error);
        return respond({
          error: "Advert Studio encountered an error. Please try again."
        }, 500);
      }
    }
    /* ---------------- CAD Studio code generation (Pro only) ---------------- */

if (request.method === "POST" && path === "/api/cad/generate") {
  let body;

  try {
    body = await request.json();
  } catch {
    return respond({ error: "Please send valid JSON." }, 400);
  }

  try {
    const userId = await authenticateFirebase(request, env);
    const plan = await getActivePlan(env, userId);

    if (plan !== "pro") {
      return respond({
        error: "AI CAD generation is available on the Pro plan. Please upgrade to Pro to use it."
      }, 403);
    }

    const description = String(
      body.prompt || body.description || ""
    ).trim();

    if (description.length < 5) {
      return respond({
        error: "Describe the 3D object you want to create."
      }, 400);
    }

    if (description.length > 2000) {
      return respond({
        error: "Please keep your CAD description under 2,000 characters."
      }, 400);
    }

    if (!env.DB) {
      return respond({
        error: "Usage tracking is temporarily unavailable."
      }, 503);
    }

    const usage = await reserveDailyUsage(env, userId, plan);

    try {
      



const prompt = [
  "Generate reliable 3D CAD JavaScript for Cascade Studio.",
  "Return ONLY a JSON object with exactly two fields: title (short string) and code (string). No markdown fences.",
  "The code must use only the documented CascadeStudio standard library.",

"Supported primitives: " + Object.values(CAD_API.primitives).filter(item => !item.startsWith("new Sketch")).join(", ") + ".",

  "Supported sketch API: new Sketch([x,y], plane?).LineTo([x,y]).Fillet(r).ArcTo([midX,midY],[endX,endY]).BSplineTo(points).End(true).Face(). The plane may be XY, XZ, or YZ.",

"Supported operations: " + Object.values(CAD_API.operations).join(", ") + ".",

"Use Pipe(shape, path) with exactly two arguments. The path must be a valid wire of the type expected by the installed Cascade Studio engine. Never assume extra arguments or pass a Sketch object, face, or general shape as the path.",
"Use only documented CascadeStudio operations with their correct argument order and compatible geometry types.",
"For swept parts, construct a valid path and ensure the profile is correctly oriented to the path at its start.",
"Avoid duplicate path points, sharp unintended kinks, self-intersections, pinched ends, and invalid or incomplete solids.",
"For a 90-degree elbow, create a quarter-circle bend with the requested centerline radius and straight sections at both ends. Do not claim exact dimensions when using an approximate spline.",
"Translate(offset, shape) takes the translation vector first and the shape second. Rotate(axis, degrees, shape) takes the axis first, then degrees, then the shape.",
"The script executes at top level. Never use a return statement. Assign the completed model to a variable named finalModel.",
"Create hollow swept parts only using the exact Pipe signature and geometry types supported by the installed Cascade Studio engine. Reuse a centerline path only when the API supports it. Subtract the inner solid from the outer solid using Difference(outerPipe, [innerPipe]) when both solids are valid.",
"Subtract the inner pipe from the outer pipe using Difference(outerPipe, [innerPipe]) and assign the result to finalModel.",
"Never use return finalModel; at the end of the generated script.",

"Supported selectors and measurements: Edges(shape).ofType(type).parallel(axis).max(axis).min(axis).indices(), " + CAD_API.measurements.filter(name => name !== "Edges").map(name => name + "(shape)").join(", ") + ".",

  "For square or rectangular profiles, prefer Sketch with 2D points followed by End(true).Face() and Extrude(face,[0,0,height]).",
  "Example rectangular block: const face = new Sketch([-20,-20]).LineTo([20,-20]).LineTo([20,20]).LineTo([-20,20]).End(true).Face(); const finalModel = Extrude(face,[0,0,20]);",
  "Polygon requires an array of 3D points. Correct example: const face = Polygon([[0,0,0],[40,0,0],[40,40,0],[0,40,0]]); const finalModel = Extrude(face,[0,0,20]); Never pass a flat list of coordinates to Polygon.",
  "For a hollow cylinder, use Difference(Cylinder(outerRadius,height),[Translate([0,0,-1],Cylinder(innerRadius,height+2))]). Ensure the inner radius is smaller than the outer radius.",
  "For holes, use cylindrical cutters that extend at least 1 mm beyond both faces. Subtract the cutters from the actual body using Difference.",
  "Assign the completed requested model to a variable named finalModel. Preserve the result of every Boolean operation.",

  "Use const or let declarations and numeric coordinates. Use positive dimensions and sensible proportions. Negative coordinates are allowed in point arrays and translations.",
  "Do not use imports, custom functions, loops, conditions, browser APIs, arbitrary JavaScript, or markdown fences. Use only the exact supported function names listed in this prompt. Cube is not supported; use Box(side,side,side) for a cube. Never invent aliases, functions, or overloads.",
  
  "CRITICAL CODE RESTRICTIONS:",
  "Never use for, while, or do loops.",
  "Never use map, forEach, filter, or reduce.",
  "Do not generate helper functions or callbacks.",
  "Create geometry using direct, explicit calls to the documented CascadeStudio API.",
  "For repeated features such as bolt holes, write each operation explicitly.",
  "Return only valid JSON containing the model title and code.",
  "Before returning, check that the code contains no forbidden JavaScript keywords or unsupported syntax.",
  "Generate actual 3D geometry, not a description. Choose operations appropriate to the requested shape and its dimensions.",
  "User requested object:",
  description
].join("\n");




     
      const cadGenerationStartedAt = Date.now();

console.log("CAD generation started", JSON.stringify({
  promptLength: description.length,
  timestamp: new Date().toISOString()
}));

let raw;

try {
  raw = await callNvidiaAI(env, prompt, true);

  console.log("CAD generation completed", JSON.stringify({
    durationMs: Date.now() - cadGenerationStartedAt,
    responseLength: raw.length
  }));
} catch (error) {
  console.error("CAD generation failed", JSON.stringify({
    durationMs: Date.now() - cadGenerationStartedAt,
    message: String(error?.message || error),
    name: error?.name || null,
    status: error?.status || null
  }));

  throw error;
}

      let result;

try {
  result = parseJsonObject(
    raw,
    "Initial CAD response was not valid JSON."
  );
} catch (error) {
  console.error("CAD initial response invalid JSON", JSON.stringify({
    message: String(error?.message || error),
    responseLength: String(raw || "").length,
    responsePreview: String(raw || "").slice(0, 1200)
  }));

  throw new Error(
    "Malie received an invalid CAD response. Please try again."
  );
}

      let title = String(
        result.title || "Generated CAD Model"
      ).replace(/[<>&"']/g, "").slice(0, 80);

      let code;

      try {
        code = validateCadCode(result.code);
      } catch (validationError) {
        const repairPrompt = [
          "Repair the invalid CascadeStudio CAD JavaScript below.",
          "The validation error was: " + String(validationError.message || validationError),
          "Return ONLY valid JSON with exactly two fields: title and code.",
          "Preserve the requested object's shape and dimensions.",
          "Use only documented CascadeStudio geometry operations with the correct argument order and geometry types.",
"Do not use loops, helper functions, comments, template strings, unsupported JavaScript, or top-level return statements.",
"Ensure every parenthesis and square bracket is correctly balanced and nested.",
"Supported sweep operation: Pipe(shape, path). Use exactly the documented two-argument signature unless the actual Cascade Studio implementation confirms otherwise.",
"Do not call GetWire unless it is confirmed to be a supported function and its required input type is known.",
"Before using Pipe, ensure the profile and path are the exact geometry types required by Cascade Studio. Do not pass a Sketch, face, or general shape where a wire is required.",
"For a hollow swept part, create compatible outer and inner profiles and use the documented sweep operation. Subtract the inner swept volume from the outer volume only when both operations return valid compatible shapes.",

"Supported primitives: " + Object.values(CAD_API.primitives).filter(item => !item.startsWith("new Sketch")).join(", ") + ".",
"Supported operations: " + Object.values(CAD_API.operations).join(", ") + ".",
"Do not invent function names, overloads, argument types, or extra arguments. Cube is unsupported; use Box(side,side,side) for a cube. Return one complete script with unique variable names and assign the final result to finalModel.",

"For hollow sweeps, use compatible outer and inner profiles and subtract the inner sweep from the outer sweep.",
"Preserve every requested component, dimension, hole, cavity, connection, and geometric feature while repairing the script.",
"Make the smallest necessary correction to fix the validation error.",
"Do not replace a detailed assembly with a primitive approximation unless the requested geometry is impossible with the documented API.",
"Return the entire corrected executable script, not a fragment or pseudocode.",
"Preserve the user's requested dimensions and shape when repairing invalid code.",
"Assign the completed model to finalModel.",
          "Requested object:",
          description,
          "Previous title:",
          title,
          "Previous invalid code:",
          String(result.code || "")
        ].join("\n");

        const repairedRaw = await callNvidiaAI(env, repairPrompt, true);

let repairedResult;

try {
  repairedResult = parseJsonObject(
    repairedRaw,
    "Repair response was not valid JSON."
  );
} catch (error) {
  console.error("CAD repair returned invalid JSON", JSON.stringify({
    message: String(error?.message || error),
    responseLength: String(repairedRaw || "").length,
    responsePreview: String(repairedRaw || "").slice(0, 1200)
  }));

  throw new Error(
    "The CAD repair model returned an invalid response. Please try a simpler model description."
  );
}

        title = String(
          repairedResult.title || title
        ).replace(/[<>&"']/g, "").slice(0, 80);

        code = validateCadCode(repairedResult.code);
      }
 

      return respond({
        ok: true,
        title,
        code,
        usage
      });
    } catch (error) {
      await refundDailyUsage(env, userId, usage.date);
      throw error;
    }
  } catch (error) {
    const status = [400, 401, 403, 429, 500, 503].includes(error?.status)
      ? error.status
      : 502;
console.error(
  "CAD generation error:",
  JSON.stringify({
    message: String(error?.message || error),
    status: error?.status || null,
    name: error?.name || null
  })
);

    return respond({
      error: error?.status
        ? error.message
        : (error?.message || "CAD generation failed. Please try again.")
    }, status);
  }
    }

    /* ---------------- Engineering Drawing Studio (Pro only) ---------------- */
    if (request.method === "POST" && path === "/api/drawing") {
      try {
        const userId = await authenticateFirebase(request, env);
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond(
            { error: "The Engineering Drawing Studio is available on the Pro plan. Please upgrade to Pro to use it." },
            403
          );
        }

        let body;
        try { body = await request.json(); }
        catch { return respond({ error: "Please send valid JSON." }, 400); }

        /* ----- Upload one image, get every view back ----- */
        if (body.action === "image") {
          const image = String(body.image || "");
          validateDrawingImage(image);

          const projection = body.projection === "first" ? "first" : "third";
          const theme = body.theme === "dark" ? "dark" : "paper";
          const drawUsage = await reserveDrawingUpload(env, userId);
          let usage;
          try {
            usage = await reserveDailyUsage(env, userId, plan);
          } catch (error) {
            await refundDrawingUpload(env, userId, drawUsage.date);
            throw error;
          }

          try {
            const ask = async (prompt) => {
              try {
                return await callAI(env, prompt, true, image, { maxTokens: 4096, model: DRAWING_VISION_MODEL, temperature: 0.1 });
              } catch (primaryError) {
                console.error("Drawing vision model failed, using fallback:", String(primaryError?.message || primaryError));
                return await callAI(env, prompt, true, image, { maxTokens: 4096, model: VISION_MODEL, temperature: 0.1 });
              }
            };

            const failMsg = "The AI could not describe that part. Please try a clearer image.";
            const firstText = await ask(DRAWING_PROMPT);
            const firstRaw = parseJsonObject(firstText, failMsg);
            let model = sanitizeModel(firstRaw);

            // Second pass: the AI checks its own result against the image.
            try {
              const secondText = await ask(DRAWING_REFINE_PROMPT + "\n" + JSON.stringify(firstRaw).slice(0, 6000));
              const refined = sanitizeModel(parseJsonObject(secondText, failMsg));
              model = refined;
            } catch (refineError) {
              console.error("Drawing refine pass skipped:", String(refineError?.message || refineError));
            }

            const views = buildAllViews(model, projection, theme);
            const { sectionPlane, ...viewSvgs } = views;

            return respond({
              ok: true,
              supported: true,
              mode: "image",
              title: model.title,
              unit: model.unit,
              estimated: model.estimated,
              confidence: model.confidence,
              notes: model.notes,
              warnings: model.warnings,
              projection,
              sectionPlane,
              views: viewSvgs,
              auxMessage:
                "An auxiliary view is only needed for inclined surfaces, so none is generated.",
              model,
              usage,
              drawingUsage: drawUsage
            });
          } catch (error) {
            await refundDailyUsage(env, userId, usage.date);
            await refundDrawingUpload(env, userId, drawUsage.date);
            throw error;
          }
        }

        /* ----- Original: draw the stepped flanged bush from typed dimensions ----- */
        const raw = body.params || {};
        const p = {
          d: Number(raw.d),
          D1: Number(raw.D1),
          L1: Number(raw.L1),
          D2: Number(raw.D2),
          T: Number(raw.T)
        };

        for (const [key, value] of Object.entries(p)) {
          if (!Number.isFinite(value) || value <= 0 || value > 1000) {
            return respond({ error: `Dimension "${key}" must be a number greater than 0 and not more than 1000.` }, 400);
          }
        }
        if (!(p.d < p.D1)) return respond({ error: "The bore diameter must be smaller than the body diameter." }, 400);
        if (!(p.D1 < p.D2)) return respond({ error: "The flange diameter must be larger than the body diameter." }, 400);

        const views = ["front", "top", "side", "iso", "section", "aux"];
        const view = views.includes(body.view) ? body.view : "section";
        const projection = body.projection === "first" ? "first" : "third";
        const unit = raw.unit === "in" ? "in" : "mm";
        const title = String(raw.title || "Stepped flanged bush").replace(/[<>&"']/g, "").slice(0, 60);

        if (view === "aux") {
          return respond({
            ok: true,
            supported: false,
            message: "An auxiliary view is only needed for inclined surfaces. This part is a body of revolution with no inclined faces, so no auxiliary view is generated."
          });
        }

        return respond({
          ok: true,
          supported: true,
          view,
          projection,
          params: { ...p, unit, title },
          totalLength: p.L1 + p.T
        });
      } catch (error) {
        console.error("Drawing error:", String(error?.message || error));
        return respond({ error: error.message || "Could not generate the drawing." }, getErrorStatus(error));
      }
    }

    /* ---------------- Malie Video Studio (Pro only) ---------------- */
    if (
      request.method === "POST" &&
      (path === "/api/video/access" || path === "/api/video/generate")
    ) {
      try {
        // 1. Who is this? (verified with Firebase, never trusted from the browser)
        const userId = await authenticateFirebase(request, env);

        // 2. Active Pro plan, read from the subscriptions table
        const plan = await getActivePlan(env, userId);

        if (plan !== "pro") {
          return respond(
            { error: "Malie Video Studio is available on the Pro plan only." },
            403
          );
        }

        const limit = getVideoDailyLimit(env);
        const secondsLimit = getVideoDailySeconds(env);
        const usageDate = getUsageDate();

        /* ----- Access check + remaining allowance (no generation, no cost) ----- */
        if (path === "/api/video/access") {
          const used = await getVideoUsage(env, userId, usageDate);

          return respond({
            ok: true,
            plan: "pro",
            usage: videoUsageSummary(limit, used.videos, used.seconds, secondsLimit, usageDate),
            options: {
              durations: VIDEO_STUDIO.durations.filter(
                seconds => seconds <= secondsLimit
              ),
              aspectRatios: VIDEO_STUDIO.aspectRatios,
              resolution: VIDEO_STUDIO.resolution,
              promptMin: VIDEO_STUDIO.promptMin,
              promptMax: VIDEO_STUDIO.promptMax
            }
          });
        }

        /* ----- Generate ----- */
        let body;
        try {
          body = await request.json();
        } catch {
          return respond({ error: "Please send valid JSON." }, 400);
        }

        const prompt = String(body.prompt || "").trim();
        const duration = Number(body.duration);
        const aspectRatio = String(body.aspectRatio || "");

        if (prompt.length < VIDEO_STUDIO.promptMin) {
          return respond(
            { error: `Please describe your video in at least ${VIDEO_STUDIO.promptMin} characters.` },
            400
          );
        }

        if (prompt.length > VIDEO_STUDIO.promptMax) {
          return respond(
            { error: `Please keep your description under ${VIDEO_STUDIO.promptMax} characters.` },
            400
          );
        }

        if (!VIDEO_STUDIO.durations.includes(duration)) {
          return respond({ error: "Choose a duration of 5 or 10 seconds." }, 400);
        }

        if (!VIDEO_STUDIO.aspectRatios.includes(aspectRatio)) {
          return respond({ error: "Choose Landscape (16:9), Portrait (9:16) or Square (1:1)." }, 400);
        }

        // Resolution is fixed on the server (480p) for cost control; any value
        // sent by the browser is ignored.

        if (!env.AI || typeof env.AI.run !== "function") {
          return respond(
            { error: "Video generation is not configured yet. The site owner must add the Workers AI binding named AI." },
            500
          );
        }

        // 3. Reserve one of today's videos BEFORE spending money. Atomic, so
        //    parallel requests cannot exceed the daily allowance.
        const usedAfterReserve = await reserveVideo(env, userId, usageDate, duration, limit, secondsLimit);

        let videoUrl;

        try {
          videoUrl = await runVideoModel(env, { prompt, duration, aspectRatio });
        } catch (error) {
          // 4. Give the video back when the failure is definite.
          if (error.refund !== false) {
            await refundVideo(env, userId, usageDate, duration);
          }
          throw error;
        }

        return respond({
          ok: true,
          videoUrl,
          durationSeconds: duration,
          aspectRatio,
          resolution: VIDEO_STUDIO.resolution,
          usage: videoUsageSummary(limit, usedAfterReserve.videos, usedAfterReserve.seconds, secondsLimit, usageDate)
        });
      } catch (error) {
        console.error("Video Studio error:", String(error?.message || error));

        // Errors we raised on purpose carry a status and a safe message.
        // Anything else (for example a missing table) gets a generic reply.
        const status = [400, 401, 403, 429, 500, 502, 503, 504].includes(error?.status)
          ? error.status
          : 500;

        return respond(
          {
            error: error?.status
              ? error.message
              : "Video Studio could not process your request. Please try again."
          },
          status
        );
      }
    }

    /* ---------------- App icons ---------------- */
    if (request.method === "GET" && APP_ICONS[path]) {
      return new Response(Uint8Array.from(atob(APP_ICONS[path]), c => c.charCodeAt(0)), {
        headers: {
          "Content-Type": "image/png",
          "Cache-Control": "public, max-age=604800"
        }
      });
    }

    /* ---------------- Fallbacks ---------------- */
    if (path.startsWith("/api/")) {
      return respond({ error: "API endpoint not found." }, 404);
    }

    if (env.ASSETS) {
      return env.ASSETS.fetch(request);
    }

    return new Response("Malie AI is running.", { headers: corsHeaders });
  }
};
