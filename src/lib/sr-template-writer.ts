import JSZip from "jszip";
import tplUrl from "@/assets/sr-template.xlsx?url";
import { SR_OUTPUT_HEADERS } from "./sale-return-converter";

/**
 * The ERP import template contains a hidden "Select Option Sheet" plus x14
 * data-validation dropdowns on the data sheet. Generating a fresh workbook with
 * SheetJS drops all of that, and the ERP import then rejects the file.
 * So we load the pristine template zip and only replace <sheetData>, keeping
 * every other part (styles, hidden sheet, validations, comments) byte-identical.
 */

let tplCache: ArrayBuffer | null = null;

async function loadTemplate(): Promise<ArrayBuffer> {
  if (tplCache) return tplCache;
  const res = await fetch(tplUrl);
  if (!res.ok) throw new Error("Sale Return template load nahi hua");
  tplCache = await res.arrayBuffer();
  return tplCache;
}

function colName(i: number): string {
  let n = i + 1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

const COL_NAMES = SR_OUTPUT_HEADERS.map((_, i) => colName(i));

function esc(v: string): string {
  return v
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    // strip control chars that make Excel reject the file
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

function cellXml(ref: string, val: unknown): string {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number" && Number.isFinite(val)) {
    return `<c r="${ref}"><v>${val}</v></c>`;
  }
  const s = String(val);
  if (s === "") return "";
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(s)}</t></is></c>`;
}

function buildSheetData(headerRowXml: string, rows: Record<string, unknown>[]): string {
  const parts: string[] = ["<sheetData>", headerRowXml];
  for (let i = 0; i < rows.length; i++) {
    const r = i + 2;
    const row = rows[i];
    let cells = "";
    for (let c = 0; c < SR_OUTPUT_HEADERS.length; c++) {
      cells += cellXml(`${COL_NAMES[c]}${r}`, row[SR_OUTPUT_HEADERS[c]]);
    }
    if (!cells) continue;
    parts.push(`<row r="${r}" spans="1:${SR_OUTPUT_HEADERS.length}" customFormat="1" ht="12.75">${cells}</row>`);
  }
  parts.push("</sheetData>");
  return parts.join("");
}

export async function buildSRWorkbookFromTemplate(
  rows: Record<string, unknown>[]
): Promise<ArrayBuffer> {
  const tpl = await loadTemplate();
  const zip = await JSZip.loadAsync(tpl);
  const sheetPath = "xl/worksheets/sheet1.xml";
  const xml = await zip.file(sheetPath)!.async("string");

  const sdStart = xml.indexOf("<sheetData>");
  const sdEnd = xml.indexOf("</sheetData>") + "</sheetData>".length;
  const existing = xml.slice(sdStart, sdEnd);
  const row1Match = existing.match(/<row [^>]*r="1"[\s\S]*?<\/row>/);
  const headerRowXml = row1Match ? row1Match[0] : "";

  const lastCol = COL_NAMES[SR_OUTPUT_HEADERS.length - 1];
  let out =
    xml.slice(0, sdStart) + buildSheetData(headerRowXml, rows) + xml.slice(sdEnd);
  out = out.replace(
    /<dimension ref="[^"]*"\/>/,
    `<dimension ref="A1:${lastCol}${Math.max(rows.length + 1, 1)}"/>`
  );

  zip.file(sheetPath, out);
  return (await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })) as ArrayBuffer;
}
