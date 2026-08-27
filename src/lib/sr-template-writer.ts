import JSZip from "jszip";
import tplUrl from "@/assets/sr-template.xlsx?url";
import { SR_OUTPUT_HEADERS } from "./sale-return-converter";

/**
 * The ERP import template contains a hidden "Select Option Sheet" plus x14
 * data-validation dropdowns on the data sheet. Generating a fresh workbook with
 * SheetJS drops all of that, and the ERP import then rejects the file.
 *
 * So we load the pristine template zip and update:
 * 1. <sheetData> in xl/worksheets/sheet1.xml with data rows referencing shared strings (t="s").
 * 2. xl/sharedStrings.xml with all text values added to the Shared Strings Table.
 *
 * This ensures 100% compatibility with Suvit / Tally ERP importers that require
 * Shared String format (t="s") and reject or skip inlineStr entries.
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
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    // strip control chars that make Excel reject the file
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

interface SharedStringsManager {
  sis: string[];
  stringToIdx: Map<string, number>;
  totalUsage: number;
  getStringIndex(s: string): number;
  buildSstXml(): string;
}

function createSharedStringsManager(initialSstXml: string, initialUsageCount: number): SharedStringsManager {
  const matches = initialSstXml.match(/<si>[\s\S]*?<\/si>/g);
  const sis: string[] = matches ? Array.from(matches) : [];
  const stringToIdx = new Map<string, number>();

  sis.forEach((si, idx) => {
    const txt = si.replace(/<[^>]+>/g, "");
    if (!stringToIdx.has(txt)) {
      stringToIdx.set(txt, idx);
    }
  });

  let totalUsage = initialUsageCount;

  return {
    sis,
    stringToIdx,
    get totalUsage() {
      return totalUsage;
    },
    getStringIndex(s: string): number {
      totalUsage++;
      const existing = stringToIdx.get(s);
      if (existing !== undefined) {
        return existing;
      }
      const newIdx = sis.length;
      sis.push(`<si><t xml:space="preserve">${esc(s)}</t></si>`);
      stringToIdx.set(s, newIdx);
      return newIdx;
    },
    buildSstXml(): string {
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totalUsage}" uniqueCount="${sis.length}">${sis.join("")}</sst>`;
    },
  };
}

function cellXml(ref: string, val: unknown, sst: SharedStringsManager): string {
  if (val === null || val === undefined || val === "") return "";
  if (typeof val === "number" && Number.isFinite(val)) {
    return `<c r="${ref}"><v>${val}</v></c>`;
  }
  const s = String(val);
  if (s === "") return "";
  const idx = sst.getStringIndex(s);
  return `<c r="${ref}" t="s"><v>${idx}</v></c>`;
}

function buildSheetData(
  headerRowXml: string,
  rows: Record<string, unknown>[],
  sst: SharedStringsManager
): string {
  const parts: string[] = ["<sheetData>", headerRowXml];
  for (let i = 0; i < rows.length; i++) {
    const r = i + 2;
    const row = rows[i];
    let cells = "";
    for (let c = 0; c < SR_OUTPUT_HEADERS.length; c++) {
      cells += cellXml(`${COL_NAMES[c]}${r}`, row[SR_OUTPUT_HEADERS[c]], sst);
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
  const s1Xml = await zip.file(sheetPath)!.async("string");
  const s2Xml = (await zip.file("xl/worksheets/sheet2.xml")?.async("string")) || "";
  const sstXml = (await zip.file("xl/sharedStrings.xml")?.async("string")) || "";

  const s2Refs = (s2Xml.match(/t="s"/g) || []).length;
  const sdStart = s1Xml.indexOf("<sheetData>");
  const sdEnd = s1Xml.indexOf("</sheetData>") + "</sheetData>".length;
  const existing = s1Xml.slice(sdStart, sdEnd);
  const row1Match = existing.match(/<row [^>]*r="1"[\s\S]*?<\/row>/);
  const headerRowXml = row1Match ? row1Match[0] : "";
  const headerRefs = (headerRowXml.match(/t="s"/g) || []).length;

  const sstManager = createSharedStringsManager(sstXml, s2Refs + headerRefs);

  const lastCol = COL_NAMES[SR_OUTPUT_HEADERS.length - 1];
  let out =
    s1Xml.slice(0, sdStart) +
    buildSheetData(headerRowXml, rows, sstManager) +
    s1Xml.slice(sdEnd);
  out = out.replace(
    /<dimension ref="[^"]*"\/>/,
    `<dimension ref="A1:${lastCol}${Math.max(rows.length + 1, 1)}"/>`
  );

  zip.file(sheetPath, out);
  zip.file("xl/sharedStrings.xml", sstManager.buildSstXml());

  return (await zip.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })) as ArrayBuffer;
}

