// Item Master — persistent item database (localStorage).
// New items from Purchase/Sale/SR uploads auto-merge; existing items get missing fields filled in.
// Strict deduplication ensures NO duplicate item names exist in DB or Differential Export.
import * as XLSX from "xlsx";

export const ITEM_MASTER_HEADERS = [
  "Item Name",
  "Alias Name",
  "Print Item Name",
  "Item Category",
  "Primary Unit Of Measurement",
  "Primary Unit of Conversion Rate",
  "Secondary Unit Of Measurement",
  "Secondary Unit of Conversion Rate",
  "Decimal Places",
  "SKU/Barcode",
  "Description",
  "Re Order Level",
  "Re Order UOM",
  "GST Applicable",
  "GST Rate",
  "HSN Code",
  "GST Cess Rate",
  "RCM Applicable",
  "MRP",
  "MRP Discount Type",
  "MRP Discount Value",
  "Selling Price Type",
  "Selling Price",
  "Secondary Unit Selling Price Type",
  "Secondary Unit Selling Price",
  "Income Ledger To Be Associated",
  "Discount Type",
  "Discount Value",
  "Purchase Price Type",
  "Purchase Price",
  "Purchase Discount Type",
  "Purchase Discount Value",
  "Expense Ledger To Be Associated",
  "Decimal Places For Rate",
  "Last Sale Price",
  "Last Purchase Price",
  "Opening Quantity Unit of Measurement",
  "Opening Quantity",
  "Opening Rate (Without GST)",
];

export interface ItemMasterEntry {
  name: string;
  hsn: string;
  gstRate: number;
  mrp: number;
  purchasePrice: number;
  lastSalePrice?: number;
  lastPurchasePrice?: number;
  description?: string;
}

const LS_KEY = "xlsxConverter.itemMaster.v1";

/** Clean string and remove invisible zero-width characters and collapse whitespace */
export function normalizeItemName(raw: unknown): string {
  if (raw === null || raw === undefined) return "";
  return String(raw)
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/** Primary lookup key for strict item deduplication */
export function normalizeItemKey(raw: unknown): string {
  return normalizeItemName(raw).toUpperCase();
}

/** Alphanumeric stripped key for fuzzy matching between slight punctuation differences */
export function cleanItemKey(raw: unknown): string {
  return normalizeItemName(raw).toLowerCase().replace(/[^a-z0-9]/g, "");
}

export function cleanHSN(v: unknown): string {
  if (v === null || v === undefined) return "";
  // Keep up to 8 digits numeric HSN
  return String(v).replace(/[^\d]/g, "").slice(0, 8);
}

function num(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[₹,\s]/g, ""));
  return isNaN(n) ? 0 : n;
}

/**
 * Load master items from localStorage.
 * Automatically cleans and deduplicates any previously stored entries by normalized item key.
 */
export function loadMaster(): Map<string, ItemMasterEntry> {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return new Map();
    const arr: ItemMasterEntry[] = JSON.parse(raw);
    const map = new Map<string, ItemMasterEntry>();

    let hadDuplicates = false;
    for (const e of arr) {
      const cleanName = normalizeItemName(e.name);
      if (!cleanName) continue;
      const key = normalizeItemKey(cleanName);
      const existing = map.get(key);

      if (!existing) {
        map.set(key, {
          name: cleanName,
          hsn: cleanHSN(e.hsn ?? ""),
          gstRate: num(e.gstRate),
          mrp: num(e.mrp),
          purchasePrice: num(e.purchasePrice),
          lastSalePrice: num(e.lastSalePrice),
          lastPurchasePrice: num(e.lastPurchasePrice || e.purchasePrice),
          description: typeof e.description === "string" ? e.description.trim() : "",
        });
      } else {
        hadDuplicates = true;
        // Merge missing attributes into existing unique item
        if (!existing.hsn && e.hsn) existing.hsn = cleanHSN(e.hsn);
        if (!existing.gstRate && e.gstRate) existing.gstRate = num(e.gstRate);
        if (!existing.mrp && e.mrp) existing.mrp = num(e.mrp);
        if (!existing.purchasePrice && e.purchasePrice) existing.purchasePrice = num(e.purchasePrice);
        if (!existing.lastSalePrice && e.lastSalePrice) existing.lastSalePrice = num(e.lastSalePrice);
        if (!existing.lastPurchasePrice && (e.lastPurchasePrice || e.purchasePrice)) {
          existing.lastPurchasePrice = num(e.lastPurchasePrice || e.purchasePrice);
        }
        if (!existing.description && e.description) existing.description = String(e.description).trim();
      }
    }

    // Persist cleaned deduplicated state if any duplicates or messy keys were resolved
    if (hadDuplicates || map.size !== arr.length) {
      saveMaster(map);
    }

    return map;
  } catch {
    return new Map();
  }
}

export function saveMaster(map: Map<string, ItemMasterEntry>) {
  try {
    const uniqueList: ItemMasterEntry[] = [];
    const seen = new Set<string>();

    for (const e of map.values()) {
      const cleanName = normalizeItemName(e.name);
      if (!cleanName) continue;
      const key = normalizeItemKey(cleanName);
      if (seen.has(key)) continue;
      seen.add(key);
      uniqueList.push({
        name: cleanName,
        hsn: cleanHSN(e.hsn),
        gstRate: num(e.gstRate),
        mrp: num(e.mrp),
        purchasePrice: num(e.purchasePrice),
        lastSalePrice: num(e.lastSalePrice),
        lastPurchasePrice: num(e.lastPurchasePrice || e.purchasePrice),
        description: typeof e.description === "string" ? e.description.trim() : "",
      });
    }

    localStorage.setItem(LS_KEY, JSON.stringify(uniqueList));
  } catch {}
}

export function clearMaster() {
  try { localStorage.removeItem(LS_KEY); } catch {}
}

interface MergeCandidate {
  name: string;
  hsn?: string;
  gstRate?: number;
  mrp?: number;
  purchasePrice?: number;
  lastPurchasePrice?: number;
  lastSalePrice?: number;
  description?: string;
}

function mergeCandidates(cands: MergeCandidate[]): { added: number; updated: number; total: number; newItems: string[] } {
  const map = loadMaster();
  const newItems: string[] = [];
  const newlyAddedKeys = new Set<string>();
  let updated = 0;

  for (const c of cands) {
    const name = normalizeItemName(c.name);
    if (!name) continue;
    const key = normalizeItemKey(name);
    const existing = map.get(key);

    if (!existing) {
      if (!newlyAddedKeys.has(key)) {
        newlyAddedKeys.add(key);
        const newEntry: ItemMasterEntry = {
          name,
          hsn: cleanHSN(c.hsn ?? ""),
          gstRate: num(c.gstRate),
          mrp: num(c.mrp),
          purchasePrice: num(c.purchasePrice),
          lastSalePrice: num(c.lastSalePrice),
          lastPurchasePrice: num(c.purchasePrice),
          description: typeof c.description === "string" ? c.description.trim() : "",
        };
        map.set(key, newEntry);
        newItems.push(name);
      }
    } else {
      let changed = false;
      if (!existing.hsn && c.hsn) { existing.hsn = cleanHSN(c.hsn); changed = true; }
      else if (existing.hsn && c.hsn && cleanHSN(c.hsn).length > existing.hsn.length) {
        existing.hsn = cleanHSN(c.hsn);
        changed = true;
      }
      if (!existing.gstRate && c.gstRate) { existing.gstRate = num(c.gstRate); changed = true; }
      if (c.mrp && num(c.mrp) > 0 && (!existing.mrp || existing.mrp <= 0)) {
        existing.mrp = num(c.mrp);
        changed = true;
      }
      if (c.purchasePrice && num(c.purchasePrice) > 0) {
        if (!existing.purchasePrice || existing.purchasePrice <= 0) {
          existing.purchasePrice = num(c.purchasePrice);
          changed = true;
        }
        if (!existing.lastPurchasePrice || existing.lastPurchasePrice <= 0) {
          existing.lastPurchasePrice = num(c.purchasePrice);
          changed = true;
        }
      }
      if (c.lastSalePrice && num(c.lastSalePrice) > 0) {
        if (!existing.lastSalePrice || existing.lastSalePrice <= 0) {
          existing.lastSalePrice = num(c.lastSalePrice);
          changed = true;
        }
      }
      if (!existing.description && c.description) {
        existing.description = String(c.description).trim();
        changed = true;
      }
      if (changed) updated++;
    }
  }

  saveMaster(map);
  return { added: newItems.length, updated, total: map.size, newItems };
}

/** Merge from Purchase converted rows. */
export function mergeFromPurchaseRows(rows: Record<string, unknown>[]) {
  const cands: MergeCandidate[] = rows.map((r) => {
    const name = normalizeItemName(r["Item Name Or Alias Name Or SKU*"] ?? r["Item Name"] ?? r["Product Name"] ?? "");
    const hsn = cleanHSN(r["Hsn Code"] ?? r["HSN Code"] ?? r["HSN"] ?? r["CF - Item Custom Field Name 1"] ?? "");
    const gstRate = num(r["GST %"]);
    const mrp = num(r["MRP"]);
    const rate = num(r["Rate per Unit (Without GST)"] ?? r["Purchase Price"] ?? r["Rate per Unit (Without GST)*"]);
    const desc = typeof r["Additional description for Item"] === "string"
      ? String(r["Additional description for Item"]).trim()
      : (typeof r["Description"] === "string" ? String(r["Description"]).trim() : "");
    return {
      name,
      hsn,
      gstRate,
      mrp,
      purchasePrice: rate,
      lastPurchasePrice: rate,
      lastSalePrice: 0,
      description: desc,
    };
  }).filter((c) => c.name);
  return mergeCandidates(cands);
}

/**
 * Directly import a Purchase file (Raw LeverEdge Product Wise Purchase or Converted Purchase Item Excel)
 * into the Item Master Database.
 */
export function importPurchaseFileToMaster(buf: ArrayBuffer): {
  added: number;
  updated: number;
  total: number;
  newItems: string[];
} {
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return { added: 0, updated: 0, total: loadMaster().size, newItems: [] };
  const ws = wb.Sheets[sheetName];
  const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (all.length === 0) return { added: 0, updated: 0, total: loadMaster().size, newItems: [] };

  const norm = (s: unknown) => String(s ?? "").toLowerCase().replace(/[^a-z0-9%/.]/g, " ").replace(/\s+/g, " ").trim();

  // Look for header row
  let headerIdx = -1;
  for (let i = 0; i < Math.min(25, all.length); i++) {
    const r = (all[i] || []).map(norm);
    if (r.includes("item name") || r.includes("product name") || r.includes("item name or alias name or sku")) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx === -1) {
    headerIdx = 0;
  }

  const rawHeaders = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
  const headers = rawHeaders.map(norm);

  const findCol = (...aliases: string[]) => {
    for (const a of aliases) {
      const na = norm(a);
      const idx = headers.indexOf(na);
      if (idx >= 0) return idx;
    }
    for (const a of aliases) {
      const na = norm(a);
      const idx = headers.findIndex((h) => h.includes(na));
      if (idx >= 0) return idx;
    }
    return -1;
  };

  const colItem = findCol("item name or alias name or sku", "item name", "product name", "product description", "item");
  const colHsn = findCol("hsn code", "hsn", "cf - item custom field name 1", "hsn sac", "sac code");
  const colMrp = findCol("mrp", "m.r.p.", "maximum retail price");
  const colRate = findCol("rate per unit without gst", "purchase price", "rate per unit", "invoice price case", "price case");
  const colQty = findCol("purchase qty units", "purchase qty", "quantity", "qty");
  const colNetAmt = findCol("netamt", "net amt", "net amount");
  const colCgst = findCol("cgst %", "cgst%", "cgst percent");
  const colSgst = findCol("sgst %", "sgst%", "sgst percent");
  const colIgst = findCol("igst %", "igst%", "igst percent");
  const colGst = findCol("gst %", "gst%", "total tax %", "tax %");
  const colCode = findCol("item code", "sku", "basepack code", "sku barcode");
  const colPack = findCol("pack size", "packsize");

  const candidates: MergeCandidate[] = [];

  for (let r = headerIdx + 1; r < all.length; r++) {
    const row = all[r];
    if (!row || !row.length) continue;
    if (colItem < 0) continue;
    const rawName = row[colItem];
    const name = normalizeItemName(rawName);
    if (!name || /grand\s*total/i.test(name)) continue;

    const hsnVal = colHsn >= 0 ? cleanHSN(row[colHsn]) : "";
    const mrpVal = colMrp >= 0 ? num(row[colMrp]) : 0;

    let gstVal = 0;
    if (colGst >= 0 && num(row[colGst]) > 0) {
      gstVal = num(row[colGst]);
    } else {
      const cgst = colCgst >= 0 ? num(row[colCgst]) : 0;
      const sgst = colSgst >= 0 ? num(row[colSgst]) : 0;
      const igst = colIgst >= 0 ? num(row[colIgst]) : 0;
      gstVal = Math.round((cgst + sgst + igst) * 100) / 100;
    }

    let rateVal = 0;
    if (colRate >= 0 && num(row[colRate]) > 0) {
      rateVal = num(row[colRate]);
    } else if (colNetAmt >= 0 && colQty >= 0 && num(row[colQty]) > 0) {
      const net = num(row[colNetAmt]);
      const qty = num(row[colQty]);
      const taxable = net / (1 + gstVal / 100);
      rateVal = Math.round((taxable / qty) * 10000) / 10000;
    }

    let desc = "";
    if (colPack >= 0 && row[colPack]) {
      desc = `Pack: ${String(row[colPack]).trim()}`;
    }
    if (colCode >= 0 && row[colCode]) {
      desc = desc ? `${desc} | Code: ${String(row[colCode]).trim()}` : `Code: ${String(row[colCode]).trim()}`;
    }

    candidates.push({
      name,
      hsn: hsnVal,
      gstRate: gstVal,
      mrp: mrpVal,
      purchasePrice: rateVal,
      lastPurchasePrice: rateVal,
      lastSalePrice: 0,
      description: desc,
    });
  }

  return mergeCandidates(candidates);
}

/** Merge from Sale / Sale-Return converted rows (optional fallback if needed). */
export function mergeFromSaleRows(rows: Record<string, unknown>[]) {
  const cands: MergeCandidate[] = rows.map((r) => ({
    name: normalizeItemName(r["Item Name Or Alias Name Or SKU*"] ?? r["Item Name"] ?? r["Product Name"] ?? ""),
    hsn: cleanHSN(r["Hsn Code"] ?? r["HSN Code"] ?? r["CF - Item Custom Field Name 1"] ?? r["HSN"] ?? ""),
    gstRate: num(r["GST %"]),
    mrp: num(r["MRP"]),
    purchasePrice: 0,
    lastSalePrice: num(r["Rate per Unit (Without GST)*"] ?? r["Rate per Unit (Without GST)"]),
    description: typeof r["Description"] === "string" ? String(r["Description"]).trim() : "",
  })).filter((c) => c.name);
  return mergeCandidates(cands);
}

export function getMasterCount(): number { return loadMaster().size; }

/**
 * Parse an uploaded existing Item Master file (Excel or CSV).
 * Extracts existing item names, aliases, and SKUs into a normalized uppercase set.
 */
export function parseExistingItemMaster(buf: ArrayBuffer): {
  existingNames: Set<string>;
  totalFound: number;
} {
  const wb = XLSX.read(buf, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return { existingNames: new Set(), totalFound: 0 };
  const sheet = wb.Sheets[sheetName];
  const rows: unknown[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (rows.length === 0) return { existingNames: new Set(), totalFound: 0 };

  // Find header row (check within first 20 rows)
  let headerIdx = -1;
  let nameColIdx = -1;
  let skuColIdx = -1;
  let aliasColIdx = -1;

  for (let r = 0; r < Math.min(20, rows.length); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const val = String(row[c] ?? "").toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
      if (
        val === "item name" ||
        val === "item" ||
        val === "product" ||
        val === "product name" ||
        val.includes("item name") ||
        val.includes("product name") ||
        val.includes("item name or alias name")
      ) {
        headerIdx = r;
        nameColIdx = c;
        break;
      }
    }
    if (nameColIdx !== -1) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c] ?? "").toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
        if (val.includes("sku") || val.includes("barcode")) skuColIdx = c;
        if (val.includes("alias") && c !== nameColIdx) aliasColIdx = c;
      }
      break;
    }
  }

  // Fallback: If no recognized header text, default to column 0, starting from row 1
  if (nameColIdx === -1) {
    headerIdx = 0;
    nameColIdx = 0;
  }

  const existingNames = new Set<string>();
  let distinctRawCount = 0;
  const seenDistinct = new Set<string>();

  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.length) continue;

    const registerKey = (val: unknown) => {
      const raw = normalizeItemName(val);
      if (!raw) return;
      const key = normalizeItemKey(raw);
      if (!seenDistinct.has(key)) {
        seenDistinct.add(key);
        distinctRawCount++;
      }
      existingNames.add(key);
      existingNames.add(raw.toUpperCase());
      const clean = cleanItemKey(raw);
      if (clean) existingNames.add(clean);
    };

    registerKey(row[nameColIdx]);
    if (skuColIdx !== -1) registerKey(row[skuColIdx]);
    if (aliasColIdx !== -1) registerKey(row[aliasColIdx]);
  }

  return {
    existingNames,
    totalFound: distinctRawCount,
  };
}

/** Check if an item exists in the exclusion set using multiple normalized variations */
export function isItemExcluded(name: string, excludeNames?: Set<string>): boolean {
  if (!excludeNames || excludeNames.size === 0) return false;
  const cleanName = normalizeItemName(name);
  if (!cleanName) return true;

  const key = normalizeItemKey(cleanName);
  if (excludeNames.has(key)) return true;
  if (excludeNames.has(cleanName.toUpperCase())) return true;

  const clean = cleanItemKey(cleanName);
  if (clean && excludeNames.has(clean)) return true;

  return false;
}

export function getNewItemsSummary(excludeNames?: Set<string>): {
  totalInDB: number;
  newCount: number;
  existingCount: number;
  newItems: ItemMasterEntry[];
} {
  const map = loadMaster();
  const seenKeys = new Set<string>();
  const uniqueAll: ItemMasterEntry[] = [];

  for (const e of map.values()) {
    const cleanName = normalizeItemName(e.name);
    if (!cleanName) continue;
    const key = normalizeItemKey(cleanName);
    if (seenKeys.has(key)) continue;
    seenKeys.add(key);
    uniqueAll.push({ ...e, name: cleanName });
  }

  if (!excludeNames || excludeNames.size === 0) {
    return {
      totalInDB: uniqueAll.length,
      newCount: uniqueAll.length,
      existingCount: 0,
      newItems: uniqueAll,
    };
  }

  const newItems = uniqueAll.filter((e) => !isItemExcluded(e.name, excludeNames));

  return {
    totalInDB: uniqueAll.length,
    newCount: newItems.length,
    existingCount: uniqueAll.length - newItems.length,
    newItems,
  };
}

export function buildItemMasterWorkbook(excludeNames?: Set<string>): {
  buffer: ArrayBuffer;
  totalExported: number;
  excludedCount: number;
  exportedItems: ItemMasterEntry[];
} {
  const map = loadMaster();
  const data: unknown[][] = [ITEM_MASTER_HEADERS];
  let excludedCount = 0;
  const exportedItems: ItemMasterEntry[] = [];
  const seenExportKeys = new Set<string>();

  for (const e of map.values()) {
    const cleanName = normalizeItemName(e.name);
    if (!cleanName) continue;
    const key = normalizeItemKey(cleanName);

    // Strict deduplication guard — never export the same item name twice
    if (seenExportKeys.has(key)) continue;

    if (isItemExcluded(cleanName, excludeNames)) {
      excludedCount++;
      continue;
    }

    seenExportKeys.add(key);
    const sanitizedEntry: ItemMasterEntry = {
      ...e,
      name: cleanName,
      hsn: cleanHSN(e.hsn),
      gstRate: num(e.gstRate),
      mrp: num(e.mrp),
      purchasePrice: num(e.purchasePrice),
      lastSalePrice: num(e.lastSalePrice),
      lastPurchasePrice: num(e.lastPurchasePrice || e.purchasePrice),
      description: typeof e.description === "string" ? e.description.trim() : "",
    };
    exportedItems.push(sanitizedEntry);

    const row: Record<string, unknown> = {};
    for (const h of ITEM_MASTER_HEADERS) row[h] = "";

    row["Item Name"] = sanitizedEntry.name;
    row["Alias Name"] = "";
    row["Print Item Name"] = "";
    row["Item Category"] = "GOODS";
    row["Primary Unit Of Measurement"] = "PCS-PIECES";
    row["Primary Unit of Conversion Rate"] = "";
    row["Secondary Unit Of Measurement"] = "";
    row["Secondary Unit of Conversion Rate"] = "";
    row["Decimal Places"] = 2;
    row["SKU/Barcode"] = "";
    row["Description"] = sanitizedEntry.description || "";
    row["Re Order Level"] = "";
    row["Re Order UOM"] = "";
    row["GST Applicable"] = "Yes";
    row["GST Rate"] = sanitizedEntry.gstRate;
    row["HSN Code"] = sanitizedEntry.hsn;
    row["GST Cess Rate"] = "";
    row["RCM Applicable"] = "No";
    row["MRP"] = sanitizedEntry.mrp;
    row["MRP Discount Type"] = "₹";
    row["MRP Discount Value"] = 0;
    row["Selling Price Type"] = "Without GST";
    row["Selling Price"] = 0;
    row["Secondary Unit Selling Price Type"] = "Without GST";
    row["Secondary Unit Selling Price"] = "";
    row["Income Ledger To Be Associated"] = "Sale";
    row["Discount Type"] = "₹";
    row["Discount Value"] = 0;
    row["Purchase Price Type"] = "Without GST";
    row["Purchase Price"] = sanitizedEntry.purchasePrice;
    row["Purchase Discount Type"] = "₹";
    row["Purchase Discount Value"] = 0;
    row["Expense Ledger To Be Associated"] = "Purchase";
    row["Decimal Places For Rate"] = 4;
    row["Last Sale Price"] = sanitizedEntry.lastSalePrice ?? 0;
    row["Last Purchase Price"] = sanitizedEntry.lastPurchasePrice || sanitizedEntry.purchasePrice || 0;
    row["Opening Quantity Unit of Measurement"] = "PCS-PIECES";
    row["Opening Quantity"] = 0;
    row["Opening Rate (Without GST)"] = 0;
    data.push(ITEM_MASTER_HEADERS.map((h) => row[h]));
  }

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Worksheet");
  const buffer = XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;

  return {
    buffer,
    totalExported: exportedItems.length,
    excludedCount,
    exportedItems,
  };
}
