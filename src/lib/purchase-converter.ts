import * as XLSX from "xlsx";

export const PUR_OUTPUT_HEADERS = [
  "Voucher Number*","Voucher Date*","Invoice Number*","Invoice Date*","Supplier Name Or Alias Name*",
  "CF - Party Custom Field Name 1","CF - Party Custom Field Name 2","CF - Party Custom Field Name 3",
  "GSTIN",
  "Billing Address Line 1","Billing Address Line 2","Billing Address Country","Billing Address State","Billing Address City","Billing Address Pincode",
  "Dispatch Name","Dispatch GSTIN","Dispatch Address Line 1","Dispatch Address Line 2","Dispatch Address Country","Dispatch Address State","Dispatch Address City","Dispatch Address Pincode",
  "Broker Name","Brokerage %","Brokerage On Value",
  "Transport Name","Document number","Document Date","Vehicle Number",
  "CF - Trn. Custom Field Name 1","CF - Trn. Custom Field Name 2","CF - Trn. Custom Field Name 3","CF - Trn. Custom Field Name 4","CF - Trn. Custom Field Name 5",
  "Item Name Or Alias Name Or SKU*",
  "CF - Item Custom Field Name 1","CF - Item Custom Field Name 2","CF - Item Custom Field Name 3","CF - Item Custom Field Name 4","CF - Item Custom Field Name 5",
  "Additional description for Item",
  "Ledger Name","Hsn Code","Unit",
  "Quantity*","Free Quantity Unit","Free Quantity",
  "MRP","Rate per Unit (Without GST)",
  "Discount 1 Type","Discount 1","Discount 2 Type","Discount 2",
  "GST %","Classification Nature Type","RCM Applicable","ITC Applicable*",
  "Additional Charges 1 Ledger","Additional Charges 1 Type","Additional Charges 1 Amount (Without GST)","Additional Charges 1 GST %",
  "Additional Charges 2 Ledger","Additional Charges 2 Type","Additional Charges 2 Amount (Without GST)","Additional Charges 2 GST %",
  "Cess","TCS Ledger","TCS Rate","Taxable Value For TCS",
  "Add / less 1 Ledger","Add / less 1 Type","Add / less 1 Amount",
  "Add / less 2 Ledger","Add / less 2 Type","Add / less 2 Amount",
  "Round off Amount","TDS Ledger","TDS Rate","Taxable Value For TDS",
  "Note",
  "Payment 1 Ledger","Payment 1 Date","Payment 1 Amount","Payment 1 Mode","Payment 1 Reference Number",
  "Payment 2 Ledger","Payment 2 Date","Payment 2 Amount","Payment 2 Mode","Payment 2 Reference Number",
];

const PUR_ALIASES: Record<string, string[]> = {
  srNo:          ["sr no", "sr. no", "sr no.", "s no", "sno", "serial no"],
  supplierName:  ["supplier name", "vendor name", "party name"],
  supplierInv:   ["supplier invoice no", "supplier invoice number", "supplier inv no", "sup inv no", "supp invoice no"],
  invoiceNumber: ["invoice number", "invoice no", "inv number", "inv no"],
  invoiceDate:   ["invoice date", "inv date", "bill date"],
  itemName:      ["item name", "product name", "item description"],
  hsn:           ["hsn code", "hsn", "hsn/sac", "hsn sac"],
  upc:           ["upc", "units per case", "u.p.c"],
  mrp:           ["mrp", "m.r.p.", "maximum retail price"],
  purchaseQty:   ["purchase qty units", "purchase qty", "qty", "quantity"],
  freeQty:       ["frms/memo units", "frms memo units", "free qty", "free quantity", "memo units"],
  invoicePrice:  ["invoice(price/case)", "invoice price/case", "price/case", "price per case", "invoice price"],
  cgstPct:       ["cgst %", "cgst%", "cgst percent"],
  sgstPct:       ["sgst %", "sgst%", "sgst percent"],
  igstPct:       ["igst %", "igst%", "igst percent"],
  othDisc:       ["othdisc", "oth disc", "other discount", "oth discount"],
  othAdj:        ["othadj", "oth adj", "other adjustment", "oth adjustment"],
  netAmt:        ["netamt", "net amt", "net amount"],
};

const norm = (s: unknown) =>
  String(s ?? "").toLowerCase().replace(/[^a-z0-9%/.]/g, " ").replace(/\s+/g, " ").trim();

function cleanHSN(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/[^\d]/g, "").slice(0, 6);
}

function findHeaderIndex(headers: string[], key: keyof typeof PUR_ALIASES): number {
  const aliases = PUR_ALIASES[key].map(norm);
  const normHeaders = headers.map(norm);
  for (let i = 0; i < normHeaders.length; i++) if (aliases.includes(normHeaders[i])) return i;
  for (let i = 0; i < normHeaders.length; i++)
    for (const a of aliases) if (normHeaders[i].includes(a) || a.includes(normHeaders[i])) return i;
  return -1;
}

function fmtDate(v: unknown): string {
  if (typeof v === "number") {
    const date = XLSX.SSF.parse_date_code(v);
    if (date) {
      const dd = String(date.d).padStart(2, "0");
      const mm = String(date.m).padStart(2, "0");
      return `${dd}/${mm}/${date.y}`;
    }
  }
  if (v instanceof Date) {
    return `${String(v.getDate()).padStart(2,"0")}/${String(v.getMonth()+1).padStart(2,"0")}/${v.getFullYear()}`;
  }
  if (typeof v === "string") {
    const iso = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
    const dmy = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dmy) {
      const y = dmy[3].length === 2 ? "20" + dmy[3] : dmy[3];
      return `${dmy[1].padStart(2,"0")}/${dmy[2].padStart(2,"0")}/${y}`;
    }
  }
  return String(v ?? "");
}

const num = (v: unknown): number => {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
};

const round2 = (n: number) => Math.round(n * 100) / 100;

export interface PurConvertStats {
  sourceRows: number;
  exportRows: number;
  rejectedInvalid: number;
  errors: { row: number; reason: string }[];
  mappedHeaders: Record<string, string>;
}

export interface PurConvertResult {
  rows: Record<string, unknown>[];
  stats: PurConvertStats;
}

export function convertPurchase(sourceBuf: ArrayBuffer): PurConvertResult {
  const wb = XLSX.read(sourceBuf, { type: "array", cellDates: false, cellFormula: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  // Detect header row: must contain "Item Name" or "Supplier Name"
  let headerIdx = -1;
  for (let i = 0; i < all.length; i++) {
    const r = (all[i] || []).map(norm);
    if (r.includes("item name") || r.includes("supplier name")) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0)
    throw new Error("Header row not found (expected 'Item Name' / 'Supplier Name').");

  const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
  const rawData = all.slice(headerIdx + 1);

  // Filter: skip empty rows, Grand Total rows, and last summary row
  const dataRows = rawData.filter((r) => {
    const first = String(r[0] ?? "").trim();
    if (!first || first === "" || /grand\s*total/i.test(first)) return false;
    if (r.every((c) => c === null || c === "")) return false;
    return true;
  });

  const idx = {
    srNo:         findHeaderIndex(headers, "srNo"),
    supplierName: findHeaderIndex(headers, "supplierName"),
    supplierInv:  findHeaderIndex(headers, "supplierInv"),
    invoiceNumber: findHeaderIndex(headers, "invoiceNumber"),
    invoiceDate:  findHeaderIndex(headers, "invoiceDate"),
    itemName:     findHeaderIndex(headers, "itemName"),
    hsn:          findHeaderIndex(headers, "hsn"),
    upc:          findHeaderIndex(headers, "upc"),
    mrp:          findHeaderIndex(headers, "mrp"),
    purchaseQty:  findHeaderIndex(headers, "purchaseQty"),
    freeQty:      findHeaderIndex(headers, "freeQty"),
    invoicePrice: findHeaderIndex(headers, "invoicePrice"),
    cgstPct:      findHeaderIndex(headers, "cgstPct"),
    sgstPct:      findHeaderIndex(headers, "sgstPct"),
    igstPct:      findHeaderIndex(headers, "igstPct"),
    othDisc:      findHeaderIndex(headers, "othDisc"),
    othAdj:       findHeaderIndex(headers, "othAdj"),
    netAmt:       findHeaderIndex(headers, "netAmt"),
  };

  const mappedHeaders: Record<string, string> = {};
  for (const [k, i] of Object.entries(idx)) if (i >= 0) mappedHeaders[k] = headers[i];

  const required: (keyof typeof idx)[] = ["supplierName", "supplierInv", "invoiceDate", "itemName", "purchaseQty"];
  const missing = required.filter((k) => idx[k] < 0);
  if (missing.length)
    throw new Error("Missing columns: " + missing.map((k) => `${k} (tried: ${PUR_ALIASES[k].join(", ")})`).join("; "));

  const errors: { row: number; reason: string }[] = [];
  const out: Record<string, unknown>[] = [];
  let voucherSeq = 0;

  dataRows.forEach((r, i) => {
    const rowNum = headerIdx + 2 + i;
    const supplier = r[idx.supplierName];
    const itemName = r[idx.itemName];
    const supplierInv = r[idx.supplierInv];

    if (!supplier || !itemName || !supplierInv) {
      errors.push({ row: rowNum, reason: "Missing Supplier/Item/Invoice" });
      return;
    }

    const qty = num(r[idx.purchaseQty]);
    if (qty === 0) {
      errors.push({ row: rowNum, reason: "Quantity = 0" });
      return;
    }

    voucherSeq++;
    const invDate    = fmtDate(r[idx.invoiceDate]);
    const cgstPct    = idx.cgstPct >= 0 ? num(r[idx.cgstPct]) : 0;
    const sgstPct    = idx.sgstPct >= 0 ? num(r[idx.sgstPct]) : 0;
    const igstPct    = idx.igstPct >= 0 ? num(r[idx.igstPct]) : 0;
    const totalGST   = round2(cgstPct + sgstPct + igstPct);
    const othDisc    = idx.othDisc >= 0 ? num(r[idx.othDisc]) : 0;
    const othAdj     = idx.othAdj  >= 0 ? num(r[idx.othAdj])  : 0;
    const discount   = round2(othDisc + othAdj);
    const freeQtyVal  = idx.freeQty >= 0 ? num(r[idx.freeQty]) : 0;
    const mrp         = idx.mrp     >= 0 ? round2(num(r[idx.mrp])) : 0;
    const netAmt      = idx.netAmt  >= 0 ? num(r[idx.netAmt]) : 0;
    // Taxable value = NetAmt ÷ (1 + GST%) — backs out GST from net amount
    const gstDivisor  = 1 + totalGST / 100;
    const taxableVal  = round2(netAmt / gstDivisor);
    // Rate per unit = Taxable value ÷ Qty (4 decimal places for precision)
    const ratePerUnit = qty > 0 ? Math.round((taxableVal / qty) * 10000) / 10000 : 0;

    let classification = "Intrastate Purchase Exempt";
    if (cgstPct > 0 && sgstPct > 0) classification = "Intrastate Purchase Taxable";
    else if (igstPct > 0) classification = "Interstate Purchase Taxable";

    const row: Record<string, unknown> = {};
    for (const h of PUR_OUTPUT_HEADERS) row[h] = "";

    const invNumSrc = idx.invoiceNumber >= 0 ? r[idx.invoiceNumber] : null;
    row["Voucher Number*"]                 = invNumSrc ? String(invNumSrc) : `PURCHASE-${String(voucherSeq).padStart(5, "0")}`;
    row["Voucher Date*"]                   = invDate;
    row["Invoice Number*"]                 = String(supplierInv);
    row["Invoice Date*"]                   = invDate;
    row["Supplier Name Or Alias Name*"]    = "HINDUSTAN UNILEVER";
    row["Billing Address Country"]         = "India";
    row["Dispatch Address Country"]        = "India";
    row["Item Name Or Alias Name Or SKU*"] = String(itemName);
    if (idx.hsn >= 0) {
      const hsnVal = cleanHSN(r[idx.hsn]);
      row["Hsn Code"]                      = hsnVal;
      row["CF - Item Custom Field Name 1"] = hsnVal;
    }
    row["Ledger Name"]                     = "Purchase";
    row["Unit"]                            = "PCS-PIECES";
    row["Quantity*"]                       = qty;
    row["Free Quantity"]                   = freeQtyVal;
    row["MRP"]                             = mrp;
    row["Rate per Unit (Without GST)"]     = ratePerUnit;
    row["Discount 1 Type"]                 = "T";
    row["Discount 1"]                      = discount;
    row["Discount 2 Type"]                 = "";
    row["Discount 2"]                      = "";
    row["GST %"]                           = totalGST;
    row["Classification Nature Type"]      = classification;
    row["RCM Applicable"]                  = "NO";
    row["ITC Applicable*"]                 = totalGST > 0 ? "YES" : "NO";
    row["Taxable Value For TCS"]           = taxableVal;
    row["Taxable Value For TDS"]           = "";
    row["Payment 1 Date"]                  = "";
    row["Payment 1 Amount"]                = "";
    row["Payment 1 Mode"]                  = "";

    out.push(row);
  });

  return {
    rows: out,
    stats: {
      sourceRows: dataRows.length,
      exportRows: out.length,
      rejectedInvalid: errors.length,
      errors,
      mappedHeaders,
    },
  };
}

const PUR_FMT_2DP = ["MRP","Discount 1","Quantity*","Free Quantity","GST %","Taxable Value For TCS","Taxable Value For TDS","Payment 1 Amount"];

function applyNumFormats(ws: XLSX.WorkSheet, numDataRows: number) {
  for (const col of PUR_FMT_2DP) {
    const c = PUR_OUTPUT_HEADERS.indexOf(col);
    if (c < 0) continue;
    for (let row = 1; row <= numDataRows; row++) {
      const addr = XLSX.utils.encode_cell({ r: row, c });
      const cell = ws[addr];
      if (cell && cell.v !== "" && cell.v !== null && cell.v !== undefined) {
        cell.t = "n";
        cell.z = "0.00";
      }
    }
  }
  // Rate per Unit — 4 decimal places, General number type
  const rateCol = PUR_OUTPUT_HEADERS.indexOf("Rate per Unit (Without GST)");
  if (rateCol >= 0) {
    for (let row = 1; row <= numDataRows; row++) {
      const addr = XLSX.utils.encode_cell({ r: row, c: rateCol });
      const cell = ws[addr];
      if (cell && cell.v !== "" && cell.v !== null && cell.v !== undefined) {
        cell.t = "n";
        cell.z = "0.0000";
      }
    }
  }
}

export function buildPurchaseWorkbook(rows: Record<string, unknown>[]): ArrayBuffer {
  const data = [PUR_OUTPUT_HEADERS, ...rows.map((r) => PUR_OUTPUT_HEADERS.map((h) => r[h] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  applyNumFormats(ws, rows.length);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Purchase Item");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}
