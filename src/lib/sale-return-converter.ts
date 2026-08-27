import * as XLSX from "xlsx";
import { type RegisterDiscount, normBillNo, lookupDiscount, mergeSameItemRowsInBill } from "@/lib/sale-converter";
import { loadMaster, normalizeItemKey, cleanItemKey } from "@/lib/item-master";

export const SR_OUTPUT_HEADERS = [
  "Credit Note Number*",
  "Credit Note Date*",
  "Customer Name Or Alias Name*",
  "CF - Party Custom Field Name 1",
  "CF - Party Custom Field Name 2",
  "CF - Party Custom Field Name 3",
  "Original Invoice Number",
  "Original Invoice Date",
  "Country Code For Mobile Number",
  "Mobile Number",
  "GSTIN",
  "Billing Address Line 1",
  "Billing Address Line 2",
  "Billing Address Country",
  "Billing Address State",
  "Billing Address City",
  "Billing Address Pincode",
  "Shipping Name",
  "Shipping GSTIN",
  "Shipping Address Line 1",
  "Shipping Address Line 2",
  "Shipping Address Country",
  "Shipping Address State",
  "Shipping Address City",
  "Shipping Address Pincode",
  "Broker Name",
  "Brokerage %",
  "Brokerage On Value",
  "Transport Name",
  "Document number",
  "Document Date",
  "Vehicle Number",
  "PO No.",
  "PO Date",
  "Credit Period",
  "Credit Period Type",
  "CF - Trn. Custom Field Name 1",
  "CF - Trn. Custom Field Name 2",
  "CF - Trn. Custom Field Name 3",
  "CF - Trn. Custom Field Name 4",
  "CF - Trn. Custom Field Name 5",
  "Item Name Or Alias Name Or SKU*",
  "CF - Item Custom Field Name 1",
  "CF - Item Custom Field Name 2",
  "CF - Item Custom Field Name 3",
  "CF - Item Custom Field Name 4",
  "CF - Item Custom Field Name 5",
  "Additional description for Item",
  "Ledger Name",
  "Hsn Code",
  "Unit",
  "Quantity*",
  "Free Quantity Unit",
  "Free Quantity",
  "MRP",
  "Rate per Unit (Without GST)*",
  "Discount 1 Type",
  "Discount 1",
  "Discount 2 Type",
  "Discount 2",
  "GST %",
  "Classification Nature Type",
  "RCM Applicable",
  "Additional Charges 1 Ledger",
  "Additional Charges 1 Type",
  "Additional Charges 1 Amount (Without GST)",
  "Additional Charges 1 GST %",
  "Additional Charges 2 Ledger",
  "Additional Charges 2 Type",
  "Additional Charges 2 Amount (Without GST)",
  "Additional Charges 2 GST %",
  "Cess",
  "TCS Ledger",
  "TCS Rate",
  "Taxable Value For TCS",
  "Add / less 1 Ledger",
  "Add / less 1 Type",
  "Add / less 1 Amount",
  "Add / less 2 Ledger",
  "Add / less 2 Type",
  "Add / less 2 Amount",
  "Round off Amount",
  "TDS Ledger",
  "TDS Rate",
  "Taxable Value For TDS",
  "Note",
  "Terms & Conditions",
  "Payment 1 Ledger",
  "Payment 1 Date",
  "Payment 1 Amount",
  "Payment 1 Mode",
  "Payment 1 Reference Number",
  "Payment 2 Ledger",
  "Payment 2 Date",
  "Payment 2 Amount",
  "Payment 2 Mode",
  "Payment 2 Reference Number",
];

const SR_ALIASES: Record<string, string[]> = {
  srNumber: [
    "sr number", "sales return no", "sales return number", "sales return",
    "credit note number", "credit note no", "credit note no.", "credit note",
    "cn number", "cn no", "return number", "return no", "voucher no", "voucher number"
  ],
  srDate: [
    "sales return date", "return date", "sr date", "credit note date", "cn date", "doc date", "voucher date"
  ],
  billNumber: [
    "bill number", "bill no", "bill no.", "orig bill no", "orig bill no.", "orig bill number", "original bill no", "original bill number",
    "original invoice no", "orig invoice no", "orig invoice number", "original invoice number",
    "bill ref no", "bill ref. no.", "billrefno", "ref bill no", "against bill no", "against inv no",
    "invoice number", "invoice no", "invoice no.", "ref no", "reference no"
  ],
  billDate: [
    "bill date", "orig bill date", "original bill date", "orig invoice date", "original invoice date", "invoice date"
  ],
  party: [
    "party name", "customer name", "account name", "ledger name", "party/customer", "buyer name", "party", "customer"
  ],
  gstin: [
    "party gstin number", "party gstin", "party gst no", "gstin number", "gstin", "gst number", "gst no"
  ],
  salesperson: [
    "salesperson name", "sales person name", "sales person", "salesperson",
    "salesman name", "sales man name", "salesman", "sales man",
    "broker name", "broker", "rep name", "sales representative", "agent name", "agent"
  ],
  product: [
    "product name", "product description", "item description", "item name",
    "description", "product", "item", "sku", "particulars"
  ],
  hsn: [
    "hsn code", "hsn", "hsn/sac", "hsn sac", "sac", "sac code"
  ],
  mrp: [
    "product mrp", "mrp", "m.r.p.", "unit mrp", "item mrp"
  ],
  basicRate: [
    "basic rate", "rate per unit without gst", "rate per unit", "rate without gst", "item rate", "sales rate", "price", "unit rate", "rate"
  ],
  qty: [
    "qty", "quantity", "return qty", "sales return qty", "ret qty", "billed qty", "units"
  ],
  freeQty: [
    "free qty", "free quantity", "free units", "free"
  ],
  schemeDisc: [
    "scheme discount", "scheme disc", "sch disc", "sch discount"
  ],
  cashDisc: [
    "cash discount", "cash disc", "cd disc"
  ],
  totDisc: [
    "tot discount", "tot disc", "total discount", "total disc", "discount", "disc"
  ],
  taxableValue: [
    "taxable value", "taxable amt", "taxable amount", "taxable"
  ],
  taxValue: [
    "tax value", "tax amount", "tax amt", "total tax", "tax"
  ],
  gstpct: [
    "total tax %", "gst %", "gst%", "tax %", "gst percent", "rate of tax"
  ],
  netsales: [
    "net credit value after discount reversal", "sales return value", "net credit value", "net sales", "net amount", "net value", "net return", "net"
  ],
  grosssales: [
    "gross amount", "gross sales", "gross value", "gross return", "gross"
  ],
  vehicle: [
    "vehicle", "vehicle name", "vehicle no", "vehicle number", "transport"
  ],
};

const norm = (s: unknown) =>
  String(s ?? "").toLowerCase().replace(/[^a-z0-9%]/g, " ").replace(/\s+/g, " ").trim();

function findHeaderIndex(headers: string[], key: keyof typeof SR_ALIASES, excludeCol = -1): number {
  const aliases = SR_ALIASES[key].map(norm);
  const normHeaders = headers.map(norm);
  // Pass 1: exact match
  for (let i = 0; i < normHeaders.length; i++) {
    if (i === excludeCol) continue;
    if (aliases.includes(normHeaders[i])) return i;
  }
  // Pass 2: contains match
  for (let i = 0; i < normHeaders.length; i++) {
    if (i === excludeCol) continue;
    for (const a of aliases) {
      if (normHeaders[i].includes(a) || (a.length >= 4 && a.includes(normHeaders[i]))) return i;
    }
  }
  return -1;
}

const GST_SLABS = [0, 5, 12, 18, 28];
function cleanHSN(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/[^\d]/g, "").slice(0, 8);
}

function cleanGSTIN(v: unknown): string {
  const s = String(v ?? "").toUpperCase().replace(/\s+/g, "");
  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]Z[0-9A-Z]$/.test(s) ? s : "";
}

function snapGST(taxVal: number, taxableVal: number): number {
  if (taxableVal <= 0) return 0;
  const pct = (taxVal / taxableVal) * 100;
  return GST_SLABS.reduce((prev, cur) =>
    Math.abs(cur - pct) < Math.abs(prev - pct) ? cur : prev
  );
}

function fmtDate(v: unknown): string {
  if (!v) return "";
  if (v instanceof Date) {
    const dd = String(v.getDate()).padStart(2, "0");
    const mm = String(v.getMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${v.getFullYear()}`;
  }
  if (typeof v === "number") {
    const date = XLSX.SSF.parse_date_code(v);
    if (date) {
      const dd = String(date.d).padStart(2, "0");
      const mm = String(date.m).padStart(2, "0");
      return `${dd}/${mm}/${date.y}`;
    }
  }
  if (typeof v === "string") {
    const iso = v.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
    const dmy = v.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dmy) {
      const y = dmy[3].length === 2 ? "20" + dmy[3] : dmy[3];
      return `${dmy[1].padStart(2, "0")}/${dmy[2].padStart(2, "0")}/${y}`;
    }
  }
  return String(v ?? "").trim();
}

const num = (v: unknown): number => {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
};

const round2 = (n: number) => Math.round(n * 100) / 100;

export interface SRConvertStats {
  sourceRows: number;
  exportRows: number;
  rejectedInvalid: number;
  errors: { row: number; reason: string }[];
  mappedHeaders: Record<string, string>;
  discountBillsMatched: number;
  discountBills: {
    bill: string;
    ushop: number;
    shikhar: number;
    btpr?: number;
    cash?: number;
    adjustments?: number;
    totalAddLess1?: number;
    roundOff?: number;
  }[];
}

export interface SRConvertResult {
  rows: Record<string, unknown>[];
  stats: SRConvertStats;
}

export function convertSaleReturn(
  sourceBuf: ArrayBuffer,
  discMap?: Map<string, RegisterDiscount> | null,
  extraHsn?: Map<string, string> | null,
): SRConvertResult {
  const wb = XLSX.read(sourceBuf, { type: "array", cellDates: false, cellFormula: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const CORE_SIGNALS = [
    "sr number", "sales return no", "sales return number", "sales return date", "return date", "credit note",
    "product name", "item name", "product description", "product code", "product mrp",
    "basic rate", "taxable value", "sales return value", "gross amount",
    "qty", "quantity", "return qty",
    "bill number", "orig bill no", "original bill no", "invoice number"
  ];

  let headerIdx = -1;
  let maxScore = 0;

  for (let i = 0; i < Math.min(all.length, 60); i++) {
    const row = all[i] || [];
    const normRow = row.map(norm);
    let score = 0;
    for (const cell of normRow) {
      if (!cell) continue;
      for (const sig of CORE_SIGNALS) {
        if (cell === sig || cell.includes(sig)) {
          score++;
          break;
        }
      }
    }
    if (score >= 3 && score > maxScore) {
      maxScore = score;
      headerIdx = i;
    }
  }

  if (headerIdx < 0 && all.length > 0) {
    let maxCols = 0;
    for (let i = 0; i < Math.min(all.length, 30); i++) {
      const count = (all[i] || []).filter((c) => typeof c === "string" && c.trim().length > 0).length;
      if (count > maxCols) {
        maxCols = count;
        headerIdx = i;
      }
    }
  }

  if (headerIdx < 0)
    throw new Error("Header row not found in Sale Return (expected 'SR Number' / 'Sales Return Date' / 'Bill No' / 'Product Name').");

  const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
  const dataRows = all.slice(headerIdx + 1).filter((r) => r.some((c) => c !== null && c !== ""));

  const srNumberCol = findHeaderIndex(headers, "srNumber");
  const billNumberCol = findHeaderIndex(headers, "billNumber", srNumberCol);

  const idx = {
    srNumber:     srNumberCol,
    srDate:       findHeaderIndex(headers, "srDate"),
    billNumber:   billNumberCol,
    billDate:     findHeaderIndex(headers, "billDate"),
    party:        findHeaderIndex(headers, "party"),
    gstin:        findHeaderIndex(headers, "gstin"),
    salesperson:  findHeaderIndex(headers, "salesperson"),
    product:      findHeaderIndex(headers, "product"),
    hsn:          findHeaderIndex(headers, "hsn"),
    mrp:          findHeaderIndex(headers, "mrp"),
    basicRate:    findHeaderIndex(headers, "basicRate"),
    qty:          findHeaderIndex(headers, "qty"),
    freeQty:      findHeaderIndex(headers, "freeQty"),
    schemeDisc:   findHeaderIndex(headers, "schemeDisc"),
    cashDisc:     findHeaderIndex(headers, "cashDisc"),
    totDisc:      findHeaderIndex(headers, "totDisc"),
    taxableValue: findHeaderIndex(headers, "taxableValue"),
    taxValue:     findHeaderIndex(headers, "taxValue"),
    gstpct:       findHeaderIndex(headers, "gstpct"),
    netsales:     findHeaderIndex(headers, "netsales"),
    grosssales:   findHeaderIndex(headers, "grosssales"),
    vehicle:      findHeaderIndex(headers, "vehicle"),
  };

  const mappedHeaders: Record<string, string> = {};
  for (const [k, i] of Object.entries(idx)) if (i >= 0) mappedHeaders[k] = headers[i];

  const required: (keyof typeof idx)[] = ["srNumber", "srDate", "party", "product", "qty"];
  const missing = required.filter((k) => idx[k] < 0);
  if (missing.length)
    throw new Error("Missing columns in Sale Return: " + missing.map((k) => `${k} (tried: ${SR_ALIASES[k].join(", ")})`).join("; "));

  // Fast index Item Master for HSN and metadata fallback
  const masterMap = loadMaster();
  const cleanMasterMap = new Map<string, { hsn?: string }>();
  for (const entry of masterMap.values()) {
    const cKey = cleanItemKey(entry.name);
    if (cKey && !cleanMasterMap.has(cKey)) {
      cleanMasterMap.set(cKey, entry);
    }
  }

  const errors: { row: number; reason: string }[] = [];
  const out: Record<string, unknown>[] = [];

  let lastParty = "";
  let lastDate: unknown = null;
  let lastSalesperson = "";

  dataRows.forEach((r, i) => {
    const rowNum = headerIdx + 2 + i;
    const srNum = r[idx.srNumber];
    const party = (idx.party >= 0 && r[idx.party]) ? String(r[idx.party]).trim() : lastParty;
    if (party) lastParty = party;

    const product = idx.product >= 0 && r[idx.product] ? String(r[idx.product]).trim() : "";

    if (!srNum || !party || !product) {
      errors.push({ row: rowNum, reason: "Missing SR Number/Party/Product" });
      return;
    }

    const rawQty = idx.qty >= 0 ? num(r[idx.qty]) : 0;
    const rawFreeQty = idx.freeQty >= 0 ? num(r[idx.freeQty]) : 0;
    const absQty = Math.abs(rawQty) || (rawFreeQty > 0 ? Math.abs(rawFreeQty) : 0);

    if (absQty === 0) {
      errors.push({ row: rowNum, reason: "Quantity = 0" });
      return;
    }

    const srDateVal = (idx.srDate >= 0 && r[idx.srDate]) ? r[idx.srDate] : lastDate;
    if (srDateVal) lastDate = srDateVal;

    const salesperson = (idx.salesperson >= 0 && r[idx.salesperson]) ? String(r[idx.salesperson]).trim() : lastSalesperson;
    if (salesperson) lastSalesperson = salesperson;

    const taxable    = idx.taxableValue >= 0 ? Math.abs(num(r[idx.taxableValue])) : 0;
    const taxVal     = idx.taxValue    >= 0 ? Math.abs(num(r[idx.taxValue]))     : 0;
    let gstPct       = idx.gstpct >= 0 ? num(r[idx.gstpct]) : 0;
    if (gstPct <= 0 && taxVal > 0 && taxable > 0) {
      gstPct = snapGST(taxVal, taxable);
    }

    const totalDisc  = round2(
      (idx.schemeDisc >= 0 ? Math.abs(num(r[idx.schemeDisc])) : 0) +
      (idx.cashDisc   >= 0 ? Math.abs(num(r[idx.cashDisc]))   : 0) +
      (idx.totDisc    >= 0 ? Math.abs(num(r[idx.totDisc]))     : 0)
    );

    const row: Record<string, unknown> = {};
    for (const h of SR_OUTPUT_HEADERS) row[h] = "";

    const srNumStr = String(srNum).trim();
    let billNoRaw = idx.billNumber >= 0 ? r[idx.billNumber] : null;
    let billNoStr = billNoRaw != null ? String(billNoRaw).trim() : "";

    if (!billNoStr || billNoStr.toUpperCase().startsWith("SRT")) {
      for (let c = 0; c < r.length; c++) {
        if (c === idx.srNumber) continue;
        const cellVal = normBillNo(r[c]);
        if (!cellVal) continue;
        if (cellVal.toUpperCase().startsWith("SRT")) continue;
        if (/^\d{3,8}$/.test(cellVal) || /^GST[\w\d\-]+$/i.test(cellVal)) {
          billNoStr = cellVal;
          break;
        }
      }
    }

    row["Credit Note Number*"]             = srNumStr;
    row["Credit Note Date*"]               = fmtDate(srDateVal);
    row["Customer Name Or Alias Name*"]    = party;
    row["Original Invoice Number"]         = billNoStr;
    row["Original Invoice Date"]           = idx.billDate >= 0 ? fmtDate(r[idx.billDate]) : "";
    row["GSTIN"]                           = idx.gstin >= 0 ? cleanGSTIN(r[idx.gstin]) : "";
    row["Billing Address Country"]         = "INDIA";
    row["Billing Address State"]           = "GUJARAT";
    row["Billing Address City"]            = "SURAT";
    row["Shipping Address Country"]        = "INDIA";
    row["Shipping Address State"]          = "GUJARAT";
    row["Shipping Address City"]           = "SURAT";

    row["Broker Name"]                     = salesperson;
    row["Brokerage %"]                     = "";
    row["Brokerage On Value"]              = "";

    row["Vehicle Number"]                  = (idx.vehicle >= 0 ? String(r[idx.vehicle] ?? "").trim() : "") || "KATARGAM";
    row["Item Name Or Alias Name Or SKU*"] = product;

    // HSN Code resolution: uploaded file column → Item Master (local) → cloud/sale-data map
    let hsnCode = idx.hsn >= 0 ? cleanHSN(r[idx.hsn]) : "";
    if (!hsnCode && product) {
      const key = normalizeItemKey(product);
      const cKey = cleanItemKey(product);
      const masterEntry = masterMap.get(key) || cleanMasterMap.get(cKey);
      if (masterEntry?.hsn) hsnCode = cleanHSN(masterEntry.hsn);
      if (!hsnCode && extraHsn) {
        hsnCode = cleanHSN(extraHsn.get(key) || extraHsn.get(cKey) || "");
      }
    }
    row["Hsn Code"]                        = hsnCode;


    row["Ledger Name"]                     = "Sale";
    row["Unit"]                            = "PCS-PIECES";
    row["Quantity*"]                       = absQty;
    row["Free Quantity Unit"]              = "";
    row["Free Quantity"]                   = "";
    row["MRP"]                             = idx.mrp >= 0 ? round2(Math.abs(num(r[idx.mrp]))) : "";

    let ratePerUnit = idx.basicRate >= 0 ? round2(Math.abs(num(r[idx.basicRate]))) : 0;
    if (ratePerUnit <= 0 && absQty > 0) {
      const netVal   = idx.netsales >= 0 ? Math.abs(num(r[idx.netsales])) : 0;
      const grossVal = idx.grosssales >= 0 ? Math.abs(num(r[idx.grosssales])) : 0;
      if (taxable > 0) ratePerUnit = Math.round((taxable / absQty) * 10000) / 10000;
      else if (netVal > 0) ratePerUnit = Math.round((netVal / absQty) * 10000) / 10000;
      else if (grossVal > 0) ratePerUnit = Math.round((grossVal / absQty) * 10000) / 10000;
      else if (idx.mrp >= 0 && num(r[idx.mrp]) > 0) ratePerUnit = Math.abs(num(r[idx.mrp]));
    }
    if (ratePerUnit <= 0 && idx.mrp >= 0 && num(r[idx.mrp]) > 0) ratePerUnit = Math.abs(num(r[idx.mrp]));
    if (ratePerUnit <= 0) ratePerUnit = 0.01;

    row["Rate per Unit (Without GST)*"]    = ratePerUnit;
    if (totalDisc > 0) {
      row["Discount 1 Type"]               = "₹";
      row["Discount 1"]                    = totalDisc;
    } else {
      row["Discount 1 Type"]               = "";
      row["Discount 1"]                    = "";
    }
    row["Discount 2 Type"]                 = "";
    row["Discount 2"]                      = "";
    row["GST %"]                           = gstPct || "";
    row["Classification Nature Type"]      = "NA";
    row["RCM Applicable"]                  = "Intrastate Sales Taxable";

    out.push(row);
  });

  // Aggregate duplicate rows with the same item name in the same credit note (Weighted Average Base Rate)
  const mergedOut = mergeSameItemRowsInBill(out, "Credit Note Number*", "Item Name Or Alias Name Or SKU*");

  // Natural alphabetical & numerical sequential sort by Credit Note Number
  const uniqueBills = Array.from(new Set(mergedOut.map((r) => String(r["Credit Note Number*"] ?? ""))));
  uniqueBills.sort((a, b) => {
    const numA = parseInt(a.replace(/[^0-9]/g, ""), 10);
    const numB = parseInt(b.replace(/[^0-9]/g, ""), 10);
    if (!isNaN(numA) && !isNaN(numB) && numA !== numB) return numA - numB;
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
  });

  const billOrderMap = new Map<string, number>();
  uniqueBills.forEach((b, idx) => billOrderMap.set(b, idx));

  mergedOut.sort((a, b) => {
    const invA = String(a["Credit Note Number*"] ?? "");
    const invB = String(b["Credit Note Number*"] ?? "");
    const orderA = billOrderMap.get(invA) ?? 0;
    const orderB = billOrderMap.get(invB) ?? 0;
    return orderA - orderB;
  });

  // Group / Sort and place Sales Register discount ONLY on the FIRST row of each unique Credit Note
  const srDiscountPlaced = new Set<string>();
  let discountBillsMatched = 0;
  const discountBills: {
    bill: string;
    ushop: number;
    shikhar: number;
    btpr?: number;
    cash?: number;
    adjustments?: number;
    totalAddLess1?: number;
    roundOff?: number;
  }[] = [];

  if (discMap) {
    for (const row of mergedOut) {
      const srNumStr = String(row["Credit Note Number*"] ?? "").trim();
      const origBill = String(row["Original Invoice Number"] ?? "").trim();
      if (!srNumStr || srDiscountPlaced.has(srNumStr)) continue;

      const disc = origBill ? lookupDiscount(discMap, origBill) : lookupDiscount(discMap, srNumStr);
      if (disc) {
        const addLess1 = disc.totalAddLess1 !== undefined
          ? disc.totalAddLess1
          : round2((disc.ushop || 0) + (disc.btpr || 0));

        if (addLess1 > 0) {
          row["Add / less 1 Ledger"] = "USHOP DIS";
          row["Add / less 1 Type"]   = "₹";
          row["Add / less 1 Amount"] = -round2(addLess1);
        }
        if (disc.shikhar > 0) {
          row["Add / less 2 Ledger"] = "SHIKHAR DIS";
          row["Add / less 2 Type"]   = "₹";
          row["Add / less 2 Amount"] = -round2(disc.shikhar);
        }
        if (disc.roundOff !== undefined && disc.roundOff !== 0) {
          row["Round off Amount"] = round2(disc.roundOff);
        }

        discountBillsMatched++;
        discountBills.push({
          bill: origBill || srNumStr,
          ushop: disc.ushop,
          shikhar: round2(disc.shikhar),
          btpr: disc.btpr,
          cash: disc.cash ?? 0,
          adjustments: disc.adjustments ?? 0,
          totalAddLess1: addLess1,
          roundOff: disc.roundOff ?? 0,
        });
      }
      srDiscountPlaced.add(srNumStr);
    }
  }

  return {
    rows: mergedOut,
    stats: {
      sourceRows: dataRows.length,
      exportRows: out.length,
      rejectedInvalid: errors.length,
      errors,
      mappedHeaders,
      discountBillsMatched,
      discountBills,
    },
  };
}

const FMT_2DP_COLS = ["MRP", "Discount 1", "Discount 2", "Quantity*", "Free Quantity", "GST %", "Add / less 1 Amount", "Add / less 2 Amount", "Round off Amount"];

function applyNumFormats(ws: XLSX.WorkSheet, numDataRows: number) {
  for (const col of FMT_2DP_COLS) {
    const c = SR_OUTPUT_HEADERS.indexOf(col);
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
  const rateCol = SR_OUTPUT_HEADERS.indexOf("Rate per Unit (Without GST)*");
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

export function buildSRWorkbook(rows: Record<string, unknown>[]): ArrayBuffer {
  const data = [SR_OUTPUT_HEADERS, ...rows.map((r) => SR_OUTPUT_HEADERS.map((h) => r[h] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  applyNumFormats(ws, rows.length);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sale Return Accounting");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

export function buildSRCSV(rows: Record<string, unknown>[]): string {
  const data = [SR_OUTPUT_HEADERS, ...rows.map((r) => SR_OUTPUT_HEADERS.map((h) => r[h] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  return XLSX.utils.sheet_to_csv(ws);
}

const SR_SPLIT_MAX = 4000;

function splitSRAtBillBoundary(
  rows: Record<string, unknown>[],
  targetSize = SR_SPLIT_MAX
): Record<string, unknown>[][] {
  if (rows.length <= targetSize) return [rows];
  const chunks: Record<string, unknown>[][] = [];
  let remaining = rows;

  const getBill = (row: Record<string, unknown>) => String(row["Credit Note Number*"] ?? "").trim();

  while (remaining.length > targetSize) {
    let splitIdx = targetSize;

    // Check if targetSize falls inside a return bill
    if (getBill(remaining[splitIdx]) === getBill(remaining[splitIdx - 1])) {
      const currentBill = getBill(remaining[splitIdx]);

      // Scan backwards to find where currentBill starts
      let backwardIdx = splitIdx;
      while (backwardIdx > 0 && getBill(remaining[backwardIdx - 1]) === currentBill) {
        backwardIdx--;
      }

      // Scan forwards to find where currentBill ends
      let forwardIdx = splitIdx;
      while (forwardIdx < remaining.length && getBill(remaining[forwardIdx]) === currentBill) {
        forwardIdx++;
      }

      const distBackward = splitIdx - backwardIdx;
      const distForward = forwardIdx - splitIdx;

      // Choose boundary closer to ~4000 rows
      if (backwardIdx > 0 && (forwardIdx >= remaining.length || distBackward <= distForward)) {
        splitIdx = backwardIdx;
      } else if (forwardIdx < remaining.length) {
        splitIdx = forwardIdx;
      } else if (backwardIdx > 0) {
        splitIdx = backwardIdx;
      }
    }

    if (splitIdx <= 0 || splitIdx >= remaining.length) {
      chunks.push(remaining);
      remaining = [];
      break;
    }

    chunks.push(remaining.slice(0, splitIdx));
    remaining = remaining.slice(splitIdx);
  }

  if (remaining.length > 0) chunks.push(remaining);
  return chunks;
}

export function buildSRWorkbooks(rows: Record<string, unknown>[]): ArrayBuffer[] {
  const chunks = splitSRAtBillBoundary(rows);
  return chunks.map((chunk) => buildSRWorkbook(chunk));
}

/**
 * Preferred export: fills the original ERP template (keeps hidden option sheet
 * and all dropdown data validations) so the import utility accepts the file.
 */
export async function buildSRWorkbooksFromTemplate(
  rows: Record<string, unknown>[]
): Promise<ArrayBuffer[]> {
  const { buildSRWorkbookFromTemplate } = await import("./sr-template-writer");
  const chunks = splitSRAtBillBoundary(rows);
  const out: ArrayBuffer[] = [];
  for (const chunk of chunks) {
    try {
      out.push(await buildSRWorkbookFromTemplate(chunk));
    } catch (e) {
      console.error("Template fill failed, falling back to plain workbook", e);
      out.push(buildSRWorkbook(chunk));
    }
  }
  return out;
}


export function validateSaleReturnFile(buf: ArrayBuffer): { valid: boolean; error?: string } {
  try {
    const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: false, sheetRows: 60 });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as unknown[][];
    if (!all || all.length === 0) return { valid: false, error: "File khali (empty) hai." };

    const CORE_SIGNALS = [
      "sr number", "sales return no", "sales return number", "sales return date", "return date", "credit note",
      "product name", "item name", "product description", "product code", "product mrp",
      "basic rate", "taxable value", "sales return value", "gross amount",
      "qty", "quantity", "return qty",
      "bill number", "orig bill no", "original bill no", "invoice number"
    ];

    let headerIdx = -1;
    let maxScore = 0;
    for (let i = 0; i < Math.min(all.length, 60); i++) {
      const row = all[i] || [];
      const normRow = row.map(norm);
      let score = 0;
      for (const cell of normRow) {
        if (!cell) continue;
        for (const sig of CORE_SIGNALS) {
          if (cell === sig || cell.includes(sig)) {
            score++;
            break;
          }
        }
      }
      if (score >= 3 && score > maxScore) {
        maxScore = score;
        headerIdx = i;
      }
    }

    if (headerIdx < 0) {
      return { valid: false, error: "Sale Return header row nahi mila. Expected 'SR Number', 'Return Date', 'Party Name', etc." };
    }

    const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
    const srNumberCol = findHeaderIndex(headers, "srNumber");
    const srDate = findHeaderIndex(headers, "srDate");
    const party = findHeaderIndex(headers, "party");
    const product = findHeaderIndex(headers, "product");
    const qty = findHeaderIndex(headers, "qty");

    const missing: string[] = [];
    if (srNumberCol < 0) missing.push("SR Number / Credit Note No");
    if (srDate < 0) missing.push("Return Date / SR Date");
    if (party < 0) missing.push("Customer Name / Party");
    if (product < 0) missing.push("Product / Item Description");
    if (qty < 0) missing.push("Quantity / Return Qty");

    if (missing.length > 0) {
      return {
        valid: false,
        error: `Sale Return file me required columns missing hain: ${missing.join(", ")}. Kripya valid LeverEDGE Sales Return file upload karein.`
      };
    }
    return { valid: true };
  } catch (e) {
    return { valid: false, error: "File read error: " + (e instanceof Error ? e.message : String(e)) };
  }
}


