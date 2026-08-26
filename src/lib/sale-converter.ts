import * as XLSX from "xlsx";

// Output template column order — MUST exactly match Hisab Kitab sale.xlsx headers (96 columns).
export const OUTPUT_HEADERS = [
  "Invoice Number*","Date*","Customer Name Or Alias Name*",
  "CF - Party Custom Field Name 1","CF - Party Custom Field Name 2","CF - Party Custom Field Name 3",
  "Country Code For Mobile Number","Mobile Number","GSTIN",
  "Billing Address Line 1","Billing Address Line 2","Billing Address Country","Billing Address State","Billing Address City","Billing Address Pincode",
  "Shipping Name","Shipping GSTIN","Shipping Address Line 1","Shipping Address Line 2","Shipping Address Country","Shipping Address State","Shipping Address City","Shipping Address Pincode",
  "Broker Name","Brokerage %","Brokerage On Value",
  "Transport Name","Document number","Document Date","Vehicle Number","Eway Bill Number","Eway Bill Date",
  "PO No.","PO Date","Credit Period","Credit Period Type",
  "CF - Trn. Custom Field Name 1","CF - Trn. Custom Field Name 2","CF - Trn. Custom Field Name 3","CF - Trn. Custom Field Name 4","CF - Trn. Custom Field Name 5",
  "Item Name Or Alias Name Or SKU*",
  "CF - Item Custom Field Name 1","CF - Item Custom Field Name 2","CF - Item Custom Field Name 3","CF - Item Custom Field Name 4","CF - Item Custom Field Name 5",
  "Additional description for Item","Ledger Name","Hsn Code","Unit",
  "Quantity*","Free Quantity Unit","Free Quantity",
  "MRP","Rate per Unit (Without GST)",
  "Discount 1 Type","Discount 1","Discount 2 Type","Discount 2",
  "GST %","Classification Nature Type","RCM Applicable",
  "Additional Charges 1 Ledger","Additional Charges 1 Type","Additional Charges 1 Amount (Without GST)","Additional Charges 1 GST %",
  "Additional Charges 2 Ledger","Additional Charges 2 Type","Additional Charges 2 Amount (Without GST)","Additional Charges 2 GST %",
  "Cess","TCS Ledger","TCS Rate","Taxable Value For TCS",
  "Add / less 1 Ledger","Add / less 1 Type","Add / less 1 Amount",
  "Add / less 2 Ledger","Add / less 2 Type","Add / less 2 Amount",
  "Round off Amount","TDS Ledger","TDS Rate","Taxable Value For TDS",
  "Note","Terms & Conditions",
  "Payment 1 Ledger","Payment 1 Date","Payment 1 Amount","Payment 1 Mode","Payment 1 Reference Number",
  "Payment 2 Ledger","Payment 2 Date","Payment 2 Amount","Payment 2 Mode","Payment 2 Reference Number",
];

// Header aliases for fuzzy/alias matching from source.
const ALIASES: Record<string, string[]> = {
  invoice: [
    "bill number", "bill no", "invoice number", "invoice no", "billno", "invoiceno",
    "bill no.", "invoice no.", "bill num", "inv no", "inv number", "doc no", "bill",
    "doc number", "invoice num", "bill no.*"
  ],
  date: [
    "bill date", "invoice date", "doc date", "sales date", "bill date.", "invoice date.", "date", "doc date."
  ],
  party: [
    "party name", "customer name", "party", "customer", "account name", "ledger name",
    "customer / party", "party/customer", "buyer name", "client name", "party name or alias name"
  ],
  gstin: [
    "party gstin number", "party gstin", "party gst no", "party gst number",
    "gstin number", "gstin", "gst number", "gst no", "party gstin no", "party gst"
  ],
  salesperson: [
    "salesperson name", "sales person name", "sales person", "salesperson",
    "salesman name", "sales man name", "salesman", "sales man",
    "broker name", "broker", "rep name", "sales representative", "agent name",
    "agent", "executive name", "sales executive", "employee name"
  ],
  item: [
    "product description", "item description", "product name", "item name",
    "description", "product", "item", "sku", "particulars", "item particulars",
    "item name or alias name or sku", "item name or alias name or sku*"
  ],
  hsn: [
    "hsn code", "hsn", "hsn/sac", "hsn sac", "sac code", "sac", "hsn_code", "hsn no", "hsn/sac code"
  ],
  units: [
    "units", "quantity", "qty", "unit", "billed qty", "billed quantity",
    "sales qty", "sales units", "tot qty", "total qty", "sale qty", "sale units"
  ],
  freeqty: [
    "free qty", "free quantity", "free units", "free", "scheme qty", "sch qty"
  ],
  mrp: [
    "product mrp", "mrp", "m.r.p.", "unit mrp", "item mrp"
  ],
  basicrate: [
    "basic rate", "rate", "rate per unit", "rate without gst", "item rate",
    "sales rate", "price", "unit rate", "unit price", "basic price"
  ],
  schemedisc: [
    "scheme disc", "scheme discount", "sch disc", "sch discount", "scheme disc amount"
  ],
  rsdisc: [
    "rs discount", "rs disc", "rs product discount"
  ],
  totaldisc: [
    "total disc", "total discount", "item discount", "tot discount", "tot disc",
    "discount", "disc", "cd disc", "cash discount"
  ],
  gstpct: [
    "total tax %", "gst %", "gst%", "tax %", "gst percent", "rate of tax", "tax percentage", "tax rate"
  ],
  netsales: [
    "net sales", "net amount", "net value", "net", "invoice amount", "net amt",
    "net credit value (after discount reversal)", "net credit value", "net credit amt", "net credit"
  ],
  grosssales: [
    "gross sales", "gross amount", "gross value", "gross sale", "gross", "gross amt", "gross value amount"
  ],
  taxableamt: [
    "taxable value", "taxable amount", "taxable amt", "taxable", "taxable val"
  ],
  taxvalue: [
    "tax value", "tax amount", "tax amt", "total tax", "tax", "gst amount", "gst amt"
  ],
};

const norm = (s: unknown) =>
  String(s ?? "").toLowerCase().replace(/[^a-z0-9%]/g, " ").replace(/\s+/g, " ").trim();

/**
 * Alias-priority header lookup: earlier aliases win over later ones, regardless of
 * the column order in the source file. (e.g. "Total Disc" beats "TOT Disc".)
 */
function findHeaderIndex(headers: string[], key: keyof typeof ALIASES): number {
  const aliases = ALIASES[key].map(norm);
  const normHeaders = headers.map(norm);
  // exact match, alias by alias (priority order)
  for (const a of aliases) {
    const i = normHeaders.indexOf(a);
    if (i >= 0) return i;
  }
  // partial / substring match, alias by alias
  for (const a of aliases) {
    for (let i = 0; i < normHeaders.length; i++) {
      if (!normHeaders[i]) continue;
      if (normHeaders[i].includes(a) || (a.length >= 4 && a.includes(normHeaders[i]))) return i;
    }
  }
  return -1;
}


function cleanHSN(v: unknown): string {
  if (v === null || v === undefined) return "";
  const digits = String(v).replace(/[^\d]/g, "");
  return digits.slice(0, 8);
}

export function normBillNo(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(Math.round(v));
  return String(v).trim();
}

export function getNormalizedBillKeys(docNo: string): string[] {
  const s = String(docNo ?? "").trim();
  if (!s) return [];
  const keys = new Set<string>();
  keys.add(s);
  keys.add(s.toLowerCase());
  keys.add(s.toUpperCase());

  const clean = s.replace(/[^a-zA-Z0-9]/g, "");
  if (clean) {
    keys.add(clean);
    keys.add(clean.toLowerCase());
    keys.add(clean.toUpperCase());
  }

  // Strip trailing slash year e.g. "30392/24-25" -> "30392"
  const strippedYear = s.split(/[\/\-_]/)[0].trim();
  if (strippedYear && strippedYear !== s) {
    keys.add(strippedYear);
    keys.add(strippedYear.toLowerCase());
    keys.add(strippedYear.toUpperCase());
  }

  const numOnly = s.replace(/[^0-9]/g, "").replace(/^0+/, "");
  if (numOnly) {
    keys.add(numOnly);
    keys.add("GST" + numOnly);
    keys.add("gst" + numOnly);
    keys.add("GST-" + numOnly);
    keys.add("GST/" + numOnly);
    keys.add("INV" + numOnly);
  }

  if (/^gst/i.test(s)) {
    const withoutGst = s.replace(/^gst[\/\-_ ]*/i, "").trim();
    if (withoutGst) {
      keys.add(withoutGst);
      const withoutGstNoZero = withoutGst.replace(/^0+/, "");
      if (withoutGstNoZero) {
        keys.add(withoutGstNoZero);
        keys.add("GST" + withoutGstNoZero);
        keys.add("gst" + withoutGstNoZero);
      }
    }
  }

  return Array.from(keys);
}

function cleanGSTIN(v: unknown): string {
  const s = String(v ?? "").toUpperCase().replace(/\s+/g, "");
  return /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]Z[0-9A-Z]$/.test(s) ? s : "";
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
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[₹,\s]/g, ""));
  return isNaN(n) ? 0 : n;
};

export interface RegisterDiscount {
  ushop: number;
  shikhar: number;
  btpr: number;
  cash?: number;
  adjustments?: number;
  totalAddLess1?: number;
  roundOff?: number;
}

export function registerDiscountKey(map: Map<string, RegisterDiscount>, rawKey: unknown, disc: RegisterDiscount) {
  const s = String(rawKey ?? "").trim();
  if (!s) return;
  const allKeys = getNormalizedBillKeys(s);
  for (const k of allKeys) {
    map.set(k, disc);
  }
}

export function lookupDiscount(map: Map<string, RegisterDiscount>, invoiceNo: unknown): RegisterDiscount | undefined {
  if (!invoiceNo || !map) return undefined;
  const s = String(invoiceNo).trim();
  if (!s) return undefined;

  if (map.has(s)) return map.get(s);

  const allKeys = getNormalizedBillKeys(s);
  for (const k of allKeys) {
    if (map.has(k)) return map.get(k);
  }

  return undefined;
}

/**
 * Merge multiple rows for the same item within the same bill/invoice into a single row.
 * - Total Quantity = Sum of quantities
 * - Total Base Amount = Sum of (quantity * ratePerUnit)
 * - Rate per Unit = Total Base Amount / Total Quantity (weighted average rate, 4 decimals)
 * - Free Quantity = Sum of free quantities
 * - Discount = Sum of discount amounts
 */
export function mergeSameItemRowsInBill(
  rows: Record<string, unknown>[],
  billKey: string,
  itemKey: string
): Record<string, unknown>[] {
  if (rows.length === 0) return rows;

  const merged: Record<string, unknown>[] = [];
  const billGroups = new Map<string, Record<string, unknown>[]>();
  const billOrder: string[] = [];

  for (const r of rows) {
    const bill = String(r[billKey] ?? "").trim();
    let list = billGroups.get(bill);
    if (!list) {
      list = [];
      billGroups.set(bill, list);
      billOrder.push(bill);
    }
    list.push(r);
  }

  for (const bill of billOrder) {
    const bRows = billGroups.get(bill) || [];
    const itemMap = new Map<string, Record<string, unknown>[]>();
    const itemOrder: string[] = [];

    for (const r of bRows) {
      const item = String(r[itemKey] ?? "").trim().toLowerCase();
      if (!itemMap.has(item)) {
        itemMap.set(item, [r]);
        itemOrder.push(item);
      } else {
        itemMap.get(item)!.push(r);
      }
    }

    for (const item of itemOrder) {
      const group = itemMap.get(item)!;
      if (group.length === 1) {
        merged.push(group[0]);
      } else {
        const primary = { ...group[0] };
        let totalQty = 0;
        let totalBaseAmt = 0;
        let totalFreeQty = 0;
        let totalDisc1 = 0;
        let totalDisc2 = 0;
        let maxMrp = 0;

        for (const r of group) {
          const q = typeof r["Quantity*"] === "number"
            ? r["Quantity*"]
            : parseFloat(String(r["Quantity*"] ?? 0)) || 0;

          const rawRate = r["Rate per Unit (Without GST)"] ?? r["Rate per Unit (Without GST)*"];
          const rate = typeof rawRate === "number"
            ? rawRate
            : parseFloat(String(rawRate ?? 0)) || 0;

          totalQty += q;
          totalBaseAmt += (q * rate);

          const freeQ = typeof r["Free Quantity"] === "number"
            ? r["Free Quantity"]
            : parseFloat(String(r["Free Quantity"] ?? 0)) || 0;
          totalFreeQty += freeQ;

          const d1 = typeof r["Discount 1"] === "number"
            ? r["Discount 1"]
            : parseFloat(String(r["Discount 1"] ?? 0)) || 0;
          totalDisc1 += d1;

          const d2 = typeof r["Discount 2"] === "number"
            ? r["Discount 2"]
            : parseFloat(String(r["Discount 2"] ?? 0)) || 0;
          totalDisc2 += d2;

          const mrp = typeof r["MRP"] === "number"
            ? r["MRP"]
            : parseFloat(String(r["MRP"] ?? 0)) || 0;
          if (mrp > maxMrp) maxMrp = mrp;
        }

        primary["Quantity*"] = totalQty;
        const avgRate = totalQty > 0 ? Math.round((totalBaseAmt / totalQty) * 10000) / 10000 : 0;
        if ("Rate per Unit (Without GST)" in primary) {
          primary["Rate per Unit (Without GST)"] = avgRate;
        }
        if ("Rate per Unit (Without GST)*" in primary) {
          primary["Rate per Unit (Without GST)*"] = avgRate;
        }

        if (totalFreeQty > 0) {
          primary["Free Quantity"] = totalFreeQty;
        }
        if (totalDisc1 > 0) {
          primary["Discount 1 Type"] = "₹";
          primary["Discount 1"] = Math.round(totalDisc1 * 100) / 100;
        }
        if (totalDisc2 > 0) {
          primary["Discount 2 Type"] = "₹";
          primary["Discount 2"] = Math.round(totalDisc2 * 100) / 100;
        }
        if (maxMrp > 0) {
          primary["MRP"] = maxMrp;
        }

        merged.push(primary);
      }
    }
  }

  return merged;
}

/**
 * Parse the Sales Register file (LeverEDGE_..._Sales_Register_...) and return a Map of
 * Bill No → { ushop, shikhar, btpr, cash, adjustments, totalAddLess1, roundOff } discount amounts.
 * 
 * Formula:
 * USHOP DIS (Add/Less 1) = BTPR SchDisc + Ushop Redemption
 * SHIKHAR DIS (Add/Less 2) = Shikhar Scheme / SchDisc
 */
export function parseSalesRegister(buf: ArrayBuffer, forReturn = false): { discountMap: Map<string, RegisterDiscount>; rows: Record<string, unknown>[]; billValues: Map<string, number> } {
  const wb = XLSX.read(buf, { type: "array", cellDates: false, cellFormula: false, cellHTML: false, cellText: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  // Scan up to 100 rows to find the header row with the highest number of recognizable signals.
  const HEADER_SIGNALS = [
    "bill no", "bill number", "invoice no", "invoice number", "ushop", "shikhar", "btpr", "btrp", "cash", "adjustments", "round off"
  ];
  let headerIdx = -1;
  let maxMatches = 0;
  for (let i = 0; i < Math.min(all.length, 100); i++) {
    const r = (all[i] || []).map((c) => norm(c));
    const matches = r.filter((cell) => cell && HEADER_SIGNALS.some((sig) => cell === sig || cell.includes(sig))).length;
    if (matches > maxMatches) {
      maxMatches = matches;
      headerIdx = i;
    }
  }
  if (headerIdx < 0) {
    for (let i = 0; i < Math.min(all.length, 30); i++) {
      const nonEmpties = (all[i] || []).filter((c) => c !== null && c !== "" && isNaN(Number(c))).length;
      if (nonEmpties >= 3) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx < 0) headerIdx = 0;
  }

  const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
  let dataRows = all.slice(headerIdx + 1).filter((r) => r.some((c) => c !== null && c !== ""));

  while (dataRows.length > 0) {
    const last = dataRows[dataRows.length - 1];
    const firstCellNorm = norm(last[0]);
    if (firstCellNorm === "" || firstCellNorm.includes("grand total") || firstCellNorm.includes("total")) {
      dataRows = dataRows.slice(0, -1);
    } else {
      break;
    }
  }

  const normHeaders = headers.map(norm);

  // Helper for specific discount column lookup with keyword guard to prevent false positive on generic "disc"
  const findDiscountCol = (requiredKeywords: string[], ...aliases: string[]): number => {
    const normAliases = aliases.map(norm);
    // 1. Exact match with any alias
    for (let i = 0; i < normHeaders.length; i++) {
      if (!normHeaders[i]) continue;
      if (normAliases.includes(normHeaders[i])) return i;
    }
    // 2. Keyword + alias contains match
    for (let i = 0; i < normHeaders.length; i++) {
      const h = normHeaders[i];
      if (!h) continue;
      const hasReq = requiredKeywords.some((kw) => h.includes(kw));
      if (!hasReq) continue;
      for (const a of normAliases) {
        if (h.includes(a) || a.includes(h)) return i;
      }
    }
    // 3. Fallback: Contains any required keyword
    for (let i = 0; i < normHeaders.length; i++) {
      const h = normHeaders[i];
      if (!h) continue;
      if (requiredKeywords.some((kw) => h.includes(kw))) return i;
    }
    return -1;
  };

  const findBillCol = (...aliases: string[]): number => {
    const normAliases = aliases.map(norm);
    const nonBillIgnore = ["sr no", "sl no", "serial no", "s no", "order no", "phone no", "mobile no", "item no", "party no", "customer no"];
    // 1. Exact match
    for (let i = 0; i < normHeaders.length; i++) {
      if (!normHeaders[i] || nonBillIgnore.includes(normHeaders[i])) continue;
      if (normAliases.includes(normHeaders[i])) return i;
    }
    // 2. Contains match
    for (let i = 0; i < normHeaders.length; i++) {
      const h = normHeaders[i];
      if (!h || nonBillIgnore.includes(h)) continue;
      for (const a of normAliases) {
        if (h.includes(a)) return i;
      }
    }
    return -1;
  };

  const billNoCol   = findBillCol("bill no", "bill number", "invoice no", "invoice number", "billno", "invoiceno", "bill no.", "invoice no.", "doc number", "doc no", "document no", "document number", "bill ref no", "billrefno", "bill_no", "inv_no", "voucher no");
  const billValCol  = findBillCol("bill value", "bill val", "net amount", "invoice amount", "net value", "bill amt", "total amount", "net amt", "invoice value", "net credit value");

  // USHOP: Must contain "ushop" or "redemption"
  const ushopCol    = findDiscountCol(["ushop", "redemption"], "ushop redemption", "ushop schdisc", "ushop sch disc", "ushop scheme discount", "ushop scheme disc", "ushop discount", "ushop disc", "ushop coupon", "ushop scheme", "ushop sch", "ushop", "u-shop", "u shop", "redemption", "redemption discount", "ushop amount", "ushop amt", "ushop dic");

  // SHIKHAR: Must contain "shikhar" or "shikar"
  const shikharCol  = findDiscountCol(["shikhar", "shikar"], "shikhar scheme", "shikhar schdisc", "shikhar sch disc", "shikhar scheme discount", "shikhar scheme disc", "shikhar discount", "shikhar disc", "shikhar", "shikar", "shikhar dic", "shikar disc", "shikhar amount", "shikhar amt");

  // BTPR: Must contain "btpr" or "btrp" or "bt pr"
  const btprCol     = findDiscountCol(["btpr", "btrp", "bt pr"], "btpr schdisc", "btrp schdisc", "btpr sch disc", "btrp sch disc", "btpr scheme discount", "btrp scheme discount", "btpr scheme", "btrp scheme", "btpr discount", "btrp discount", "btpr disc", "btrp disc", "btpr", "btrp", "bt pr", "btpr amount", "btrp amount", "btpr amt", "btrp amt", "btpr val", "btrp val");

  // CASH: Must contain "cash" or "cd" or "c.d"
  const cashCol     = findDiscountCol(["cash", "cd", "c d"], "cashdisc", "cash disc", "cash discount", "cash schdisc", "cash sch disc", "cash scheme disc", "cash scheme discount", "cash amt", "cash amount", "cd disc", "cd discount", "cd amt", "cd amount", "cash dis", "cd", "c.d. disc", "c.d disc", "cash dic");

  // ADJUSTMENTS: Must contain "adj" or "adjust"
  const adjCol      = findDiscountCol(["adj", "adjust"], "adjustments", "adjustment", "adj amt", "adj amount", "adjust", "adj", "other discount", "other disc", "other adjustments");

  // ROUND OFF: Must contain "round" or "rnd"
  const roundOffCol = findDiscountCol(["round", "rnd"], "round off", "round off amt", "round off amount", "roundoff", "round-off", "rnd off", "rnd off amt", "round_off", "roff");

  const result = new Map<string, RegisterDiscount>();

  for (const r of dataRows) {
    const billNo = normBillNo(r[billNoCol]);
    if (!billNo || norm(billNo).includes("grand total") || norm(billNo).includes("total")) continue;

    const billVal = billValCol >= 0 ? num(r[billValCol]) : 0;
    // Sale data me ONLY positive (+) Bill Value rows ka data liya jana chahiye
    if (!forReturn && billValCol >= 0 && billVal <= 0) continue;
    // Return data me ONLY negative (-) Bill Value rows ka data liya jana chahiye
    if (forReturn && billValCol >= 0 && billVal >= 0) continue;

    const ushop    = ushopCol    >= 0 ? Math.abs(num(r[ushopCol]))    : 0;
    const shikhar  = shikharCol  >= 0 ? Math.abs(num(r[shikharCol]))  : 0;
    const btpr     = btprCol     >= 0 ? Math.abs(num(r[btprCol]))     : 0;
    const cash     = cashCol     >= 0 ? Math.abs(num(r[cashCol]))     : 0;
    const adj      = adjCol      >= 0 ? Math.abs(num(r[adjCol]))      : 0;
    const roundOff = roundOffCol >= 0 ? num(r[roundOffCol])           : 0;

    // USHOP DIS (Add/Less 1) = BTPR SchDisc + Ushop Redemption
    const totalAddLess1 = btpr + ushop;

    if (totalAddLess1 === 0 && shikhar === 0 && roundOff === 0) continue;

    const discObj: RegisterDiscount = {
      ushop,
      shikhar,
      btpr,
      cash,
      adjustments: adj,
      totalAddLess1,
      roundOff,
    };

    const existing = lookupDiscount(result, billNo);
    if (existing) {
      existing.ushop   += ushop;
      existing.shikhar += shikhar;
      existing.btpr    += btpr;
      existing.cash    = (existing.cash ?? 0) + cash;
      existing.adjustments = (existing.adjustments ?? 0) + adj;
      existing.totalAddLess1 = (existing.totalAddLess1 ?? 0) + totalAddLess1;
      if (roundOff !== 0) existing.roundOff = roundOff;
      registerDiscountKey(result, billNo, existing);
    } else {
      registerDiscountKey(result, billNo, discObj);
    }
  }

  // Bill-level BillValue map — used for Sale Report reconciliation
  const billValues = new Map<string, number>();
  if (billValCol >= 0) {
    for (const r of dataRows) {
      const billNo = normBillNo(r[billNoCol]);
      if (!billNo || norm(billNo).includes("total")) continue;
      const billVal = num(r[billValCol]);
      if (!forReturn && billVal <= 0) continue;
      if (forReturn && billVal >= 0) continue;
      const key = billKeyOf(billNo);
      billValues.set(key, Math.round(((billValues.get(key) ?? 0) + billVal) * 100) / 100);
    }
  }

  const rows: Record<string, unknown>[] = dataRows
    .filter((r) => {
      const billNo = normBillNo(r[billNoCol]);
      if (!billNo || norm(billNo).includes("total")) return false;
      const billVal = billValCol >= 0 ? num(r[billValCol]) : 0;
      if (!forReturn && billValCol >= 0 && billVal <= 0) return false;
      if (forReturn && billValCol >= 0 && billVal >= 0) return false;
      return true;
    })
    .map((r) => {
      const o: Record<string, unknown> = {};
      headers.forEach((h, i) => { if (h) o[h] = r[i] ?? ""; });
      return o;
    });

  return { discountMap: result, rows, billValues };
}

/** Canonical key for a bill number: digits only (leading zeros stripped) when digits exist. */
export function billKeyOf(billNo: unknown): string {
  const s = String(billNo ?? "").trim().toUpperCase();
  const digits = s.replace(/[^0-9]/g, "").replace(/^0+/, "");
  return digits || s;
}

export interface ReturnRegEntry {
  billNo: string;
  srNo: string;
  retValue: number;
  schDisc: number;
  cashDisc: number;
  taxAmt: number;
  crNoteAmt: number;
}

/**
 * Parse a bill-level Sales Return Register (Sr No / Salesperson / Sales Ret. No / Bill No /
 * Cr. Note No. / Cr. Note Amt. / Ret. Value / Sch Disc / CashDisc / Tax Amt ...).
 * Returns Map billKey → aggregated return entry.
 */
export function parseSalesReturnRegister(buf: ArrayBuffer): Map<string, ReturnRegEntry> {
  const map = new Map<string, ReturnRegEntry>();
  try {
    const wb = XLSX.read(buf, { type: "array", cellDates: false, cellFormula: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

    const SIGNALS = ["sales ret no", "ret value", "cr note amt", "cr note no", "bill no", "sch disc", "tax amt"];
    let headerIdx = -1;
    let best = 0;
    for (let i = 0; i < Math.min(all.length, 60); i++) {
      const r = (all[i] || []).map((c) => norm(c));
      const n = r.filter((cell) => cell && SIGNALS.some((s) => cell === s || cell.includes(s))).length;
      if (n > best) { best = n; headerIdx = i; }
    }
    if (headerIdx < 0 || best < 3) return map;

    const headers = (all[headerIdx] as unknown[]).map((c) => norm(c));
    const col = (...names: string[]) => {
      for (const n of names) {
        const i = headers.findIndex((h) => h === norm(n));
        if (i >= 0) return i;
      }
      for (const n of names) {
        const i = headers.findIndex((h) => h && h.includes(norm(n)));
        if (i >= 0) return i;
      }
      return -1;
    };

    const cBill = col("bill no", "bill number", "billrefno", "invoice no");
    const cSr   = col("sales ret no", "sr no.", "sales return no", "cr note no");
    const cRet  = col("ret value", "return value", "ret amt");
    const cSch  = col("sch disc", "scheme disc");
    const cCash = col("cashdisc", "cash disc");
    const cTax  = col("tax amt", "tax amount");
    const cCr   = col("cr note amt", "credit note amt", "total");
    if (cBill < 0 || cRet < 0) return map;

    for (const r of all.slice(headerIdx + 1)) {
      const billNo = normBillNo(r[cBill]);
      if (!billNo || norm(billNo).includes("total")) continue;
      const key = billKeyOf(billNo);
      const prev = map.get(key);
      const entry: ReturnRegEntry = {
        billNo,
        srNo: cSr >= 0 ? String(r[cSr] ?? "").trim() : "",
        retValue: Math.abs(num(r[cRet])),
        schDisc: cSch >= 0 ? Math.abs(num(r[cSch])) : 0,
        cashDisc: cCash >= 0 ? Math.abs(num(r[cCash])) : 0,
        taxAmt: cTax >= 0 ? Math.abs(num(r[cTax])) : 0,
        crNoteAmt: cCr >= 0 ? Math.abs(num(r[cCr])) : 0,
      };
      if (prev) {
        prev.retValue += entry.retValue;
        prev.schDisc += entry.schDisc;
        prev.cashDisc += entry.cashDisc;
        prev.taxAmt += entry.taxAmt;
        prev.crNoteAmt += entry.crNoteAmt;
      } else {
        map.set(key, entry);
      }
    }
  } catch (e) {
    console.warn("[parseSalesReturnRegister]", e);
  }
  return map;
}

export interface ReconcileResult {
  billsChecked: number;
  matched: number;
  mismatched: number;
  registerTotal: number;
  reportTotal: number;
  missingInReport: number;
  missingInRegister: number;
  diffs: { bill: string; register: number; report: number; diff: number }[];
}

/**
 * Compare Sales Register BillValue (sale + return netted per bill) against the
 * Sale Report "Invoice Amount" column. Tolerance ±1 rupee (round-off).
 */
export function reconcileWithSaleReport(
  registerBillValues: Map<string, number>,
  reportBuf: ArrayBuffer,
  tolerance = 1,
): ReconcileResult {
  const res: ReconcileResult = {
    billsChecked: 0, matched: 0, mismatched: 0,
    registerTotal: 0, reportTotal: 0,
    missingInReport: 0, missingInRegister: 0, diffs: [],
  };

  const wb = XLSX.read(reportBuf, { type: "array", cellDates: false, cellFormula: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  const SIGNALS = ["document number", "invoice amount", "taxable value", "transaction type"];
  let headerIdx = -1, best = 0;
  for (let i = 0; i < Math.min(all.length, 60); i++) {
    const r = (all[i] || []).map((c) => norm(c));
    const n = r.filter((cell) => cell && SIGNALS.some((s) => cell === s || cell.includes(s))).length;
    if (n > best) { best = n; headerIdx = i; }
  }
  if (headerIdx < 0 || best < 2) throw new Error("Sale Report header row nahi mila ('Document Number' / 'Invoice Amount').");

  const headers = (all[headerIdx] as unknown[]).map((c) => norm(c));
  const find = (...names: string[]) => {
    for (const n of names) { const i = headers.findIndex((h) => h === norm(n)); if (i >= 0) return i; }
    for (const n of names) { const i = headers.findIndex((h) => h && h.includes(norm(n))); if (i >= 0) return i; }
    return -1;
  };
  const cDoc = find("document number", "doc number", "invoice number", "bill number");
  const cAmt = find("invoice amount", "bill value", "net amount", "total amount");
  if (cDoc < 0 || cAmt < 0) throw new Error("Sale Report me 'Document Number' ya 'Invoice Amount' column nahi mila.");

  const reportMap = new Map<string, number>();
  for (const r of all.slice(headerIdx + 1)) {
    const doc = normBillNo(r[cDoc]);
    if (!doc || norm(doc).includes("total")) continue;
    const key = billKeyOf(doc);
    reportMap.set(key, Math.round(((reportMap.get(key) ?? 0) + num(r[cAmt])) * 100) / 100);
  }

  const keys = new Set<string>([...registerBillValues.keys(), ...reportMap.keys()]);
  for (const k of keys) {
    const reg = registerBillValues.get(k);
    const rep = reportMap.get(k);
    if (reg === undefined) { res.missingInRegister++; res.reportTotal += rep ?? 0; continue; }
    if (rep === undefined) { res.missingInReport++; res.registerTotal += reg; continue; }
    res.billsChecked++;
    res.registerTotal += reg;
    res.reportTotal += rep;
    const diff = Math.round((reg - rep) * 100) / 100;
    if (Math.abs(diff) <= tolerance) res.matched++;
    else {
      res.mismatched++;
      if (res.diffs.length < 200) res.diffs.push({ bill: k, register: reg, report: rep, diff });
    }
  }
  res.registerTotal = Math.round(res.registerTotal * 100) / 100;
  res.reportTotal = Math.round(res.reportTotal * 100) / 100;
  return res;
}


export interface ConvertStats {
  sourceRows: number;
  exportRows: number;
  removedNegative: number;
  rejectedInvalid: number;
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
  errors: { row: number; reason: string }[];
  mappedHeaders: Record<string, string>;
  srMissingBillsAdded?: number;
  srMissingRowsAdded?: number;
  zeroQtyRebuilt?: number;
  zeroQtyUnresolved?: number;
}

export interface ConvertResult {
  rows: Record<string, unknown>[];
  stats: ConvertStats;
}

export function validateSaleDataFile(buf: ArrayBuffer): { valid: boolean; error?: string } {
  try {
    const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: false, sheetRows: 60 });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as unknown[][];
    if (!all || all.length === 0) return { valid: false, error: "File khali (empty) hai." };

    let headerIdx = -1;
    let maxMatch = 0;
    const BILL_SIGNALS = ["bill number", "bill no", "invoice number", "invoice no", "net sales", "party name", "customer name", "product description"];
    for (let i = 0; i < Math.min(all.length, 60); i++) {
      const r = (all[i] || []).map((c) => norm(c));
      const matchCount = r.filter((cell) => BILL_SIGNALS.some((sig) => cell === sig || cell.includes(sig))).length;
      if (matchCount > maxMatch) {
        maxMatch = matchCount;
        headerIdx = i;
      }
    }
    if (headerIdx < 0 || maxMatch < 2) {
      return { valid: false, error: "Bill Wise Sales header row nahi mila. Expected 'Bill Number', 'Party Name', 'Net Sales', etc." };
    }
    const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
    const invoice = findHeaderIndex(headers, "invoice");
    const date = findHeaderIndex(headers, "date");
    const party = findHeaderIndex(headers, "party");
    const salesperson = findHeaderIndex(headers, "salesperson");
    const item = findHeaderIndex(headers, "item");
    const units = findHeaderIndex(headers, "units");
    const netsales = findHeaderIndex(headers, "netsales");

    const missing: string[] = [];
    if (invoice < 0) missing.push("Bill Number / Invoice No");
    if (date < 0) missing.push("Bill Date");
    if (party < 0) missing.push("Customer Name / Party");
    if (salesperson < 0) missing.push("Salesperson Name");
    if (item < 0) missing.push("Item Name / Product Description");
    if (units < 0) missing.push("Units / Quantity");
    if (netsales < 0) missing.push("Net Sales");

    if (missing.length > 0) {
      return {
        valid: false,
        error: `Sale Data file me required columns missing hain: ${missing.join(", ")}. Kripya valid LeverEDGE Bill Wise Sales file upload karein.`
      };
    }
    return { valid: true };
  } catch (e) {
    return { valid: false, error: "File read error: " + (e instanceof Error ? e.message : String(e)) };
  }
}

export function validateSalesRegisterFile(buf: ArrayBuffer): { valid: boolean; error?: string } {
  try {
    const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: false, sheetRows: 60 });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as unknown[][];
    if (!all || all.length === 0) return { valid: false, error: "File khali (empty) hai." };

    const SIGNALS = ["bill no", "bill number", "invoice no", "bill ref no", "billrefno", "bill value", "billvalue", "net sales", "net amount"];
    let headerIdx = -1;
    let maxMatch = 0;
    for (let i = 0; i < Math.min(all.length, 60); i++) {
      const r = (all[i] || []).map((c) => norm(c));
      const matchCount = r.filter((cell) => SIGNALS.some((sig) => cell === sig || cell.includes(sig))).length;
      if (matchCount > maxMatch) {
        maxMatch = matchCount;
        headerIdx = i;
      }
    }
    if (headerIdx < 0 || maxMatch < 2) {
      return { valid: false, error: "Sales Register header row nahi mila. Expected 'BillRefNo' and 'BillValue'." };
    }
    const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
    const normHeaders = headers.map(norm);
    const findCol = (...aliases: string[]) => {
      const normAliases = aliases.map(norm);
      for (let i = 0; i < normHeaders.length; i++) {
        if (!normHeaders[i]) continue;
        if (normAliases.includes(normHeaders[i])) return i;
      }
      for (let i = 0; i < normHeaders.length; i++) {
        if (!normHeaders[i]) continue;
        for (const a of normAliases) {
          if (normHeaders[i].includes(a) || (a.length >= 4 && a.includes(normHeaders[i]))) return i;
        }
      }
      return -1;
    };
    const billNoCol = findCol("billrefno", "bill ref no", "bill no", "bill number", "invoice no", "invoice number", "billno", "invoiceno", "bill no.", "invoice no.");
    const billValCol = findCol("billvalue", "bill value", "bill val", "billval", "bill amt", "invoice amount", "net amount", "total amount", "net sales");

    const missing: string[] = [];
    if (billNoCol < 0) missing.push("BillRefNo / Bill No");
    if (billValCol < 0) missing.push("BillValue / Net Amount");

    if (missing.length > 0) {
      return {
        valid: false,
        error: `Sales Register file me required columns missing hain: ${missing.join(", ")}. Kripya valid LeverEDGE Sales Register file upload karein.`
      };
    }
    return { valid: true };
  } catch (e) {
    return { valid: false, error: "File read error: " + (e instanceof Error ? e.message : String(e)) };
  }
}

// Aliases for Sale Return columns
const SR_SUB_ALIASES: Record<string, string[]> = {
  srNumber: [
    "sr number", "sr no", "sr_no", "sr no.", "sales return no", "sales return number",
    "credit note no", "credit note number", "credit note no.", "cn number", "cn no", "return no", "voucher no"
  ],
  srDate: [
    "sales return date", "return date", "sr date", "credit note date", "cn date", "doc date", "voucher date", "date"
  ],
  billNumber: [
    "orig bill no", "orig bill no.", "orig bill number", "original bill no", "original bill number",
    "original invoice no", "orig invoice no", "orig invoice number", "original invoice number",
    "bill ref no", "bill ref. no.", "billrefno", "ref bill no", "against bill no", "against inv no",
    "bill number", "bill no", "bill no.", "invoice number", "invoice no", "invoice no.", "ref no", "reference no",
    "bill num", "orig doc no"
  ],
  billDate: [
    "orig bill date", "original bill date", "orig invoice date", "original invoice date", "bill date", "invoice date", "date"
  ],
  party: [
    "party name", "customer name", "customer", "party", "account name", "ledger name", "party/customer", "buyer name", "client name"
  ],
  gstin: [
    "party gstin number", "party gstin", "party gst no", "party gst number", "gstin number", "gstin", "gst number", "gst no", "party gstin no"
  ],
  salesperson: [
    "salesperson name", "sales person name", "sales person", "salesperson",
    "salesman name", "sales man name", "salesman", "sales man",
    "broker name", "broker", "rep name", "sales representative", "agent name", "agent"
  ],
  product: [
    "product description", "product name", "item description", "item name", "description", "product", "item", "sku", "particulars", "item particulars"
  ],
  hsn: [
    "hsn code", "hsn", "hsn/sac", "hsn sac", "sac", "hsn_code", "hsn no"
  ],
  mrp: [
    "product mrp", "mrp", "m.r.p.", "unit mrp", "item mrp"
  ],
  basicRate: [
    "basic rate", "rate", "rate per unit", "rate without gst", "item rate", "sales rate", "price", "unit rate", "unit price", "basic price"
  ],
  qty: [
    "qty", "quantity", "units", "return qty", "sales return qty", "ret qty", "billed qty", "sale return qty", "billed quantity", "tot qty", "total qty"
  ],
  freeQty: [
    "free qty", "free quantity", "free units", "free", "scheme qty", "sch qty"
  ],
  schemeDisc: [
    "scheme discount", "scheme disc", "sch disc", "sch discount", "scheme discount amount", "scheme discount amt", "scheme disc amount"
  ],
  cashDisc: [
    "cash discount", "cash disc", "cd disc", "cd discount", "cash discount amount", "cash discount amt", "cash disc amount"
  ],
  totDisc: [
    "total discount", "total disc", "tot discount", "tot disc", "discount", "disc", "item discount", "other discount"
  ],
  taxableValue: [
    "taxable value", "taxable amt", "taxable amount", "taxable", "taxable val"
  ],
  taxValue: [
    "tax value", "tax amount", "tax amt", "total tax", "tax", "gst amount", "gst amt"
  ],
  gstpct: [
    "total tax %", "gst %", "gst%", "tax %", "gst percent", "rate of tax", "tax percentage", "tax rate"
  ],
  netsales: [
    "net credit value (after discount reversal)", "net credit value", "net credit amt", "net credit", "net amt", "net sales", "net amount", "net value", "net return", "net", "invoice amount", "credit amount", "credit amt"
  ],
  grosssales: [
    "gross amount", "gross sales", "gross value", "gross return", "gross", "gross amt", "gross value amount"
  ],
};

const GST_SLABS = [0, 5, 12, 18, 28];
function snapGST(taxVal: number, taxableVal: number): number {
  if (taxableVal <= 0) return 0;
  const pct = (taxVal / taxableVal) * 100;
  return GST_SLABS.reduce((prev, cur) =>
    Math.abs(cur - pct) < Math.abs(prev - pct) ? cur : prev
  );
}

function findSRSubHeaderIndex(headers: string[], key: keyof typeof SR_SUB_ALIASES, excludeCol = -1): number {
  const aliases = SR_SUB_ALIASES[key].map(norm);
  const normHeaders = headers.map(norm);
  for (let i = 0; i < normHeaders.length; i++) {
    if (i === excludeCol) continue;
    if (aliases.includes(normHeaders[i])) return i;
  }
  for (let i = 0; i < normHeaders.length; i++) {
    if (i === excludeCol) continue;
    for (const a of aliases) {
      if (normHeaders[i].includes(a) || (a.length >= 4 && a.includes(normHeaders[i]))) return i;
    }
  }
  return -1;
}

export interface SRItemParsed {
  origBillNo: string;
  srNumber: string;
  party: string;
  gstin: string;
  salesperson: string;
  product: string;
  hsn: string;
  mrp: number;
  basicRate: number;
  qty: number;
  freeQty: number;
  schemeDisc: number;
  cashDisc: number;
  totDisc: number;
  taxableValue: number;
  taxValue: number;
  gstpct: number;
  netsales: number;
  grosssales: number;
  dateVal: unknown;
  used: boolean;
}

export function parseSRBufferItems(srBuf: ArrayBuffer): SRItemParsed[] {
  try {
    const wbSR = XLSX.read(srBuf, { type: "array", cellDates: false, cellFormula: false });
    const wsSR = wbSR.Sheets[wbSR.SheetNames[0]];
    const allSR: unknown[][] = XLSX.utils.sheet_to_json(wsSR, { header: 1, raw: true, defval: null });

    const HEADER_SIGNALS_SR = [
      "sr number", "sr no", "sr_no", "sales return date", "sales return no",
      "return date", "return no", "credit note", "bill number", "bill no",
      "invoice number", "invoice no", "party name", "customer name", "net credit value"
    ];

    let headerIdxSR = -1;
    let maxMatchSR = 0;
    for (let i = 0; i < Math.min(allSR.length, 60); i++) {
      const r = (allSR[i] || []).map((c) => norm(c));
      const count = r.filter((cell) => HEADER_SIGNALS_SR.some((sig) => cell === sig || cell.includes(sig))).length;
      if (count > maxMatchSR) {
        maxMatchSR = count;
        headerIdxSR = i;
      }
    }

    if (headerIdxSR < 0 || maxMatchSR < 2) return [];

    const headersSR = (allSR[headerIdxSR] as unknown[]).map((c) => String(c ?? ""));
    const dataRowsSR = allSR.slice(headerIdxSR + 1).filter((r) => r.some((c) => c !== null && c !== ""));

    const srNumberCol = findSRSubHeaderIndex(headersSR, "srNumber");
    const billNumberCol = findSRSubHeaderIndex(headersSR, "billNumber", srNumberCol);

    const idxSR = {
      srNumber:     srNumberCol,
      billNumber:   billNumberCol,
      srDate:       findSRSubHeaderIndex(headersSR, "srDate"),
      billDate:     findSRSubHeaderIndex(headersSR, "billDate"),
      party:        findSRSubHeaderIndex(headersSR, "party"),
      gstin:        findSRSubHeaderIndex(headersSR, "gstin"),
      salesperson:  findSRSubHeaderIndex(headersSR, "salesperson"),
      product:      findSRSubHeaderIndex(headersSR, "product"),
      hsn:          findSRSubHeaderIndex(headersSR, "hsn"),
      mrp:          findSRSubHeaderIndex(headersSR, "mrp"),
      basicRate:    findSRSubHeaderIndex(headersSR, "basicRate"),
      qty:          findSRSubHeaderIndex(headersSR, "qty"),
      freeQty:      findSRSubHeaderIndex(headersSR, "freeQty"),
      schemeDisc:   findSRSubHeaderIndex(headersSR, "schemeDisc"),
      cashDisc:     findSRSubHeaderIndex(headersSR, "cashDisc"),
      totDisc:      findSRSubHeaderIndex(headersSR, "totDisc"),
      taxableValue: findSRSubHeaderIndex(headersSR, "taxableValue"),
      taxValue:     findSRSubHeaderIndex(headersSR, "taxValue"),
      gstpct:       findSRSubHeaderIndex(headersSR, "gstpct"),
      netsales:     findSRSubHeaderIndex(headersSR, "netsales"),
      grosssales:   findSRSubHeaderIndex(headersSR, "grosssales"),
    };

    const isReturnVoucherNo = (v: string): boolean => {
      if (!v) return false;
      const u = v.toUpperCase();
      return u.startsWith("SRT") || u.startsWith("SR-") || u.startsWith("SR/") || u.startsWith("CN-") || u.startsWith("CN/") || u.startsWith("RET");
    };

    let lastParty = "";
    let lastDate: unknown = null;
    let lastSalesperson = "";
    const items: SRItemParsed[] = [];

    for (const r of dataRowsSR) {
      let origBillStr = idxSR.billNumber >= 0 ? normBillNo(r[idxSR.billNumber]) : "";

      if (!origBillStr || isReturnVoucherNo(origBillStr)) {
        origBillStr = "";
        for (let c = 0; c < r.length; c++) {
          if (c === idxSR.srNumber) continue;
          const cellVal = normBillNo(r[c]);
          if (!cellVal) continue;
          if (isReturnVoucherNo(cellVal)) continue;
          const upper = cellVal.toUpperCase();
          if (upper.includes("TOTAL") || upper === "NA" || upper === "NONE" || upper === "-") continue;
          if (/^\d{3,8}$/.test(cellVal) || /^GST[\w\d\-]+$/i.test(cellVal)) {
            origBillStr = cellVal;
            break;
          }
        }
      }

      const party = (idxSR.party >= 0 && r[idxSR.party]) ? String(r[idxSR.party]).trim() : lastParty;
      if (party) lastParty = party;

      const dateVal = (idxSR.billDate >= 0 && r[idxSR.billDate])
        ? r[idxSR.billDate]
        : (idxSR.srDate >= 0 && r[idxSR.srDate])
        ? r[idxSR.srDate]
        : lastDate;
      if (dateVal) lastDate = dateVal;

      const salesperson = (idxSR.salesperson >= 0 && r[idxSR.salesperson])
        ? String(r[idxSR.salesperson]).trim()
        : lastSalesperson;
      if (salesperson) lastSalesperson = salesperson;

      const product = idxSR.product >= 0 && r[idxSR.product] ? String(r[idxSR.product]).trim() : "";
      if (!product) continue;

      const rawQty = idxSR.qty >= 0 ? num(r[idxSR.qty]) : 0;
      const rawFreeQty = idxSR.freeQty >= 0 ? num(r[idxSR.freeQty]) : 0;
      const absQty = Math.abs(rawQty) || (rawFreeQty > 0 ? Math.abs(rawFreeQty) : 0);

      items.push({
        origBillNo: origBillStr,
        srNumber: idxSR.srNumber >= 0 ? normBillNo(r[idxSR.srNumber]) : "",
        party: party || "",
        gstin: idxSR.gstin >= 0 ? cleanGSTIN(r[idxSR.gstin]) : "",
        salesperson: salesperson || "",
        product,
        hsn: idxSR.hsn >= 0 ? String(r[idxSR.hsn] ?? "").trim() : "",
        mrp: idxSR.mrp >= 0 ? Math.abs(num(r[idxSR.mrp])) : 0,
        basicRate: idxSR.basicRate >= 0 ? Math.abs(num(r[idxSR.basicRate])) : 0,
        qty: absQty,
        freeQty: rawFreeQty > 0 ? Math.abs(rawFreeQty) : 0,
        schemeDisc: idxSR.schemeDisc >= 0 ? Math.abs(num(r[idxSR.schemeDisc])) : 0,
        cashDisc: idxSR.cashDisc >= 0 ? Math.abs(num(r[idxSR.cashDisc])) : 0,
        totDisc: idxSR.totDisc >= 0 ? Math.abs(num(r[idxSR.totDisc])) : 0,
        taxableValue: idxSR.taxableValue >= 0 ? Math.abs(num(r[idxSR.taxableValue])) : 0,
        taxValue: idxSR.taxValue >= 0 ? Math.abs(num(r[idxSR.taxValue])) : 0,
        gstpct: idxSR.gstpct >= 0 ? num(r[idxSR.gstpct]) : 0,
        netsales: idxSR.netsales >= 0 ? Math.abs(num(r[idxSR.netsales])) : 0,
        grosssales: idxSR.grosssales >= 0 ? Math.abs(num(r[idxSR.grosssales])) : 0,
        dateVal,
        used: false,
      });
    }
    return items;
  } catch (e) {
    console.warn("[parseSRBufferItems] Error:", e);
    return [];
  }
}

function findMatchingSRItems(
  srItemMap: Map<string, SRItemParsed[]>,
  billNo: string,
  itemName: string
): SRItemParsed[] {
  if (srItemMap.size === 0) return [];
  const billKeys = getNormalizedBillKeys(billNo);
  const normItem = norm(itemName);

  const candidates: SRItemParsed[] = [];
  const seen = new Set<SRItemParsed>();
  for (const k of billKeys) {
    const list = srItemMap.get(k);
    if (list) {
      for (const sr of list) {
        if (!sr.used && !seen.has(sr)) {
          seen.add(sr);
          candidates.push(sr);
        }
      }
    }
  }

  if (candidates.length === 0) return [];

  const matched: SRItemParsed[] = [];

  // 1. Exact match on unused item for this bill
  for (const sr of candidates) {
    if (sr.used) continue;
    const srNormItem = norm(sr.product);
    if (srNormItem === normItem) {
      matched.push(sr);
    }
  }

  // 2. Substring / fuzzy match on product name for same bill
  if (matched.length === 0) {
    for (const sr of candidates) {
      if (sr.used) continue;
      const srNormItem = norm(sr.product);
      if (
        (srNormItem.length >= 4 && normItem.includes(srNormItem)) ||
        (normItem.length >= 4 && srNormItem.includes(normItem))
      ) {
        matched.push(sr);
      }
    }
  }

  return matched;
}

export function convertSaleData(
  sourceBuf: ArrayBuffer,
  registerDiscounts?: Map<string, RegisterDiscount>,
  srBuf?: ArrayBuffer | null,
  returnRegister?: Map<string, ReturnRegEntry> | null
): ConvertResult {
  const wb = XLSX.read(sourceBuf, { type: "array", cellDates: false, cellFormula: false, cellHTML: false, cellText: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });

  // Find header row: scan top 60 rows for best matching row containing bill / net sales / party
  let headerIdx = -1;
  let maxMatch = 0;
  const BILL_SIGNALS = ["bill number", "bill no", "invoice number", "invoice no", "net sales", "party name", "customer name", "product description"];
  for (let i = 0; i < Math.min(all.length, 60); i++) {
    const r = (all[i] || []).map((c) => norm(c));
    const matchCount = r.filter((cell) => BILL_SIGNALS.some((sig) => cell === sig || cell.includes(sig))).length;
    if (matchCount > maxMatch) {
      maxMatch = matchCount;
      headerIdx = i;
    }
  }
  if (headerIdx < 0 || maxMatch < 2) {
    throw new Error("Header row not found in Bill Wise Sales (expected 'Bill Number', 'Net Sales', 'Party Name').");
  }

  const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
  const dataRows = all.slice(headerIdx + 1).filter((r) => r.some((c) => c !== null && c !== ""));

  const idx = {
    invoice: findHeaderIndex(headers, "invoice"),
    date: findHeaderIndex(headers, "date"),
    party: findHeaderIndex(headers, "party"),
    gstin: findHeaderIndex(headers, "gstin"),
    salesperson: findHeaderIndex(headers, "salesperson"),
    item: findHeaderIndex(headers, "item"),
    hsn: findHeaderIndex(headers, "hsn"),
    units: findHeaderIndex(headers, "units"),
    freeqty: findHeaderIndex(headers, "freeqty"),
    mrp: findHeaderIndex(headers, "mrp"),
    basicrate: findHeaderIndex(headers, "basicrate"),
    schemedisc: findHeaderIndex(headers, "schemedisc"),
    rsdisc: findHeaderIndex(headers, "rsdisc"),
    totaldisc: findHeaderIndex(headers, "totaldisc"),
    gstpct: findHeaderIndex(headers, "gstpct"),
    netsales: findHeaderIndex(headers, "netsales"),
    grosssales: findHeaderIndex(headers, "grosssales"),
    taxableamt: findHeaderIndex(headers, "taxableamt"),
    taxvalue: findHeaderIndex(headers, "taxvalue"),
  };

  const mappedHeaders: Record<string, string> = {};
  for (const [k, i] of Object.entries(idx)) if (i >= 0) mappedHeaders[k] = headers[i];

  // Salesperson Name is strictly required
  if (idx.salesperson < 0) {
    throw new Error("Salesperson Name column missing hai. Yeh file Bill Wise Sales upload me support nahi karti. Kripya Salesperson Name column wali valid file upload karein.");
  }

  const required: (keyof typeof idx)[] = ["invoice", "date", "party", "salesperson", "item", "units", "netsales"];
  const missing = required.filter((k) => idx[k] < 0);
  if (missing.length)
    throw new Error("Missing required columns in Bill Wise Sales: " + missing.map((k) => `${k} (tried: ${ALIASES[k].join(", ")})`).join("; "));

  // Parse attached Sale Return items up front for 0-unit row recovery and missing bill addition
  const parsedSRItems: SRItemParsed[] = srBuf ? parseSRBufferItems(srBuf) : [];

  // Build fast indexed Map for Sale Return items by normalized bill keys
  const srItemMap = new Map<string, SRItemParsed[]>();
  for (const sr of parsedSRItems) {
    const keys = getNormalizedBillKeys(sr.origBillNo);
    for (const k of keys) {
      let list = srItemMap.get(k);
      if (!list) {
        list = [];
        srItemMap.set(k, list);
      }
      list.push(sr);
    }
  }

  const errors: { row: number; reason: string }[] = [];
  let removedNegative = 0;
  const out: Record<string, unknown>[] = [];
  const processedBillItems = new Set<string>();

  const saleBillNumbers = new Set<string>();
  let gstPrefixCount = 0;
  let totalInvoicesChecked = 0;
  let srMissingBillsAdded = 0;
  let srMissingRowsAdded = 0;
  const srMissingBillsSet = new Set<string>();

  // Fully / partially returned rows come through with Units = 0 in the Bill Wise export.
  // These get rebuilt from the Sales Return Register (bill-level) after the main pass.
  const zeroPending: {
    row: Record<string, unknown>;
    rate: number;
    billKey: string;
    invStr: string;
    itemStr: string;
    rowNum: number;
  }[] = [];
  let zeroQtyRebuilt = 0;
  let zeroQtyUnresolved = 0;

  dataRows.forEach((r, i) => {
    const rowNum = headerIdx + 2 + i;
    const netSales = num(r[idx.netsales]);
    if (netSales < 0) {
      removedNegative++;
      return;
    }
    const rawUnits = r[idx.units];
    let qty = num(rawUnits);
    const invoice = r[idx.invoice];
    const party = r[idx.party];
    const item = r[idx.item];
    if (!invoice || !party || !item) {
      errors.push({ row: rowNum, reason: "Missing invoice/party/item" });
      return;
    }

    const invStr = normBillNo(invoice);
    saleBillNumbers.add(invStr);
    saleBillNumbers.add(invStr.toLowerCase());
    const numOnly = invStr.replace(/[^0-9]/g, "").replace(/^0+/, "");
    if (numOnly) saleBillNumbers.add(numOnly);

    totalInvoicesChecked++;
    if (/^gst/i.test(invStr)) gstPrefixCount++;

    const itemStr = String(item).trim();

    // Check if there are matching return items in parsedSRItems for this bill & product
    const matchingSRList = parsedSRItems.length > 0 ? findMatchingSRItems(srItemMap, invStr, itemStr) : [];
    let retQtySum = 0;
    let retFreeQtySum = 0;
    let retDiscSum = 0;
    let srPrimary: SRItemParsed | undefined;

    if (matchingSRList.length > 0) {
      srPrimary = matchingSRList[0];
      for (const sr of matchingSRList) {
        sr.used = true;
        const q = Math.abs(sr.qty) || (sr.freeQty > 0 ? Math.abs(sr.freeQty) : 0);
        retQtySum += q;
        retFreeQtySum += (sr.freeQty > 0 ? Math.abs(sr.freeQty) : 0);
        retDiscSum += Math.round(((sr.schemeDisc || 0) + (sr.cashDisc || 0) + (sr.totDisc || 0)) * 100) / 100;
      }
    }

    // Augment quantity: Gross Sale Qty = BWS Qty + Return Qty (e.g. 2 pcs sold + 1 pc returned = 3 pcs original invoice)
    if (qty > 0 && retQtySum > 0) {
      qty += retQtySum;
    } else if (qty <= 0 && retQtySum > 0) {
      qty = retQtySum;
    }

    // Still zero → queue for rebuild from Sales Return Register value (bill-level) after this pass
    const basicRateVal = idx.basicrate >= 0 && num(r[idx.basicrate]) > 0 ? num(r[idx.basicrate]) : 0;
    const mrpRateVal = idx.mrp >= 0 && num(r[idx.mrp]) > 0 ? num(r[idx.mrp]) : 0;
    const rateForRebuild = basicRateVal > 0 ? basicRateVal : (srPrimary && srPrimary.basicRate > 0 ? srPrimary.basicRate : mrpRateVal);

    if (qty <= 0) {
      const row: Record<string, unknown> = {};
      for (const h of OUTPUT_HEADERS) row[h] = "";

      row["Invoice Number*"] = invStr;
      row["Date*"] = fmtDate(r[idx.date]) || "";
      row["Customer Name Or Alias Name*"] = String(party).trim();
      row["GSTIN"] = idx.gstin >= 0 ? cleanGSTIN(r[idx.gstin]) : "";
      row["Billing Address Country"] = "INDIA";
      row["Billing Address State"] = "GUJARAT";
      row["Billing Address City"] = "SURAT";
      row["Shipping Address Country"] = "INDIA";
      row["Shipping Address State"] = "GUJARAT";
      row["Shipping Address City"] = "SURAT";

      const brokerName = idx.salesperson >= 0 ? String(r[idx.salesperson] ?? "").trim() : "";
      row["Broker Name"] = brokerName;
      if (brokerName) {
        row["Brokerage %"] = 0;
        row["Brokerage On Value"] = "Invoice Value";
      }
      row["Ledger Name"] = "Sale";
      row["Item Name Or Alias Name Or SKU*"] = itemStr;

      const hsnVal = idx.hsn >= 0 ? cleanHSN(r[idx.hsn]) : "";
      if (hsnVal) {
        row["Hsn Code"] = hsnVal;
      }
      if (mrpRateVal > 0) row["MRP"] = mrpRateVal;
      row["Rate per Unit (Without GST)"] = rateForRebuild;

      let gstPct = idx.gstpct >= 0 ? num(r[idx.gstpct]) : 0;
      if (gstPct <= 0 && idx.taxableamt >= 0 && idx.taxvalue >= 0) {
        const taxable = num(r[idx.taxableamt]);
        const taxVal = num(r[idx.taxvalue]);
        if (taxVal > 0 && taxable > 0) gstPct = snapGST(taxVal, taxable);
      }
      row["GST %"] = gstPct || "";
      row["Discount 1 Type"] = "T";

      zeroPending.push({
        row,
        rate: rateForRebuild,
        billKey: billKeyOf(invStr),
        invStr,
        itemStr,
        rowNum,
      });
      return;
    }

    processedBillItems.add(norm(invStr) + "|||" + norm(itemStr));

    const row: Record<string, unknown> = {};
    for (const h of OUTPUT_HEADERS) row[h] = "";

    row["Invoice Number*"] = invStr;
    row["Date*"] = fmtDate(r[idx.date]) || (srPrimary ? fmtDate(srPrimary.dateVal) : "");
    row["Customer Name Or Alias Name*"] = String(party).trim() || (srPrimary ? srPrimary.party : "");
    row["GSTIN"] = (idx.gstin >= 0 ? cleanGSTIN(r[idx.gstin]) : "") || (srPrimary ? srPrimary.gstin : "");
    row["Billing Address Country"] = "INDIA";
    row["Billing Address State"] = "GUJARAT";
    row["Billing Address City"] = "SURAT";
    row["Shipping Address Country"] = "INDIA";
    row["Shipping Address State"] = "GUJARAT";
    row["Shipping Address City"] = "SURAT";

    const brokerName = (idx.salesperson >= 0 ? String(r[idx.salesperson] ?? "").trim() : "") || (srPrimary ? srPrimary.salesperson : "");
    row["Broker Name"] = brokerName;
    if (brokerName) {
      row["Brokerage %"] = 0;
      row["Brokerage On Value"] = "Invoice Value";
    }
    row["Ledger Name"] = "Sale";
    row["Item Name Or Alias Name Or SKU*"] = itemStr;

    const hsnVal = (srPrimary && srPrimary.hsn) ? cleanHSN(srPrimary.hsn) : (idx.hsn >= 0 ? cleanHSN(r[idx.hsn]) : "");
    if (hsnVal) {
      row["Hsn Code"] = hsnVal;
    }

    row["Unit"] = "";
    row["Quantity*"] = qty;
    row["Free Quantity Unit"] = "";
    row["Free Quantity"] = "";

    const mrpVal = (srPrimary && srPrimary.mrp > 0) ? srPrimary.mrp : (idx.mrp >= 0 ? num(r[idx.mrp]) : 0);
    if (mrpVal > 0) row["MRP"] = mrpVal;

    // Rate per Unit (Without GST)
    let ratePerUnit = idx.basicrate >= 0 && num(r[idx.basicrate]) > 0 ? num(r[idx.basicrate]) : 0;
    if (ratePerUnit <= 0 && srPrimary && srPrimary.basicRate > 0) {
      ratePerUnit = srPrimary.basicRate;
    }
    if (ratePerUnit <= 0 && qty > 0) {
      const gross = srPrimary && srPrimary.grosssales > 0 ? srPrimary.grosssales : (idx.grosssales >= 0 ? num(r[idx.grosssales]) : 0);
      const taxable = srPrimary && srPrimary.taxableValue > 0 ? srPrimary.taxableValue : (idx.taxableamt >= 0 ? num(r[idx.taxableamt]) : 0);
      const net = srPrimary && srPrimary.netsales > 0 ? srPrimary.netsales : (idx.netsales >= 0 ? num(r[idx.netsales]) : 0);
      if (taxable > 0) ratePerUnit = Math.round((taxable / qty) * 10000) / 10000;
      else if (gross > 0) ratePerUnit = Math.round((gross / qty) * 10000) / 10000;
      else if (net > 0) ratePerUnit = Math.round((net / qty) * 10000) / 10000;
      else if (mrpVal > 0) ratePerUnit = mrpVal;
    }
    if (ratePerUnit <= 0 && mrpVal > 0) ratePerUnit = mrpVal;
    if (ratePerUnit <= 0) ratePerUnit = 0.01;

    row["Rate per Unit (Without GST)"] = ratePerUnit;

    let gstPct = idx.gstpct >= 0 ? num(r[idx.gstpct]) : 0;
    if (gstPct <= 0 && srPrimary && srPrimary.gstpct > 0) {
      gstPct = srPrimary.gstpct;
    }
    if (gstPct <= 0) {
      const taxable = (srPrimary && srPrimary.taxableValue > 0) ? srPrimary.taxableValue : (idx.taxableamt >= 0 ? num(r[idx.taxableamt]) : (ratePerUnit * qty));
      const taxVal = (srPrimary && srPrimary.taxValue > 0) ? srPrimary.taxValue : (idx.taxvalue >= 0 ? num(r[idx.taxvalue]) : 0);
      if (taxVal > 0 && taxable > 0) gstPct = snapGST(taxVal, taxable);
    }
    row["GST %"] = gstPct || "";

    // Set Discount 1 = Total Disc when non-zero, else Scheme Disc + RS Discount + retDiscSum
    let rowDisc = 0;
    const totD = idx.totaldisc >= 0 ? Math.abs(num(r[idx.totaldisc])) : 0;
    if (totD > 0) {
      rowDisc = totD + retDiscSum;
    } else {
      const schD = idx.schemedisc >= 0 ? Math.abs(num(r[idx.schemedisc])) : 0;
      const rsD = idx.rsdisc >= 0 ? Math.abs(num(r[idx.rsdisc])) : 0;
      rowDisc = Math.round((schD + rsD + retDiscSum) * 100) / 100;
    }
    if (rowDisc > 0) {
      row["Discount 1 Type"] = "T";
      row["Discount 1"] = rowDisc;
    } else {
      row["Discount 1 Type"] = "";
      row["Discount 1"] = "";
    }
    row["Discount 2 Type"] = "";
    row["Discount 2"] = "";

    out.push(row);
    if (srPrimary) {
      srMissingRowsAdded += matchingSRList.length;
      srMissingBillsSet.add(invStr);
    }
  });

  // Rebuild pending zero-qty rows using Sales Return Register (bill-level)
  if (zeroPending.length > 0) {
    const pendingByBill = new Map<string, typeof zeroPending>();
    for (const p of zeroPending) {
      const list = pendingByBill.get(p.billKey) ?? [];
      list.push(p);
      pendingByBill.set(p.billKey, list);
    }

    for (const [bKey, pRows] of pendingByBill.entries()) {
      const regEntry = returnRegister ? returnRegister.get(bKey) : undefined;
      if (!regEntry) {
        zeroQtyUnresolved += pRows.length;
        pRows.forEach((p) => {
          errors.push({ row: p.rowNum, reason: `Quantity = 0 for bill ${p.invStr} (Sale Return data nahi mila)` });
        });
        continue;
      }

      // Allocate each bill's Ret. Value + Sch Disc across its pending zero rows (qty = value ÷ BASIC RATE, exact when a bill has one zeroed item)
      const totalAllocateVal = (regEntry.retValue || 0) + (regEntry.schDisc || 0);
      const totalSchDisc = regEntry.schDisc || 0;

      if (pRows.length === 1) {
        const p = pRows[0];
        const rate = p.rate > 0 ? p.rate : 1;
        const qtyCalc = Math.round((totalAllocateVal / rate) * 100) / 100;
        p.row["Quantity*"] = qtyCalc > 0 ? qtyCalc : 1;
        p.row["Unit"] = "";
        p.row["Free Quantity Unit"] = "";
        p.row["Free Quantity"] = "";
        p.row["Rate per Unit (Without GST)"] = p.rate > 0 ? p.rate : Math.round(totalAllocateVal * 10000) / 10000;
        const discAmt = Math.round(totalSchDisc * 100) / 100;
        if (discAmt > 0) {
          p.row["Discount 1 Type"] = "T";
          p.row["Discount 1"] = discAmt;
        } else {
          p.row["Discount 1 Type"] = "";
          p.row["Discount 1"] = "";
        }
        out.push(p.row);
        processedBillItems.add(norm(p.invStr) + "|||" + norm(p.itemStr));
        zeroQtyRebuilt++;
      } else {
        // Multiple zero-qty rows for this bill: allocate proportionally
        const totalRates = pRows.reduce((s, p) => s + (p.rate > 0 ? p.rate : 1), 0);
        for (const p of pRows) {
          const rate = p.rate > 0 ? p.rate : 1;
          const weight = totalRates > 0 ? rate / totalRates : 1 / pRows.length;
          const itemVal = totalAllocateVal * weight;
          const itemDisc = totalSchDisc * weight;
          const qtyCalc = Math.round((itemVal / rate) * 100) / 100;
          p.row["Quantity*"] = qtyCalc > 0 ? qtyCalc : 1;
          p.row["Unit"] = "";
          p.row["Free Quantity Unit"] = "";
          p.row["Free Quantity"] = "";
          p.row["Rate per Unit (Without GST)"] = p.rate > 0 ? p.rate : Math.round(itemVal * 10000) / 10000;
          const discAmt = Math.round(itemDisc * 100) / 100;
          if (discAmt > 0) {
            p.row["Discount 1 Type"] = "T";
            p.row["Discount 1"] = discAmt;
          } else {
            p.row["Discount 1 Type"] = "";
            p.row["Discount 1"] = "";
          }
          out.push(p.row);
          processedBillItems.add(norm(p.invStr) + "|||" + norm(p.itemStr));
          zeroQtyRebuilt++;
        }
      }
    }
  }

  const prefersGstPrefix = totalInvoicesChecked > 0 && (gstPrefixCount / totalInvoicesChecked) >= 0.5;

  // Add any remaining unused rows from Sale Return file that belong to valid invoices
  const isReturnVoucherNo = (v: string): boolean => {
    if (!v) return false;
    const u = v.toUpperCase();
    return u.startsWith("SRT") || u.startsWith("SR-") || u.startsWith("SR/") || u.startsWith("CN-") || u.startsWith("CN/") || u.startsWith("RET");
  };

  parsedSRItems.forEach((srItem) => {
    if (srItem.used) return;
    const origBillStr = srItem.origBillNo;
    if (!origBillStr || isReturnVoucherNo(origBillStr)) return;

    // Strict Bill-Wise Match: Only process SR items for bills that exist in Bill-Wise Sales
    const origKeys = getNormalizedBillKeys(origBillStr);
    const billInBWS = origKeys.some((k) => saleBillNumbers.has(k) || saleBillNumbers.has(k.toLowerCase()));
    if (!billInBWS) return;

    let finalInvoiceNo = origBillStr;
    if (/^\d+$/.test(origBillStr) && prefersGstPrefix && !finalInvoiceNo.toUpperCase().startsWith("GST")) {
      finalInvoiceNo = "GST" + origBillStr;
    }

    const prodStr = srItem.product.trim();
    if (!prodStr) return;

    const itemKey = norm(finalInvoiceNo) + "|||" + norm(prodStr);
    const itemKeyOrig = norm(origBillStr) + "|||" + norm(prodStr);
    if (processedBillItems.has(itemKey) || processedBillItems.has(itemKeyOrig)) {
      srItem.used = true;
      return;
    }

    const absQty = Math.abs(srItem.qty) || (srItem.freeQty > 0 ? Math.abs(srItem.freeQty) : 0) || 1;

    const row: Record<string, unknown> = {};
    for (const h of OUTPUT_HEADERS) row[h] = "";

    row["Invoice Number*"] = finalInvoiceNo;
    row["Date*"] = fmtDate(srItem.dateVal);
    row["Customer Name Or Alias Name*"] = srItem.party || "Customer";
    row["GSTIN"] = srItem.gstin;
    row["Billing Address Country"] = "INDIA";
    row["Billing Address State"] = "GUJARAT";
    row["Billing Address City"] = "SURAT";
    row["Shipping Address Country"] = "INDIA";
    row["Shipping Address State"] = "GUJARAT";
    row["Shipping Address City"] = "SURAT";

    row["Broker Name"] = srItem.salesperson;
    if (srItem.salesperson) {
      row["Brokerage %"] = 0;
      row["Brokerage On Value"] = "Invoice Value";
    }
    row["Ledger Name"] = "Sale";
    row["Item Name Or Alias Name Or SKU*"] = prodStr;
    const hsnVal = srItem.hsn ? cleanHSN(srItem.hsn) : "";
    if (hsnVal) {
      row["Hsn Code"] = hsnVal;
    }
    row["Unit"] = "";
    row["Quantity*"] = absQty;
    row["Free Quantity Unit"] = "";
    row["Free Quantity"] = "";
    if (srItem.mrp > 0) row["MRP"] = srItem.mrp;

    let ratePerUnit = srItem.basicRate > 0 ? srItem.basicRate : 0;
    if (ratePerUnit <= 0 && absQty > 0) {
      if (srItem.taxableValue > 0) ratePerUnit = Math.round((srItem.taxableValue / absQty) * 10000) / 10000;
      else if (srItem.grosssales > 0) ratePerUnit = Math.round((srItem.grosssales / absQty) * 10000) / 10000;
      else if (srItem.netsales > 0) ratePerUnit = Math.round((srItem.netsales / absQty) * 10000) / 10000;
      else if (srItem.mrp > 0) ratePerUnit = srItem.mrp;
    }
    if (ratePerUnit <= 0 && srItem.mrp > 0) ratePerUnit = srItem.mrp;
    if (ratePerUnit <= 0) ratePerUnit = 0.01;

    row["Rate per Unit (Without GST)"] = ratePerUnit;

    let gstPct = srItem.gstpct > 0 ? srItem.gstpct : 0;
    if (gstPct <= 0 && srItem.taxValue > 0 && srItem.taxableValue > 0) {
      gstPct = snapGST(srItem.taxValue, srItem.taxableValue);
    }
    row["GST %"] = gstPct || "";

    const totalDisc = Math.round(((srItem.schemeDisc || 0) + (srItem.cashDisc || 0) + (srItem.totDisc || 0)) * 100) / 100;
    if (totalDisc > 0) {
      row["Discount 1 Type"] = "T";
      row["Discount 1"] = totalDisc;
    } else {
      row["Discount 1 Type"] = "";
      row["Discount 1"] = "";
    }
    row["Discount 2 Type"] = "";
    row["Discount 2"] = "";

    out.push(row);
    processedBillItems.add(itemKey);
    srItem.used = true;
    srMissingRowsAdded++;
    srMissingBillsSet.add(finalInvoiceNo);
  });

  srMissingBillsAdded = srMissingBillsSet.size;

  // Aggregate duplicate rows with the same item name in the same invoice (Weighted Average Base Rate)
  const mergedOut = mergeSameItemRowsInBill(out, "Invoice Number*", "Item Name Or Alias Name Or SKU*");

  // Natural alphabetical & numerical sequential sort by Invoice Number
  const uniqueBills = Array.from(new Set(mergedOut.map((r) => String(r["Invoice Number*"] ?? ""))));
  uniqueBills.sort((a, b) => {
    const numA = parseInt(a.replace(/[^0-9]/g, ""), 10);
    const numB = parseInt(b.replace(/[^0-9]/g, ""), 10);
    if (!isNaN(numA) && !isNaN(numB) && numA !== numB) return numA - numB;
    return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
  });

  const billOrderMap = new Map<string, number>();
  uniqueBills.forEach((b, idx) => billOrderMap.set(b, idx));

  mergedOut.sort((a, b) => {
    const invA = String(a["Invoice Number*"] ?? "");
    const invB = String(b["Invoice Number*"] ?? "");
    const orderA = billOrderMap.get(invA) ?? 0;
    const orderB = billOrderMap.get(invB) ?? 0;
    return orderA - orderB;
  });

  // POST-SORT DISCOUNT PLACEMENT: Place discounts on the FIRST row of each unique invoice
  const billDiscountPlaced = new Set<string>();
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

  if (registerDiscounts) {
    for (const row of mergedOut) {
      const inv = String(row["Invoice Number*"] ?? "").trim();
      if (!inv) continue;

      const numKey = inv.replace(/[^0-9]/g, "").replace(/^0+/, "") || inv;
      if (billDiscountPlaced.has(numKey)) continue;

      const disc = lookupDiscount(registerDiscounts, inv);
      if (disc) {
        const addLess1 = disc.totalAddLess1 !== undefined
          ? disc.totalAddLess1
          : ((disc.ushop || 0) + (disc.btpr || 0));

        if (addLess1 > 0) {
          row["Add / less 1 Ledger"] = "USHOP DIS";
          row["Add / less 1 Type"]   = "₹";
          row["Add / less 1 Amount"] = -Math.round(addLess1 * 100) / 100;
        }
        if (disc.shikhar > 0) {
          row["Add / less 2 Ledger"] = "SHIKHAR DIS";
          row["Add / less 2 Type"]   = "₹";
          row["Add / less 2 Amount"] = -Math.round(disc.shikhar * 100) / 100;
        }
        if (disc.roundOff !== undefined && disc.roundOff !== 0) {
          row["Round off Amount"] = Math.round(disc.roundOff * 100) / 100;
        }

        discountBillsMatched++;
        discountBills.push({
          bill: inv,
          ushop: disc.ushop,
          shikhar: disc.shikhar,
          btpr: disc.btpr,
          cash: disc.cash ?? 0,
          adjustments: disc.adjustments ?? 0,
          totalAddLess1: addLess1,
          roundOff: disc.roundOff ?? 0,
        });
      }
      billDiscountPlaced.add(numKey);
    }
  }

  return {
    rows: mergedOut,
    stats: {
      sourceRows: dataRows.length,
      exportRows: out.length,
      removedNegative,
      rejectedInvalid: errors.length,
      discountBillsMatched,
      discountBills,
      errors,
      mappedHeaders,
      srMissingBillsAdded,
      srMissingRowsAdded,
      zeroQtyRebuilt,
      zeroQtyUnresolved,
    },
  };
}

const NUM_FMT_COLS = ["MRP", "Discount 1", "Add / less 1 Amount", "Add / less 2 Amount", "Round off Amount"];
const NUM_FMT = "0.00";
const RATE_FMT_COL = "Rate per Unit (Without GST)";
const RATE_FMT = "0.0000";

function applyNumFormats(ws: XLSX.WorkSheet, headers: string[], numDataRows: number) {
  const colIndices: { c: number; fmt: string }[] = [];
  for (const col of NUM_FMT_COLS) {
    const c = headers.indexOf(col);
    if (c >= 0) colIndices.push({ c, fmt: NUM_FMT });
  }
  const rateCol = headers.indexOf(RATE_FMT_COL);
  if (rateCol >= 0) colIndices.push({ c: rateCol, fmt: RATE_FMT });

  if (colIndices.length === 0) return;

  for (let row = 1; row <= numDataRows; row++) {
    for (let ci = 0; ci < colIndices.length; ci++) {
      const { c, fmt } = colIndices[ci];
      const addr = XLSX.utils.encode_cell({ r: row, c });
      const cell = ws[addr];
      if (cell && cell.v !== "" && cell.v !== null && cell.v !== undefined) {
        cell.t = "n";
        cell.z = fmt;
      }
    }
  }
}

export function buildOutputWorkbook(rows: Record<string, unknown>[]): ArrayBuffer {
  const data = [OUTPUT_HEADERS, ...rows.map((r) => OUTPUT_HEADERS.map((h) => r[h] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(data);
  applyNumFormats(ws, OUTPUT_HEADERS, rows.length);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sale Item");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

const SPLIT_MAX = 4000;

function splitAtBillBoundary(
  rows: Record<string, unknown>[],
  billKey: string,
  targetSize = SPLIT_MAX,
): Record<string, unknown>[][] {
  if (rows.length <= targetSize) return [rows];
  const chunks: Record<string, unknown>[][] = [];
  let remaining = rows;

  const getBill = (row: Record<string, unknown>) => String(row[billKey] ?? "").trim();

  while (remaining.length > targetSize) {
    let splitIdx = targetSize;

    // Check if targetSize falls inside a bill (same bill as preceding row)
    if (getBill(remaining[splitIdx]) === getBill(remaining[splitIdx - 1])) {
      const currentBill = getBill(remaining[splitIdx]);

      // Scan backwards to find the first row of currentBill
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

      // Pick whichever boundary is closer to ~4000 rows
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

export function buildOutputWorkbooks(rows: Record<string, unknown>[]): ArrayBuffer[] {
  const chunks = splitAtBillBoundary(rows, "Invoice Number*");
  return chunks.map((chunk) => buildOutputWorkbook(chunk));
}

export function todayFolderName(d: Date = new Date()): string {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${d.getFullYear()}-${mm}-${dd}`;
}
