import * as XLSX from "xlsx";

const num = (v: unknown): number => {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[₹,\s]/g, ""));
  return isNaN(n) ? 0 : n;
};

const norm = (s: unknown) =>
  String(s ?? "").toLowerCase().replace(/[^a-z0-9%]/g, " ").replace(/\s+/g, " ").trim();

export function normBillNo(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "number") return String(Math.round(v));
  return String(v).trim();
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
    const s = v.trim();
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
    const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
    if (dmy) {
      const y = dmy[3].length === 2 ? "20" + dmy[3] : dmy[3];
      return `${dmy[1].padStart(2, "0")}/${dmy[2].padStart(2, "0")}/${y}`;
    }
    return s;
  }
  return String(v ?? "").trim();
}

export interface SaleReportRow {
  partyName: string;
  docNo: string;
  date: string;
  taxableVal?: number;
  cgst?: number;
  sgst?: number;
  igst?: number;
  addLess?: number;
  roundOff?: number;
  invoiceAmt: number;
  paidAmt?: number;
  gstin?: string;
  brokerName?: string;
  user?: string;
  partyCode?: string;
}

export interface LeverSaleRow {
  billRefNo: string;
  billValue: number;
  partyName: string;
}

export interface LeverSaleRetRow {
  salesRetNo: string;
  total: number;
  partyName: string;
}

export type MatchStatus = "matched" | "diff" | "only_report" | "only_lever";

export interface MatchRow {
  docNo: string;
  partyName: string;
  reportAmt: number | null;
  leverAmt: number | null;
  diff: number;
  status: MatchStatus;
}

export interface MatchingResult {
  saleRows: MatchRow[];
  saleRetRows: MatchRow[];
  reportSaleTotal: number;
  reportSaleRetTotal: number;
  leverSaleTotal: number;
  leverSaleRetTotal: number;
}

// Aliases for deep scanning Hisab Kitab Sale Report columns
const HISAB_KITAB_ALIASES: Record<string, string[]> = {
  partyName: [
    "party name", "customer name", "party", "customer", "account name", "ledger name",
    "party/customer", "buyer name", "client name", "party name or alias name"
  ],
  gstin: [
    "gstin", "party gstin", "party gstin number", "gst no", "gst number", "party gst no", "party gst number"
  ],
  invoiceNumber: [
    "document number", "document no", "document no.", "doc number", "doc no", "doc no.",
    "invoice number", "invoice no", "bill number", "bill no", "voucher no",
    "billno", "invoiceno", "bill no.", "invoice no.", "inv no", "inv number",
    "invoice number*", "bill num", "credit note no", "cn no", "return no", "bill ref no", "billrefno"
  ],
  invoiceDate: [
    "invoice date", "date", "bill date", "doc date", "sales date", "bill date.",
    "invoice date.", "date*", "return date", "credit note date", "voucher date", "document date"
  ],
  brokerName: [
    "broker name", "broker", "salesperson name", "sales person name", "sales person",
    "salesperson", "salesman name", "sales man name", "salesman", "sales man",
    "agent name", "agent", "rep name", "sales representative", "executive name"
  ],
  brokerage: [
    "brokerage", "brokerage %", "brokerage rate"
  ],
  irnNumber: [
    "irn number", "irn no", "irn"
  ],
  taxableValue: [
    "taxable value", "taxable amount", "taxable amt", "taxable"
  ],
  cgst: [
    "cgst", "cgst amount", "central tax", "cgst amt", "cgst value"
  ],
  sgst: [
    "sgst", "sgst amount", "state tax", "utgst", "sgst amt", "sgst value"
  ],
  igst: [
    "igst", "igst amount", "integrated tax", "igst amt", "igst value"
  ],
  addLess: [
    "add / less amount", "add/less amount", "add / less", "add/less", "add less amount",
    "add less", "adjustment", "adjustments", "adj amt", "adj amount", "other discount"
  ],
  roundOff: [
    "round off", "round off amt", "round off amount", "roundoff", "round-off", "rnd off", "rnd off amt"
  ],
  invoiceAmount: [
    "invoice amount", "invoice amt", "net amount", "bill amount",
    "bill amt", "total amount", "total amt", "bill value", "billvalue", "grand total", "total", "net",
    "amount", "invoice value", "net return", "gross amount"
  ],
  paidAmount: [
    "paid amount", "paid amt", "payment amount", "received amount", "paid"
  ],
  user: [
    "user", "created by", "user name", "admin"
  ],
  partyCode: [
    "party code", "customer code", "party id", "customer id"
  ],
  voucherType: [
    "transaction type", "voucher type", "doc type", "type", "voucher"
  ],
};

function findColIndex(headers: string[], key: keyof typeof HISAB_KITAB_ALIASES): number {
  const aliases = HISAB_KITAB_ALIASES[key].map(norm);
  const normHeaders = headers.map(norm);

  // Exact match
  for (let i = 0; i < normHeaders.length; i++) {
    if (aliases.includes(normHeaders[i])) return i;
  }
  // Substring match - ensure we don't cross match taxes or discounts to invoice amount
  for (let i = 0; i < normHeaders.length; i++) {
    const h = normHeaders[i];
    if (!h) continue;
    if (key === "invoiceNumber" && (h.includes("type") || h.includes("date") || h.includes("party") || h.includes("name") || h.includes("amount") || h.includes("amt") || h.includes("tax") || h.includes("disc") || h.includes("gstin"))) {
      continue;
    }
    if (key === "invoiceAmount" && (h.includes("tax") || h.includes("cgst") || h.includes("sgst") || h.includes("igst") || h.includes("disc") || h.includes("add") || h.includes("less") || h.includes("round"))) {
      continue;
    }
    for (const a of aliases) {
      if (h === a || (h.length >= 4 && h.includes(a)) || (a.length >= 4 && a.includes(h))) return i;
    }
  }
  return -1;
}

/**
 * Deep scan & parse Hisab Kitab Sale Report:
 * Skips the top 14 rows (or deep scans from row 14 / index 14 onwards) to locate
 * the column header row, then maps all columns by their exact or alias names.
 */
export function parseSaleReport(buf: ArrayBuffer): { saleRows: SaleReportRow[]; saleRetRows: SaleReportRow[] } {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: false, raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as unknown[][];

  if (!all || all.length === 0) {
    return { saleRows: [], saleRetRows: [] };
  }

  // Key signals to identify the Hisab Kitab header row
  const HEADER_SIGNALS = [
    "party name", "invoice number", "invoice date", "taxable value", "invoice amount",
    "round off", "add / less", "cgst", "sgst", "igst", "gstin", "broker name", "paid amount"
  ];

  // Search for the header row: start checking from row 14 (index 14) and all surrounding rows up to row 40
  let headerIdx = -1;
  let maxMatches = 0;

  // Scan rows from 0 to 40
  for (let i = 0; i < Math.min(all.length, 50); i++) {
    const rowCells = (all[i] || []).map((c) => norm(c));
    const matches = rowCells.filter((cell) =>
      HEADER_SIGNALS.some((sig) => cell === sig || cell.includes(sig))
    ).length;

    // Prefer row index 14 if tied or high match
    if (matches > maxMatches) {
      maxMatches = matches;
      headerIdx = i;
    }
  }

  // If deep scan found header row, use it; otherwise fallback to row index 14 (15th row)
  if (headerIdx < 0 || maxMatches < 2) {
    headerIdx = all.length > 14 ? 14 : 0;
  }

  const rawHeaders = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));
  const headers = rawHeaders.map((c) => String(c ?? "").trim());

  const idx = {
    partyName:     findColIndex(headers, "partyName"),
    gstin:         findColIndex(headers, "gstin"),
    invoiceNumber: findColIndex(headers, "invoiceNumber"),
    invoiceDate:   findColIndex(headers, "invoiceDate"),
    brokerName:    findColIndex(headers, "brokerName"),
    brokerage:     findColIndex(headers, "brokerage"),
    irnNumber:     findColIndex(headers, "irnNumber"),
    taxableValue:  findColIndex(headers, "taxableValue"),
    cgst:          findColIndex(headers, "cgst"),
    sgst:          findColIndex(headers, "sgst"),
    igst:          findColIndex(headers, "igst"),
    addLess:       findColIndex(headers, "addLess"),
    roundOff:      findColIndex(headers, "roundOff"),
    invoiceAmount: findColIndex(headers, "invoiceAmount"),
    paidAmount:    findColIndex(headers, "paidAmount"),
    user:          findColIndex(headers, "user"),
    partyCode:     findColIndex(headers, "partyCode"),
    voucherType:   findColIndex(headers, "voucherType"),
  };

  // Fallback if specific indices are -1 based on standard Hisab Kitab column order (0 to 18)
  if (idx.partyName < 0) idx.partyName = 0;
  if (idx.invoiceNumber < 0) {
    const docCol = headers.findIndex((h) => {
      const n = norm(h);
      return (n.includes("document") || n.includes("doc") || n.includes("inv") || n.includes("bill")) &&
        !n.includes("type") && !n.includes("date") && !n.includes("amt") && !n.includes("amount") && !n.includes("party");
    });
    idx.invoiceNumber = docCol >= 0 ? docCol : (idx.gstin === 1 ? 2 : 2);
  }
  if (idx.invoiceDate < 0) idx.invoiceDate = idx.invoiceNumber === 2 ? 3 : 2;
  if (idx.invoiceAmount < 0) idx.invoiceAmount = 15;

  const dataRows = all.slice(headerIdx + 1).filter((r) => r.some((c) => c !== null && c !== ""));

  const saleRows: SaleReportRow[] = [];
  const saleRetRows: SaleReportRow[] = [];

  const INVALID_DOC_STRINGS = new Set(["sale", "sales", "sales return", "salesreturn", "return", "credit note", "creditnote", "total", "grand total", "sub total"]);

  for (let i = 0; i < dataRows.length; i++) {
    const r = dataRows[i] || [];
    const partyName = idx.partyName >= 0 ? String(r[idx.partyName] ?? "").trim() : "";
    let docNo = idx.invoiceNumber >= 0 ? normBillNo(r[idx.invoiceNumber]) : "";

    // If docNo is empty or resolved to a transaction type string, search row for real invoice-like string
    if (!docNo || INVALID_DOC_STRINGS.has(norm(docNo))) {
      docNo = "";
      for (let c = 0; c < r.length; c++) {
        if (c === idx.voucherType || c === idx.partyName) continue;
        const val = normBillNo(r[c]);
        if (val && !INVALID_DOC_STRINGS.has(norm(val))) {
          if (/^GST[\w\d\-]+$/i.test(val) || /^SRT[\w\d\-]+$/i.test(val) || /^\d{3,8}$/.test(val) || /^CN[\w\d\-]+$/i.test(val)) {
            docNo = val;
            break;
          }
        }
      }
    }

    if (!docNo) continue;

    const docNoNorm = norm(docNo);
    if (
      docNoNorm === "" ||
      INVALID_DOC_STRINGS.has(docNoNorm) ||
      docNoNorm.includes("grand total") ||
      docNoNorm.includes("total amount") ||
      docNoNorm.includes("sub total") ||
      docNoNorm === "total"
    ) {
      continue;
    }

    const partyNorm = norm(partyName);
    if (partyNorm.includes("grand total") || partyNorm === "total") continue;

    const date = idx.invoiceDate >= 0 ? fmtDate(r[idx.invoiceDate]) : "";
    const taxableVal = idx.taxableValue >= 0 ? num(r[idx.taxableValue]) : 0;
    const cgst = idx.cgst >= 0 ? num(r[idx.cgst]) : 0;
    const sgst = idx.sgst >= 0 ? num(r[idx.sgst]) : 0;
    const igst = idx.igst >= 0 ? num(r[idx.igst]) : 0;
    const addLess = idx.addLess >= 0 ? num(r[idx.addLess]) : 0;
    const roundOff = idx.roundOff >= 0 ? num(r[idx.roundOff]) : 0;

    let invoiceAmt = 0;
    if (idx.invoiceAmount >= 0 && r[idx.invoiceAmount] !== null && r[idx.invoiceAmount] !== undefined && r[idx.invoiceAmount] !== "") {
      invoiceAmt = num(r[idx.invoiceAmount]);
    } else if (taxableVal > 0) {
      invoiceAmt = Math.round((taxableVal + cgst + sgst + igst + addLess + roundOff) * 100) / 100;
    }

    const row: SaleReportRow = {
      partyName,
      docNo,
      date,
      taxableVal,
      cgst,
      sgst,
      igst,
      addLess,
      roundOff,
      invoiceAmt,
      paidAmt: idx.paidAmount >= 0 ? num(r[idx.paidAmount]) : 0,
      gstin: idx.gstin >= 0 ? String(r[idx.gstin] ?? "").trim() : "",
      brokerName: idx.brokerName >= 0 ? String(r[idx.brokerName] ?? "").trim() : "",
      user: idx.user >= 0 ? String(r[idx.user] ?? "").trim() : "",
      partyCode: idx.partyCode >= 0 ? String(r[idx.partyCode] ?? "").trim() : "",
    };

    const docUpper = docNo.toUpperCase();
    const vType = idx.voucherType >= 0 ? String(r[idx.voucherType] ?? "").toLowerCase() : "";

    const isReturn =
      docUpper.startsWith("SRT") ||
      docUpper.startsWith("SR-") ||
      docUpper.startsWith("SR/") ||
      docUpper.startsWith("CN") ||
      docUpper.startsWith("RET") ||
      docUpper.startsWith("CR-") ||
      vType.includes("return") ||
      vType.includes("credit note") ||
      invoiceAmt < 0;

    if (isReturn) {
      row.invoiceAmt = -Math.abs(invoiceAmt);
      saleRetRows.push(row);
    } else {
      row.invoiceAmt = Math.abs(invoiceAmt);
      saleRows.push(row);
    }
  }

  return { saleRows, saleRetRows };
}

/**
 * Deep scan & parse LeverEDGE Sales Register file.
 * - Extracts ONLY + (positive) Bill Value for Sales.
 * - Extracts - (negative) Bill Value or Sal Ret / SRT rows for Sale Returns.
 */
export function parseLeverSalesReg(buf: ArrayBuffer): {
  saleRows: LeverSaleRow[];
  saleRetRows: LeverSaleRetRow[];
} {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: false, raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as unknown[][];

  if (!all || all.length === 0) return { saleRows: [], saleRetRows: [] };

  const HEADER_SIGNALS = [
    "bill no", "bill number", "invoice no", "invoice number", "bill ref no", "billrefno",
    "bill value", "billvalue", "net amount", "party name", "customer name", "sal ret", "salret"
  ];

  let headerIdx = -1;
  let maxMatches = 0;
  for (let i = 0; i < Math.min(all.length, 60); i++) {
    const rowCells = (all[i] || []).map((c) => norm(c));
    const matches = rowCells.filter((cell) =>
      HEADER_SIGNALS.some((sig) => cell === sig || cell.includes(sig))
    ).length;
    if (matches > maxMatches) {
      maxMatches = matches;
      headerIdx = i;
    }
  }

  if (headerIdx < 0) headerIdx = 12;

  const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));

  const findCol = (...aliases: string[]) => {
    const normAliases = aliases.map(norm);
    const normHeaders = headers.map(norm);
    for (let i = 0; i < normHeaders.length; i++) {
      if (normAliases.includes(normHeaders[i])) return i;
    }
    for (let i = 0; i < normHeaders.length; i++) {
      const h = normHeaders[i];
      if (!h) continue;
      if (h.includes("tax") || h.includes("cgst") || h.includes("sgst") || h.includes("igst") || h.includes("disc") || h.includes("rate") || h.includes("mrp") || h.includes("qty")) {
        continue;
      }
      for (const a of normAliases) {
        if (h === a || (h.length >= 4 && h.includes(a)) || (a.length >= 4 && a.includes(h))) return i;
      }
    }
    return -1;
  };

  const billNoCol  = findCol("billrefno", "bill ref no", "bill no", "bill number", "invoice no", "invoice number", "bill ref. no.", "bill no.", "invoice no.");
  const billValCol = findCol("billvalue", "bill value", "bill val", "billval", "invoice amount", "bill amt", "net amount", "total");
  const partyCol   = findCol("party name", "customer name", "account name", "ledger name", "party", "customer", "party/customer", "buyer name");
  const salRetCol  = findCol("sal ret", "salret", "sales ret", "sales return", "return amt");

  const actualBillNoCol = billNoCol >= 0 ? billNoCol : 3;
  const actualBillValCol = billValCol >= 0 ? billValCol : 34;
  const actualPartyCol = partyCol >= 0 ? partyCol : 6;
  const actualSalRetCol = salRetCol >= 0 ? salRetCol : -1;

  const dataRows = all.slice(headerIdx + 1).filter((r) => r.some((c) => c !== null && c !== ""));
  const saleMap = new Map<string, { total: number; partyName: string }>();
  const retMap = new Map<string, { total: number; partyName: string }>();

  for (const r of dataRows) {
    const billRefNo = normBillNo(r[actualBillNoCol]);
    if (!billRefNo) continue;

    const normNo = norm(billRefNo);
    if (normNo.includes("grand total") || normNo.includes("total") || normNo === "total") continue;

    const billValue = num(r[actualBillValCol]);
    const salRetVal = actualSalRetCol >= 0 ? num(r[actualSalRetCol]) : 0;
    const partyName = actualPartyCol >= 0 ? String(r[actualPartyCol] ?? "").trim() : "";
    const docUpper = billRefNo.toUpperCase();
    const isExplicitReturn = docUpper.startsWith("SRT") || docUpper.startsWith("SR-") || docUpper.startsWith("CN") || docUpper.startsWith("RET");

    if (billValue > 0 && !isExplicitReturn) {
      // ONLY positive (+) Bill Value for Sales!
      const ex = saleMap.get(billRefNo);
      if (ex) {
        ex.total = Math.round((ex.total + billValue) * 100) / 100;
        if (!ex.partyName && partyName) ex.partyName = partyName;
      } else {
        saleMap.set(billRefNo, { total: Math.round(billValue * 100) / 100, partyName });
      }
    } else if (billValue < 0 || salRetVal > 0 || isExplicitReturn) {
      // Negative (-) amount for Sale Returns!
      const retAmt = billValue < 0 ? -Math.abs(billValue) : -Math.abs(salRetVal || billValue);
      const ex = retMap.get(billRefNo);
      if (ex) {
        ex.total = Math.round((ex.total + retAmt) * 100) / 100;
        if (!ex.partyName && partyName) ex.partyName = partyName;
      } else {
        retMap.set(billRefNo, { total: Math.round(retAmt * 100) / 100, partyName });
      }
    }
  }

  const saleRows: LeverSaleRow[] = Array.from(saleMap.entries()).map(([billRefNo, { total, partyName }]) => ({
    billRefNo,
    billValue: total,
    partyName,
  }));

  const saleRetRows: LeverSaleRetRow[] = Array.from(retMap.entries()).map(([salesRetNo, { total, partyName }]) => ({
    salesRetNo,
    total,
    partyName,
  }));

  return { saleRows, saleRetRows };
}

/**
 * Deep scan & parse LeverEDGE Sales Return Register file (or Sales Return Detail report).
 */
export function parseLeverSalesRetReg(buf: ArrayBuffer): LeverSaleRetRow[] {
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: false, raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) as unknown[][];

  if (!all || all.length === 0) return [];

  const HEADER_SIGNALS = [
    "sales return no", "sr number", "sr no", "sr_no", "sales return date",
    "return no", "credit note", "credit note no", "bill number", "bill no", "total",
    "net credit value", "sales return value"
  ];

  let headerIdx = -1;
  let maxMatches = 0;
  for (let i = 0; i < Math.min(all.length, 60); i++) {
    const rowCells = (all[i] || []).map((c) => norm(c));
    const matches = rowCells.filter((cell) =>
      HEADER_SIGNALS.some((sig) => cell === sig || cell.includes(sig))
    ).length;
    if (matches > maxMatches) {
      maxMatches = matches;
      headerIdx = i;
    }
  }

  if (headerIdx < 0) headerIdx = 11;

  const headers = (all[headerIdx] as unknown[]).map((c) => String(c ?? ""));

  const findCol = (...aliases: string[]) => {
    const normAliases = aliases.map(norm);
    const normHeaders = headers.map(norm);
    for (let i = 0; i < normHeaders.length; i++) {
      if (normAliases.includes(normHeaders[i])) return i;
    }
    for (let i = 0; i < normHeaders.length; i++) {
      const h = normHeaders[i];
      if (!h) continue;
      for (const a of normAliases) {
        if (h === a || (h.length >= 4 && h.includes(a)) || (a.length >= 4 && a.includes(h))) return i;
      }
    }
    return -1;
  };

  const srCol    = findCol("sales return no", "sr number", "sr no", "sr_no", "sr no.", "return no", "credit note no", "cn no", "voucher no", "bill number", "bill no");
  const totalCol = findCol("net credit value after discount reversal", "sales return value", "net credit value", "net amount", "bill value", "net return", "return amt", "credit note amt", "amount", "total");
  const partyCol = findCol("party name", "customer name", "salesperson", "sales person", "salesman", "party", "customer");

  const actualSrCol = srCol >= 0 ? srCol : 1;
  const actualTotalCol = totalCol >= 0 ? totalCol : (headers.length > 25 ? 25 : 12);
  const actualPartyCol = partyCol >= 0 ? partyCol : (headers.length > 6 ? 6 : 1);

  const dataRows = all.slice(headerIdx + 1).filter((r) => r.some((c) => c !== null && c !== ""));
  const retMap = new Map<string, { total: number; partyName: string }>();

  for (const r of dataRows) {
    const salesRetNo = normBillNo(r[actualSrCol]);
    if (!salesRetNo) continue;

    const normNo = norm(salesRetNo);
    if (normNo.includes("grand total") || normNo.includes("total") || normNo === "total") continue;

    const rawVal = num(r[actualTotalCol]);
    const total = -Math.abs(rawVal); // Always store as negative (-) amount for return matching!
    const partyName = actualPartyCol >= 0 ? String(r[actualPartyCol] ?? "").trim() : "";

    const ex = retMap.get(salesRetNo);
    if (ex) {
      ex.total = Math.round((ex.total + total) * 100) / 100;
      if (!ex.partyName && partyName) ex.partyName = partyName;
    } else {
      retMap.set(salesRetNo, { total: Math.round(total * 100) / 100, partyName });
    }
  }

  return Array.from(retMap.entries()).map(([salesRetNo, { total, partyName }]) => ({
    salesRetNo,
    total,
    partyName,
  }));
}

/**
 * Normalized key generator for flexible invoice & SRT number matching
 * (e.g. GST29471 matches 29471, GST-29471; SRT01959 matches 01959, SRT-01959, 1959).
 */
function getNormalizedKeys(docNo: string): string[] {
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

  const numOnly = s.replace(/[^0-9]/g, "").replace(/^0+/, "");
  if (numOnly) {
    keys.add(numOnly);
    keys.add("GST" + numOnly);
    keys.add("gst" + numOnly);
    keys.add("SRT" + numOnly);
    keys.add("srt" + numOnly);
  }

  if (/^(gst|srt|cn|ret|sr)/i.test(s)) {
    const withoutPrefix = s.replace(/^(gst|srt|cn|ret|sr)[\/\-_ ]*/i, "").trim();
    if (withoutPrefix) {
      keys.add(withoutPrefix);
      const withoutPrefixNoZero = withoutPrefix.replace(/^0+/, "");
      if (withoutPrefixNoZero) {
        keys.add(withoutPrefixNoZero);
        keys.add("SRT" + withoutPrefixNoZero);
        keys.add("GST" + withoutPrefixNoZero);
      }
    }
  }

  return Array.from(keys);
}

interface MatchItem {
  docNo: string;
  partyName: string;
  reportAmt: number | null;
  leverAmt: number | null;
}

function matchDataset(
  reportItems: { docNo: string; partyName: string; amount: number }[],
  leverItems: { docNo: string; partyName: string; amount: number }[]
): { rows: MatchRow[]; reportTotal: number; leverTotal: number } {
  // Aggregate report items by normalized keys
  const reportByDoc = new Map<string, { originalDoc: string; partyName: string; amount: number }>();
  const reportKeyMap = new Map<string, string>(); // normKey -> primaryDocNo

  for (const r of reportItems) {
    const doc = r.docNo;
    const ex = reportByDoc.get(doc);
    if (ex) {
      ex.amount = Math.round((ex.amount + r.amount) * 100) / 100;
      if (!ex.partyName && r.partyName) ex.partyName = r.partyName;
    } else {
      reportByDoc.set(doc, { originalDoc: doc, partyName: r.partyName, amount: Math.round(r.amount * 100) / 100 });
      for (const k of getNormalizedKeys(doc)) {
        if (!reportKeyMap.has(k)) reportKeyMap.set(k, doc);
      }
    }
  }

  // Aggregate lever items by normalized keys
  const leverByDoc = new Map<string, { originalDoc: string; partyName: string; amount: number }>();
  const leverKeyMap = new Map<string, string>(); // normKey -> primaryDocNo

  for (const l of leverItems) {
    const doc = l.docNo;
    const ex = leverByDoc.get(doc);
    if (ex) {
      ex.amount = Math.round((ex.amount + l.amount) * 100) / 100;
      if (!ex.partyName && l.partyName) ex.partyName = l.partyName;
    } else {
      leverByDoc.set(doc, { originalDoc: doc, partyName: l.partyName, amount: Math.round(l.amount * 100) / 100 });
      for (const k of getNormalizedKeys(doc)) {
        if (!leverKeyMap.has(k)) leverKeyMap.set(k, doc);
      }
    }
  }

  const matchedPairs = new Map<string, MatchItem>();
  const usedLeverDocs = new Set<string>();

  // Match report items against lever
  for (const [repDoc, repData] of reportByDoc.entries()) {
    let matchedLeverDoc: string | undefined;

    for (const k of getNormalizedKeys(repDoc)) {
      if (leverKeyMap.has(k)) {
        matchedLeverDoc = leverKeyMap.get(k);
        break;
      }
    }

    if (matchedLeverDoc && leverByDoc.has(matchedLeverDoc)) {
      const leverData = leverByDoc.get(matchedLeverDoc)!;
      usedLeverDocs.add(matchedLeverDoc);
      matchedPairs.set(repDoc, {
        docNo: repDoc,
        partyName: repData.partyName || leverData.partyName,
        reportAmt: repData.amount,
        leverAmt: leverData.amount,
      });
    } else {
      matchedPairs.set(repDoc, {
        docNo: repDoc,
        partyName: repData.partyName,
        reportAmt: repData.amount,
        leverAmt: null,
      });
    }
  }

  // Add remaining lever items that were not in report
  for (const [levDoc, levData] of leverByDoc.entries()) {
    if (!usedLeverDocs.has(levDoc)) {
      matchedPairs.set(levDoc, {
        docNo: levDoc,
        partyName: levData.partyName,
        reportAmt: null,
        leverAmt: levData.amount,
      });
    }
  }

  const rows: MatchRow[] = Array.from(matchedPairs.values()).map((item) => {
    const diff = (item.reportAmt ?? 0) - (item.leverAmt ?? 0);
    const roundDiff = Math.round(diff * 100) / 100;
    let status: MatchStatus;

    if (item.reportAmt === null) status = "only_lever";
    else if (item.leverAmt === null) status = "only_report";
    else if (Math.abs(roundDiff) < 1.0) status = "matched";
    else status = "diff";

    return {
      docNo: item.docNo,
      partyName: item.partyName,
      reportAmt: item.reportAmt,
      leverAmt: item.leverAmt,
      diff: Math.abs(roundDiff) < 0.01 ? 0 : roundDiff,
      status,
    };
  });

  const sortKey = (a: MatchRow, b: MatchRow) => {
    const numA = parseInt(a.docNo.replace(/[^0-9]/g, ""), 10);
    const numB = parseInt(b.docNo.replace(/[^0-9]/g, ""), 10);
    if (!isNaN(numA) && !isNaN(numB) && numA !== numB) return numA - numB;
    return a.docNo.localeCompare(b.docNo, undefined, { numeric: true, sensitivity: "base" });
  };

  rows.sort(sortKey);

  const reportTotal = Array.from(reportByDoc.values()).reduce((s, v) => s + v.amount, 0);
  const leverTotal = Array.from(leverByDoc.values()).reduce((s, v) => s + v.amount, 0);

  return {
    rows,
    reportTotal: Math.round(reportTotal * 100) / 100,
    leverTotal: Math.round(leverTotal * 100) / 100,
  };
}

export function performMatching(
  saleReportRows: SaleReportRow[],
  saleRetReportRows: SaleReportRow[],
  leverSaleRows: LeverSaleRow[],
  leverSaleRetRows: LeverSaleRetRow[]
): MatchingResult {
  const saleMatch = matchDataset(
    saleReportRows.map((r) => ({ docNo: r.docNo, partyName: r.partyName, amount: r.invoiceAmt })),
    leverSaleRows.map((r) => ({ docNo: r.billRefNo, partyName: r.partyName, amount: r.billValue }))
  );

  const retMatch = matchDataset(
    saleRetReportRows.map((r) => ({ docNo: r.docNo, partyName: r.partyName, amount: r.invoiceAmt })),
    leverSaleRetRows.map((r) => ({ docNo: r.salesRetNo, partyName: r.partyName, amount: r.total }))
  );

  return {
    saleRows: saleMatch.rows,
    saleRetRows: retMatch.rows,
    reportSaleTotal: saleMatch.reportTotal,
    reportSaleRetTotal: retMatch.reportTotal,
    leverSaleTotal: saleMatch.leverTotal,
    leverSaleRetTotal: retMatch.leverTotal,
  };
}

export function buildDiffWorkbook(result: MatchingResult): ArrayBuffer {
  const wb = XLSX.utils.book_new();

  const saleDiffHeaders = ["Doc No", "Party Name", "Sale Report Amt", "LeverEdge Amt", "Difference", "Status"];
  const saleRetDiffHeaders = ["Doc No", "Party Name", "Sale Report Amt", "LeverEdge Amt", "Difference", "Status"];

  const statusLabel: Record<MatchStatus, string> = {
    matched: "Matched",
    diff: "Amount Diff",
    only_report: "Only in Report",
    only_lever: "Only in LeverEdge",
  };

  const toSheetRows = (rows: MatchRow[]) =>
    rows.map((r) => [
      r.docNo,
      r.partyName,
      r.reportAmt ?? "",
      r.leverAmt ?? "",
      r.diff !== 0 ? r.diff : 0,
      statusLabel[r.status],
    ]);

  const mismatchSale = result.saleRows.filter((r) => r.status !== "matched");
  const mismatchRet = result.saleRetRows.filter((r) => r.status !== "matched");

  // Format headers and sheets
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([saleDiffHeaders, ...toSheetRows(mismatchSale)]), "Sale Mismatch");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([saleRetDiffHeaders, ...toSheetRows(mismatchRet)]), "Sale Return Mismatch");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([saleDiffHeaders, ...toSheetRows(result.saleRows)]), "Sale All");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([saleRetDiffHeaders, ...toSheetRows(result.saleRetRows)]), "Sale Return All");

  // Summary sheet
  const saleDiffTotal = Math.round((result.reportSaleTotal - result.leverSaleTotal) * 100) / 100;
  const retDiffTotal = Math.round((result.reportSaleRetTotal - result.leverSaleRetTotal) * 100) / 100;

  const summaryData = [
    ["Summary Report", ""],
    ["", ""],
    ["Sale Report Total", result.reportSaleTotal],
    ["LeverEdge Sale Total", result.leverSaleTotal],
    ["Sale Difference", saleDiffTotal],
    ["", ""],
    ["Sale Return Report Total", result.reportSaleRetTotal],
    ["LeverEdge Return Total", result.leverSaleRetTotal],
    ["Return Difference", retDiffTotal],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "Summary");

  return XLSX.write(wb, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}
