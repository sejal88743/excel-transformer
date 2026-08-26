// Saves converted rows to Cloud if configured. Duplicates (bill_no + item + date) are skipped.
import { supabase, isSupabaseConfigured } from "@/integrations/supabase/client";
import { normalizeItemName, normalizeItemKey, cleanItemKey, cleanHSN } from "@/lib/item-master";

type Row = Record<string, unknown>;

const s = (v: unknown) => (v === null || v === undefined ? "" : String(v)).trim();
const n = (v: unknown): number | null => {
  if (v === null || v === undefined || v === "") return null;
  const x = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return isNaN(x) ? null : x;
};

async function upsertUniqueNames(table: "parties" | "salespersons" | "items", rows: Record<string, unknown>[], mergeFields = false) {
  if (!isSupabaseConfigured || !rows.length) return;
  // chunk to avoid oversized payloads
  const CHUNK = 500;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const slice = rows.slice(i, i + CHUNK);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { error } = await (supabase.from(table) as any).upsert(slice, { onConflict: "name", ignoreDuplicates: !mergeFields });
      if (error) console.warn(`[cloud-saver] ${table} upsert warning:`, error.message);
    } catch {
      // Ignore network errors when Supabase is not reachable
    }
  }
}

async function insertLines(table: "sale_lines" | "sale_return_lines" | "purchase_lines", rows: Record<string, unknown>[], conflictCols: string) {
  if (!isSupabaseConfigured || !rows.length) return { inserted: 0, skipped: 0 };
  const CHUNK = 300;
  let inserted = 0;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const slice = rows.slice(i, i + CHUNK);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase.from(table) as any)
        .upsert(slice, { onConflict: conflictCols, ignoreDuplicates: true })
        .select("id");
      if (error) { console.warn(`[cloud-saver] ${table} insert warning:`, error.message); continue; }
      inserted += data?.length ?? 0;
    } catch {
      // Ignore network errors when Supabase is not reachable
    }
  }
  return { inserted, skipped: rows.length - inserted };
}

export async function saveSaleToCloud(rows: Row[]) {
  if (!isSupabaseConfigured) return { inserted: 0, skipped: 0 };
  const parties = new Map<string, { name: string; gstin: string | null }>();
  const sps = new Set<string>();
  const items = new Map<string, Record<string, unknown>>();
  const lines: Record<string, unknown>[] = [];

  for (const r of rows) {
    const bill = s(r["Invoice Number*"]);
    const date = s(r["Date*"]);
    const rawItem = s(r["Item Name Or Alias Name Or SKU*"]);
    const item = normalizeItemName(rawItem);
    if (!bill || !date || !item) continue;
    const itemKey = normalizeItemKey(item);
    const party = s(r["Customer Name Or Alias Name*"]);
    const sp = s(r["Broker Name"]);
    if (party) parties.set(party, { name: party, gstin: s(r["GSTIN"]) || null });
    if (sp) sps.add(sp);
    const hsn = cleanHSN(r["Hsn Code"] ?? r["HSN Code"] ?? r["CF - Item Custom Field Name 1"]);
    const existing = items.get(itemKey);
    if (!existing) {
      items.set(itemKey, { name: item, hsn: hsn || null, gst_rate: n(r["GST %"]), mrp: n(r["MRP"]) });
    } else if (!existing.hsn && hsn) {
      existing.hsn = hsn;
    }
    lines.push({
      bill_no: bill,
      bill_date: date,
      item_name: item,
      party_name: party || null,
      salesperson: sp || null,
      qty: n(r["Quantity*"]),
      rate: n(r["Rate per Unit (Without GST)"]),
      mrp: n(r["MRP"]),
      gst_pct: n(r["GST %"]),
      amount: n(r["Taxable Value For TCS"]),
      data: r,
    });
  }
  await Promise.all([
    upsertUniqueNames("parties", Array.from(parties.values())),
    upsertUniqueNames("salespersons", Array.from(sps).map((name) => ({ name }))),
    upsertUniqueNames("items", Array.from(items.values()), true),
  ]);
  return insertLines("sale_lines", lines, "bill_no,item_name,bill_date");
}

export async function saveSaleReturnToCloud(rows: Row[]) {
  if (!isSupabaseConfigured) return { inserted: 0, skipped: 0 };
  const parties = new Map<string, { name: string; gstin: string | null }>();
  const sps = new Set<string>();
  const items = new Map<string, Record<string, unknown>>();
  const lines: Record<string, unknown>[] = [];

  for (const r of rows) {
    const bill = s(r["Invoice Number*"] ?? r["Voucher Number*"] ?? r["Credit Note Number*"]);
    const date = s(r["Date*"] ?? r["Voucher Date*"] ?? r["Credit Note Date*"]);
    const rawItem = s(r["Item Name Or Alias Name Or SKU*"]);
    const item = normalizeItemName(rawItem);
    if (!bill || !date || !item) continue;
    const itemKey = normalizeItemKey(item);
    const party = s(r["Customer Name Or Alias Name*"] ?? r["Supplier Name Or Alias Name*"]);
    const sp = s(r["Broker Name"]);
    if (party) parties.set(party, { name: party, gstin: s(r["GSTIN"]) || null });
    if (sp) sps.add(sp);
    const hsn = cleanHSN(r["Hsn Code"] ?? r["HSN Code"] ?? r["CF - Item Custom Field Name 1"]);
    const existing = items.get(itemKey);
    if (!existing) {
      items.set(itemKey, { name: item, hsn: hsn || null, gst_rate: n(r["GST %"]), mrp: n(r["MRP"]) });
    } else if (!existing.hsn && hsn) {
      existing.hsn = hsn;
    }
    lines.push({
      bill_no: bill,
      bill_date: date,
      item_name: item,
      party_name: party || null,
      salesperson: sp || null,
      qty: n(r["Quantity*"]),
      rate: n(r["Rate per Unit (Without GST)*"] ?? r["Rate per Unit (Without GST)"]),
      mrp: n(r["MRP"]),
      gst_pct: n(r["GST %"]),
      amount: n(r["Taxable Value For TCS"]),
      data: r,
    });
  }
  await Promise.all([
    upsertUniqueNames("parties", Array.from(parties.values())),
    upsertUniqueNames("salespersons", Array.from(sps).map((name) => ({ name }))),
    upsertUniqueNames("items", Array.from(items.values()), true),
  ]);
  return insertLines("sale_return_lines", lines, "bill_no,item_name,bill_date");
}

export async function savePurchaseToCloud(rows: Row[]) {
  if (!isSupabaseConfigured) return { inserted: 0, skipped: 0 };
  const items = new Map<string, Record<string, unknown>>();
  const parties = new Map<string, { name: string; gstin: string | null }>();
  const lines: Record<string, unknown>[] = [];

  for (const r of rows) {
    const inv = s(r["Invoice Number*"]);
    const date = s(r["Invoice Date*"]);
    const rawItem = s(r["Item Name Or Alias Name Or SKU*"]);
    const item = normalizeItemName(rawItem);
    if (!inv || !date || !item) continue;
    const itemKey = normalizeItemKey(item);
    const supplier = s(r["Supplier Name Or Alias Name*"]);
    if (supplier) parties.set(supplier, { name: supplier, gstin: s(r["GSTIN"]) || null });
    if (!items.has(itemKey)) {
      items.set(itemKey, {
        name: item,
        hsn: cleanHSN(r["Hsn Code"] ?? r["HSN Code"] ?? r["CF - Item Custom Field Name 1"] ?? "") || null,
        gst_rate: n(r["GST %"]),
        mrp: n(r["MRP"]),
        purchase_price: n(r["Rate per Unit (Without GST)"]),
      });
    }
    lines.push({
      voucher_no: s(r["Voucher Number*"]) || inv,
      invoice_no: inv,
      invoice_date: date,
      supplier_name: supplier || null,
      item_name: item,
      qty: n(r["Quantity*"]),
      rate: n(r["Rate per Unit (Without GST)"]),
      mrp: n(r["MRP"]),
      gst_pct: n(r["GST %"]),
      data: r,
    });
  }
  await Promise.all([
    upsertUniqueNames("parties", Array.from(parties.values())),
    upsertUniqueNames("items", Array.from(items.values()), true),
  ]);
  return insertLines("purchase_lines", lines, "invoice_no,item_name,invoice_date");
}

/** HSN lookup map built from the cloud `items` table (item key → HSN). */
export async function fetchCloudHsnMap(): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  if (!isSupabaseConfigured) return map;
  try {
    const { data, error } = await supabase.from("items").select("name,hsn").not("hsn", "is", null).limit(20000);
    if (error || !data) return map;
    for (const it of data as { name: string; hsn: string | null }[]) {
      const hsn = cleanHSN(it.hsn);
      if (!hsn || !it.name) continue;
      const name = normalizeItemName(it.name);
      map.set(normalizeItemKey(name), hsn);
      map.set(cleanItemKey(name), hsn);
    }
  } catch {
    // offline — fall back to local item master only
  }
  return map;
}
