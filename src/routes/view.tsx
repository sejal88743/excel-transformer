import { createFileRoute, Link } from "@tanstack/react-router";
import { useMemo, useState, useSyncExternalStore } from "react";
import { datasetStore, type Row } from "@/lib/dataset-store";

export const Route = createFileRoute("/view")({
  head: () => ({
    meta: [
      { title: "Data Viewer — Sale, Return & Purchase" },
      { name: "description", content: "View uploaded Sale, Sale Return and Purchase data." },
    ],
  }),
  component: ViewPage,
});

type Tab = "sale" | "return" | "purchase" | "register";

const SALE_COLS: { key: string; label: string; num?: boolean }[] = [
  { key: "Invoice Number*", label: "Bill No" },
  { key: "Date*", label: "Date" },
  { key: "Customer Name Or Alias Name*", label: "Customer" },
  { key: "Item Name Or Alias Name Or SKU*", label: "Item" },
  { key: "Hsn Code", label: "HSN" },
  { key: "Quantity*", label: "Qty", num: true },
  { key: "Rate per Unit (Without GST)", label: "Rate", num: true },
  { key: "Discount 1", label: "Discount", num: true },
  { key: "GST %", label: "GST%", num: true },
  { key: "Broker Name", label: "Broker" },
];

const RETURN_COLS: { key: string; label: string; num?: boolean }[] = [
  { key: "Credit Note Number*", label: "CN No" },
  { key: "Credit Note Date*", label: "Date" },
  { key: "Customer Name Or Alias Name*", label: "Customer" },
  { key: "Original Invoice Number", label: "Bill No" },
  { key: "Item Name Or Alias Name Or SKU*", label: "Item" },
  { key: "Hsn Code", label: "HSN" },
  { key: "Quantity*", label: "Qty", num: true },
  { key: "Rate per Unit (Without GST)*", label: "Rate", num: true },
  { key: "Discount 1", label: "Discount", num: true },
  { key: "GST %", label: "GST%", num: true },
  { key: "Classification Nature Type", label: "Classification" },
];

const PURCHASE_COLS: { key: string; label: string; num?: boolean }[] = [
  { key: "Voucher Number*", label: "Voucher" },
  { key: "Invoice Number*", label: "Invoice No" },
  { key: "Invoice Date*", label: "Date" },
  { key: "Supplier Name Or Alias Name*", label: "Supplier" },
  { key: "Item Name Or Alias Name Or SKU*", label: "Item" },
  { key: "Quantity*", label: "Qty", num: true },
  { key: "Rate per Unit (Without GST)", label: "Rate", num: true },
  { key: "Discount 1", label: "Discount", num: true },
  { key: "GST %", label: "GST%", num: true },
  { key: "Payment 1 Amount", label: "Net Amt", num: true },
  { key: "Classification Nature Type", label: "Classification" },
];

function useStore() {
  const version = useSyncExternalStore(
    (cb) => datasetStore.subscribe(cb),
    () => datasetStore.getVersion(),
    () => 0,
  );
  void version;
  return { sale: datasetStore.getSale(), ret: datasetStore.getReturn(), pur: datasetStore.getPurchase(), reg: datasetStore.getRegister() };
}

function n(v: unknown): number {
  if (typeof v === "number") return v;
  const x = parseFloat(String(v ?? "").replace(/,/g, ""));
  return Number.isFinite(x) ? x : 0;
}

function ViewPage() {
  const { sale, ret, pur, reg } = useStore();
  const [tab, setTab] = useState<Tab>("sale");
  const [groupBy, setGroupBy] = useState<string>("__none__");
  const [search, setSearch] = useState("");
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");

  // Register columns are derived dynamically from the first row of the uploaded file.
  const registerCols = useMemo(() => {
    if (!reg || reg.rows.length === 0) return [] as { key: string; label: string; num?: boolean }[];
    const NUMERIC_HINTS = ["amount", "value", "qty", "disc", "tax", "rate", "total", "redemption", "scheme"];
    return Object.keys(reg.rows[0]).slice(0, 20).map((k) => ({
      key: k,
      label: k,
      num: NUMERIC_HINTS.some((h) => k.toLowerCase().includes(h)),
    }));
  }, [reg]);

  const cols = tab === "sale" ? SALE_COLS : tab === "return" ? RETURN_COLS : tab === "purchase" ? PURCHASE_COLS : registerCols;
  const dataset = tab === "sale" ? sale : tab === "return" ? ret : tab === "purchase" ? pur : reg;
  const rows: Row[] = dataset?.rows ?? [];

  const parseDate = (v: unknown): Date | null => {
    const s = String(v ?? "").trim();
    if (!s) return null;
    const m = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (m) { const yy = m[3].length === 2 ? 2000 + +m[3] : +m[3]; return new Date(yy, +m[2] - 1, +m[1]); }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  };

  const dateKey = tab === "sale" ? "Date*" : tab === "return" ? "Credit Note Date*" : tab === "purchase" ? "Invoice Date*" : "Bill Date";
  const qtyKey  = tab === "register" ? "Qty" : "Quantity*";
  const getRate = (r: Row) => n(r["Rate per Unit (Without GST)*"] ?? r["Rate per Unit (Without GST)"]);
  const amtKey  = tab === "purchase" ? "Payment 1 Amount" : tab === "register" ? "Bill Value" : null;

  const filtered = useMemo(() => {
    const q = search.toLowerCase().trim();
    const from = fromDate ? new Date(fromDate) : null;
    const to   = toDate   ? new Date(toDate)   : null;
    return rows.filter((r) => {
      if (q) { const hay = cols.map((c) => String(r[c.key] ?? "")).join(" ").toLowerCase(); if (!hay.includes(q)) return false; }
      if (from || to) {
        const d = parseDate(r[dateKey]);
        if (!d) return false;
        if (from && d < from) return false;
        if (to   && d > to)   return false;
      }
      return true;
    });
  }, [rows, search, fromDate, toDate, cols, dateKey]);

  const totals = useMemo(() => {
    let qty = 0, amt = 0;
    for (const r of filtered) {
      const q = n(r[qtyKey]);
      qty += q;
      amt += amtKey ? n(r[amtKey]) : q * getRate(r);
    }
    return { qty, amt, rows: filtered.length };
  }, [filtered, amtKey]);

  const grouped = useMemo(() => {
    if (groupBy === "__none__") return null;
    const map = new Map<string, { rows: Row[]; qty: number; amt: number }>();
    for (const r of filtered) {
      const key = String(r[groupBy] ?? "—").trim() || "—";
      const g = map.get(key) ?? { rows: [], qty: 0, amt: 0 };
      const q = n(r[qtyKey]);
      g.rows.push(r);
      g.qty += q;
      g.amt += amtKey ? n(r[amtKey]) : q * getRate(r);
      map.set(key, g);
    }
    return [...map.entries()].sort((a, b) => b[1].amt - a[1].amt);
  }, [filtered, groupBy, amtKey]);

  const fmt = (x: number) => x.toLocaleString("en-IN", { maximumFractionDigits: 2 });

  const TAB_CONFIG: { key: Tab; label: string; count: number }[] = [
    { key: "sale",     label: "Sale Data",       count: sale?.rows.length ?? 0 },
    { key: "return",   label: "Sale Return",     count: ret?.rows.length  ?? 0 },
    { key: "purchase", label: "Purchase",        count: pur?.rows.length  ?? 0 },
    { key: "register", label: "Sales Register",  count: reg?.rows.length  ?? 0 },
  ];

  return (
    <div className="min-h-screen bg-background text-foreground">
      <div className="mx-auto max-w-7xl px-4 py-3">
        <div className="mb-3 flex flex-wrap items-center gap-3">
          <h1 className="text-lg font-semibold tracking-tight">Data View</h1>
          <span className="text-xs text-muted-foreground">RAM only — refresh to clear.</span>
        </div>

        <div className="mb-3 inline-flex rounded-md border border-border bg-muted/30 p-1">
          {TAB_CONFIG.map((t) => (
            <button key={t.key} onClick={() => { setTab(t.key); setGroupBy("__none__"); setSearch(""); }}
              className={`rounded px-3 py-1.5 text-sm font-medium transition ${tab === t.key ? "bg-background shadow-sm" : "text-muted-foreground hover:text-foreground"}`}>
              {t.label}
              <span className="ml-2 text-xs text-muted-foreground">({t.count})</span>
            </button>
          ))}
        </div>

        {!dataset ? (
          <div className="rounded-xl border border-dashed border-border p-10 text-center text-sm text-muted-foreground">
            Koi data nahi mila. <Link to="/" className="underline">Converter</Link> par jaakar file upload karein.
          </div>
        ) : (
          <>
            <div className="mb-2 flex flex-wrap items-center gap-3 text-xs text-muted-foreground">
              <span>📄 <span className="font-mono">{dataset.fileName}</span></span>
              <span>•</span>
              <span>{new Date(dataset.uploadedAt).toLocaleString()}</span>
            </div>

            <div className="mb-2 grid gap-2 sm:grid-cols-2 lg:grid-cols-4">
              <input placeholder="Search…" value={search} onChange={(e) => setSearch(e.target.value)}
                className="rounded-md border border-input bg-background px-3 py-2 text-sm" />
              <input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} className="rounded-md border border-input bg-background px-3 py-2 text-sm" />
              <input type="date" value={toDate}   onChange={(e) => setToDate(e.target.value)}   className="rounded-md border border-input bg-background px-3 py-2 text-sm" />
              <select value={groupBy} onChange={(e) => setGroupBy(e.target.value)} className="rounded-md border border-input bg-background px-3 py-2 text-sm">
                <option value="__none__">Group By: None</option>
                {cols.filter((c) => !c.num).map((c) => (
                  <option key={c.key} value={c.key}>Group By: {c.label}</option>
                ))}
                <option value={dateKey}>Group By: Date</option>
              </select>
            </div>

            <div className="mb-2 grid grid-cols-3 gap-2">
              <StatBox label="Rows"      value={totals.rows.toLocaleString()} />
              <StatBox label="Total Qty" value={fmt(totals.qty)} />
              <StatBox label={tab === "purchase" ? "Total Net Amount" : "Total Amount (Qty × Rate)"} value={`₹ ${fmt(totals.amt)}`} accent />
            </div>

            <div className="overflow-auto rounded-lg border border-border">
              <table className="w-full text-xs">
                <thead className="bg-muted/50">
                  <tr>
                    {cols.map((c) => (
                      <th key={c.key} className={`border-b border-border px-2 py-2 text-left font-medium ${c.num ? "text-right" : ""}`}>{c.label}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {grouped ? (
                    grouped.map(([key, g]) => (
                      <GroupBlock key={key} title={key} cols={cols} rows={g.rows} qty={g.qty} amt={g.amt} fmt={fmt} />
                    ))
                  ) : (
                    filtered.slice(0, 2000).map((r, i) => (
                      <tr key={i} className="border-b border-border/50 hover:bg-muted/30">
                        {cols.map((c) => (
                          <td key={c.key} className={`px-2 py-1 ${c.num ? "text-right tabular-nums" : ""}`}>
                            {c.num ? fmt(n(r[c.key])) : String(r[c.key] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))
                  )}
                </tbody>
                <tfoot className="bg-muted/40 font-medium">
                  <tr>
                    <td colSpan={cols.length} className="px-2 py-2 text-right">
                      Total Rows: {totals.rows.toLocaleString()} · Qty: {fmt(totals.qty)} · Amount: ₹ {fmt(totals.amt)}
                    </td>
                  </tr>
                </tfoot>
              </table>
              {!grouped && filtered.length > 2000 && (
                <div className="border-t border-border bg-muted/30 px-3 py-2 text-xs text-muted-foreground">
                  Showing first 2000 of {filtered.length.toLocaleString()} rows. Use filters/group-by to narrow down.
                </div>
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function GroupBlock({ title, cols, rows, qty, amt, fmt }: {
  title: string; cols: { key: string; label: string; num?: boolean }[];
  rows: Row[]; qty: number; amt: number; fmt: (x: number) => string;
}) {
  const [open, setOpen] = useState(false);
  return (
    <>
      <tr className="cursor-pointer bg-primary/5 hover:bg-primary/10" onClick={() => setOpen((v) => !v)}>
        <td colSpan={cols.length} className="px-2 py-1.5 font-medium">
          <span className="mr-1 inline-block w-3">{open ? "▼" : "▶"}</span>
          {title} <span className="ml-2 text-xs text-muted-foreground">({rows.length} rows · Qty {fmt(qty)} · ₹ {fmt(amt)})</span>
        </td>
      </tr>
      {open && rows.slice(0, 500).map((r, i) => (
        <tr key={i} className="border-b border-border/40">
          {cols.map((c) => (
            <td key={c.key} className={`px-2 py-1 ${c.num ? "text-right tabular-nums" : ""}`}>
              {c.num ? fmt(n(r[c.key])) : String(r[c.key] ?? "")}
            </td>
          ))}
        </tr>
      ))}
    </>
  );
}

function StatBox({ label, value, accent }: { label: string; value: string; accent?: boolean }) {
  return (
    <div className={`rounded-lg border p-3 ${accent ? "border-primary/30 bg-primary/5" : "border-border bg-muted/30"}`}>
      <div className="text-xs text-muted-foreground">{label}</div>
      <div className="mt-1 text-lg font-semibold tabular-nums">{value}</div>
    </div>
  );
}
