import { createFileRoute, Link } from "@tanstack/react-router";
import { useCallback, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  parseSaleReport,
  parseLeverSalesReg,
  parseLeverSalesRetReg,
  performMatching,
  buildDiffWorkbook,
  type MatchRow,
  type MatchStatus,
  type MatchingResult,
} from "@/lib/matching-parser";
import { downloadBuffer } from "@/lib/template-filler";

export const Route = createFileRoute("/matching")({
  head: () => ({
    meta: [{ title: "Data Matching — Excel Converter" }],
  }),
  component: MatchingPage,
});

const fmt = (n: number) =>
  new Intl.NumberFormat("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n);

const STATUS_LABEL: Record<MatchStatus, string> = {
  matched: "Matched",
  diff: "Diff",
  only_report: "Only Report",
  only_lever: "Only LeverEdge",
};

const STATUS_CLASS: Record<MatchStatus, string> = {
  matched: "bg-green-100 text-green-800",
  diff: "bg-red-100 text-red-800",
  only_report: "bg-amber-100 text-amber-800",
  only_lever: "bg-blue-100 text-blue-800",
};

type FilterType = "all" | MatchStatus;

interface FileSlot {
  name: string;
  buf: ArrayBuffer | null;
}

function FileUploadSlot({
  label,
  sub,
  slot,
  onChange,
}: {
  label: string;
  sub: string;
  slot: FileSlot;
  onChange: (file: File) => void;
}) {
  const ref = useRef<HTMLInputElement>(null);
  return (
    <div className="flex flex-col gap-1">
      <span className="text-xs font-semibold text-muted-foreground uppercase tracking-wide">{label}</span>
      <div
        className={`flex items-center gap-2 rounded-lg border px-3 py-2 text-sm cursor-pointer transition-colors ${
          slot.buf ? "border-green-400 bg-green-50" : "border-dashed border-border bg-muted/30 hover:bg-muted/60"
        }`}
        onClick={() => ref.current?.click()}
      >
        <span className="text-base">{slot.buf ? "✅" : "📂"}</span>
        <span className={`truncate flex-1 ${slot.buf ? "text-green-800 font-medium" : "text-muted-foreground"}`}>
          {slot.name || sub}
        </span>
        <button
          className="shrink-0 rounded bg-foreground/10 px-2 py-0.5 text-xs font-medium hover:bg-foreground/20"
          onClick={(e) => { e.stopPropagation(); ref.current?.click(); }}
        >
          Choose
        </button>
      </div>
      <input
        ref={ref}
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onChange(f);
          e.target.value = "";
        }}
      />
    </div>
  );
}

function StatCard({
  title,
  reportTotal,
  reportRetTotal,
  leverTotal,
  leverRetTotal,
  reportLabel,
  leverLabel,
}: {
  title: string;
  reportTotal: number | null;
  reportRetTotal: number | null;
  leverTotal: number | null;
  leverRetTotal: number | null;
  reportLabel: string;
  leverLabel: string;
}) {
  const diff = reportTotal !== null && leverTotal !== null ? reportTotal - leverTotal : null;
  const retDiff = reportRetTotal !== null && leverRetTotal !== null ? reportRetTotal - leverRetTotal : null;

  return (
    <div className="rounded-xl border bg-card p-5 shadow-sm flex flex-col gap-4">
      <h2 className="font-bold text-base text-foreground tracking-tight">{title}</h2>

      <div className="grid grid-cols-2 gap-3">
        <div className="rounded-lg bg-muted/50 p-3">
          <p className="text-xs text-muted-foreground mb-1">{reportLabel} — Sale</p>
          <p className="text-lg font-bold text-foreground">
            {reportTotal !== null ? `₹${fmt(reportTotal)}` : <span className="text-muted-foreground text-sm">—</span>}
          </p>
        </div>
        <div className="rounded-lg bg-muted/50 p-3">
          <p className="text-xs text-muted-foreground mb-1">{leverLabel} — Sale</p>
          <p className="text-lg font-bold text-foreground">
            {leverTotal !== null ? `₹${fmt(leverTotal)}` : <span className="text-muted-foreground text-sm">—</span>}
          </p>
        </div>
        <div className="rounded-lg bg-muted/50 p-3">
          <p className="text-xs text-muted-foreground mb-1">{reportLabel} — Return</p>
          <p className="text-lg font-bold text-foreground">
            {reportRetTotal !== null ? `₹${fmt(reportRetTotal)}` : <span className="text-muted-foreground text-sm">—</span>}
          </p>
        </div>
        <div className="rounded-lg bg-muted/50 p-3">
          <p className="text-xs text-muted-foreground mb-1">{leverLabel} — Return</p>
          <p className="text-lg font-bold text-foreground">
            {leverRetTotal !== null ? `₹${fmt(leverRetTotal)}` : <span className="text-muted-foreground text-sm">—</span>}
          </p>
        </div>
      </div>

      {(diff !== null || retDiff !== null) && (
        <div className="grid grid-cols-2 gap-3">
          {diff !== null && (
            <div className={`rounded-lg p-3 ${Math.abs(diff) < 1 ? "bg-green-50 border border-green-200" : "bg-red-50 border border-red-200"}`}>
              <p className="text-xs text-muted-foreground mb-1">Sale Difference</p>
              <p className={`text-base font-bold ${Math.abs(diff) < 1 ? "text-green-700" : "text-red-700"}`}>
                {diff >= 0 ? "+" : ""}{fmt(diff)}
              </p>
            </div>
          )}
          {retDiff !== null && (
            <div className={`rounded-lg p-3 ${Math.abs(retDiff) < 1 ? "bg-green-50 border border-green-200" : "bg-red-50 border border-red-200"}`}>
              <p className="text-xs text-muted-foreground mb-1">Return Difference</p>
              <p className={`text-base font-bold ${Math.abs(retDiff) < 1 ? "text-green-700" : "text-red-700"}`}>
                {retDiff >= 0 ? "+" : ""}{fmt(retDiff)}
              </p>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function MatchTable({ rows, filter, setFilter }: { rows: MatchRow[]; filter: FilterType; setFilter: (f: FilterType) => void }) {
  const [search, setSearch] = useState("");

  const counts = {
    all: rows.length,
    matched: rows.filter((r) => r.status === "matched").length,
    diff: rows.filter((r) => r.status === "diff").length,
    only_report: rows.filter((r) => r.status === "only_report").length,
    only_lever: rows.filter((r) => r.status === "only_lever").length,
  };

  const statusFiltered = filter === "all" ? rows : rows.filter((r) => r.status === filter);
  const q = search.trim().toLowerCase();
  const visible = q
    ? statusFiltered.filter(
        (r) =>
          r.docNo.toLowerCase().includes(q) ||
          r.partyName.toLowerCase().includes(q)
      )
    : statusFiltered;

  const tabs: { key: FilterType; label: string; color: string }[] = [
    { key: "all", label: `All (${counts.all})`, color: "bg-foreground/10 text-foreground" },
    { key: "diff", label: `Diff (${counts.diff})`, color: "bg-red-100 text-red-800" },
    { key: "only_report", label: `Only Report (${counts.only_report})`, color: "bg-amber-100 text-amber-800" },
    { key: "only_lever", label: `Only LeverEdge (${counts.only_lever})`, color: "bg-blue-100 text-blue-800" },
    { key: "matched", label: `Matched (${counts.matched})`, color: "bg-green-100 text-green-800" },
  ];

  return (
    <div className="flex flex-col gap-3">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex flex-wrap gap-2">
          {tabs.map((t) => (
            <button
              key={t.key}
              onClick={() => setFilter(t.key)}
              className={`rounded-full px-3 py-1 text-xs font-semibold transition-all ${
                filter === t.key ? `${t.color} ring-2 ring-offset-1 ring-current` : "bg-muted text-muted-foreground hover:bg-muted/80"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
        <input
          type="text"
          placeholder="🔍 Search Doc No / Party..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          className="rounded-lg border px-3 py-1 text-xs bg-background focus:outline-none focus:ring-1 focus:ring-foreground w-48 sm:w-60"
        />
      </div>

      <div className="overflow-auto rounded-lg border max-h-96">
        <table className="w-full text-xs">
          <thead className="sticky top-0 bg-muted border-b">
            <tr>
              <th className="px-3 py-2 text-left font-semibold">Doc No</th>
              <th className="px-3 py-2 text-left font-semibold">Party Name</th>
              <th className="px-3 py-2 text-right font-semibold">Report Amt</th>
              <th className="px-3 py-2 text-right font-semibold">LeverEdge Amt</th>
              <th className="px-3 py-2 text-right font-semibold">Diff</th>
              <th className="px-3 py-2 text-center font-semibold">Status</th>
            </tr>
          </thead>
          <tbody>
            {visible.length === 0 ? (
              <tr><td colSpan={6} className="px-3 py-6 text-center text-muted-foreground">No records</td></tr>
            ) : visible.map((r) => (
              <tr key={r.docNo} className="border-b last:border-0 hover:bg-muted/40">
                <td className="px-3 py-1.5 font-mono font-medium">{r.docNo}</td>
                <td className="px-3 py-1.5 max-w-[200px] truncate text-muted-foreground">{r.partyName}</td>
                <td className="px-3 py-1.5 text-right">{r.reportAmt !== null ? fmt(r.reportAmt) : "—"}</td>
                <td className="px-3 py-1.5 text-right">{r.leverAmt !== null ? fmt(r.leverAmt) : "—"}</td>
                <td className={`px-3 py-1.5 text-right font-semibold ${r.diff !== 0 ? "text-red-600" : "text-green-600"}`}>
                  {r.diff !== 0 ? (r.diff > 0 ? "+" : "") + fmt(r.diff) : "—"}
                </td>
                <td className="px-3 py-1.5 text-center">
                  <span className={`rounded-full px-2 py-0.5 text-[10px] font-semibold ${STATUS_CLASS[r.status]}`}>
                    {STATUS_LABEL[r.status]}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <p className="text-xs text-muted-foreground text-right">Showing {visible.length} of {rows.length} rows</p>
    </div>
  );
}

function MatchingPage() {
  const [saleReport, setSaleReport] = useState<FileSlot>({ name: "", buf: null });
  const [leverSaleReg, setLeverSaleReg] = useState<FileSlot>({ name: "", buf: null });
  const [leverSaleRetReg, setLeverSaleRetReg] = useState<FileSlot>({ name: "", buf: null });

  const [result, setResult] = useState<MatchingResult | null>(null);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const [saleFilter, setSaleFilter] = useState<FilterType>("all");
  const [retFilter, setRetFilter] = useState<FilterType>("all");
  const [activeTab, setActiveTab] = useState<"sale" | "sale_return">("sale");

  const currentSaleReport = useRef<ArrayBuffer | null>(null);
  const currentLeverSaleReg = useRef<ArrayBuffer | null>(null);
  const currentLeverSaleRetReg = useRef<ArrayBuffer | null>(null);

  const tryMatch = useCallback(async (sr: ArrayBuffer | null, lsr: ArrayBuffer | null, lsrr: ArrayBuffer | null) => {
    if (!sr || !lsr) return;
    setProcessing(true);
    setError(null);
    try {
      const { saleRows: reportSale, saleRetRows: reportSaleRet } = parseSaleReport(sr);
      const { saleRows: leverSale, saleRetRows: leverRegRet } = parseLeverSalesReg(lsr);
      const leverRetFromExtra = lsrr ? parseLeverSalesRetReg(lsrr) : [];
      
      // Merge LeverEDGE Sales Register (- Bill Value) rows with any extra Sales Return Register rows
      const leverSaleRet = [...leverRegRet, ...leverRetFromExtra];
      const res = performMatching(reportSale, reportSaleRet, leverSale, leverSaleRet);
      setResult(res);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setProcessing(false);
    }
  }, []);

  const handleSaleReport = useCallback(async (file: File) => {
    const buf = await file.arrayBuffer();
    setSaleReport({ name: file.name, buf });
    currentSaleReport.current = buf;
    tryMatch(buf, currentLeverSaleReg.current, currentLeverSaleRetReg.current);
  }, [tryMatch]);

  const handleLeverSaleReg = useCallback(async (file: File) => {
    const buf = await file.arrayBuffer();
    setLeverSaleReg({ name: file.name, buf });
    currentLeverSaleReg.current = buf;
    tryMatch(currentSaleReport.current, buf, currentLeverSaleRetReg.current);
  }, [tryMatch]);

  const handleLeverSaleRetReg = useCallback(async (file: File) => {
    const buf = await file.arrayBuffer();
    setLeverSaleRetReg({ name: file.name, buf });
    currentLeverSaleRetReg.current = buf;
    tryMatch(currentSaleReport.current, currentLeverSaleReg.current, buf);
  }, [tryMatch]);

  const handleDownload = useCallback(() => {
    if (!result) return;
    const buf = buildDiffWorkbook(result);
    const today = new Date().toISOString().slice(0, 10);
    downloadBuffer(buf, `Sale_Matching_Diff_${today}.xlsx`);
  }, [result]);

  const hasRequiredFiles = saleReport.buf && leverSaleReg.buf;

  const saleMatchedCount = result?.saleRows.filter((r) => r.status === "matched").length ?? 0;
  const saleMismatchCount = result ? result.saleRows.length - saleMatchedCount : 0;
  const retMatchedCount = result?.saleRetRows.filter((r) => r.status === "matched").length ?? 0;
  const retMismatchCount = result ? result.saleRetRows.length - retMatchedCount : 0;

  return (
    <div className="min-h-screen bg-background">
      <div className="mx-auto max-w-6xl px-4 py-3 flex flex-col gap-3">

        <h1 className="text-lg font-semibold tracking-tight">Data Matching</h1>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <StatCard
            title="📊 Sale & Return Summary"
            reportTotal={result?.reportSaleTotal ?? null}
            reportRetTotal={result?.reportSaleRetTotal ?? null}
            leverTotal={result?.leverSaleTotal ?? null}
            leverRetTotal={result?.leverSaleRetTotal ?? null}
            reportLabel="Hisab Kitab"
            leverLabel="LeverEdge"
          />
          <div className="rounded-xl border bg-card p-5 shadow-sm flex flex-col gap-3">
            <h2 className="font-bold text-base text-foreground tracking-tight">ℹ️ Matching Rules</h2>
            <ul className="text-xs text-muted-foreground space-y-1.5 list-disc pl-4">
              <li><strong className="text-foreground">Sale Match:</strong> Hisab Kitab Bill No wise <code className="text-foreground font-semibold">Invoice Amount</code> vs LeverEdge Sales Register <code className="text-emerald-700 font-semibold">(+) Bill Value</code> column.</li>
              <li><strong className="text-foreground">Sale Return Match:</strong> Hisab Kitab <code className="text-foreground font-semibold">(-) Invoice Amount</code> vs LeverEdge <code className="text-red-700 font-semibold">(-) Return Amount</code> SRT No wise.</li>
              <li>LeverEdge Sales Register ki negative rows automatically Sale Return me match hoti hain.</li>
            </ul>
          </div>
        </div>

        <div className="rounded-xl border bg-card p-5 shadow-sm">
          <h2 className="font-semibold text-sm mb-3 text-foreground">Upload Files for Matching</h2>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <FileUploadSlot
              label="1. Sale Report (Required)"
              sub="Hisab Kitab sale report (.xlsx)"
              slot={saleReport}
              onChange={handleSaleReport}
            />
            <FileUploadSlot
              label="2. LeverEdge Sales Register (Required)"
              sub="Sales Register with Bill Value (.xlsx)"
              slot={leverSaleReg}
              onChange={handleLeverSaleReg}
            />
            <FileUploadSlot
              label="3. Sales Return Register (Optional)"
              sub="salesReturnDetail (.xlsx)"
              slot={leverSaleRetReg}
              onChange={handleLeverSaleRetReg}
            />
          </div>
          {!hasRequiredFiles && (
            <p className="mt-3 text-xs text-amber-700 bg-amber-50 rounded-lg px-3 py-2">
              ⏳ Pehle 2 files (Hisab Kitab Sale Report & LeverEdge Sales Register) upload karo — matching automatically run hogi.
            </p>
          )}
        </div>

        {processing && (
          <div className="rounded-xl border bg-card p-6 text-center text-sm text-muted-foreground">
            ⏳ Matching chal rahi hai…
          </div>
        )}

        {error && (
          <div className="rounded-xl border border-red-200 bg-red-50 p-4 text-sm text-red-800">
            ❌ Error: {error}
          </div>
        )}

        {result && !processing && (
          <div className="rounded-xl border bg-card p-5 shadow-sm flex flex-col gap-4">
            <div className="flex items-center justify-between flex-wrap gap-3">
              <div className="flex gap-3 flex-wrap">
                <div className="text-xs bg-green-50 border border-green-200 rounded-lg px-3 py-1.5">
                  <span className="font-semibold text-green-800">Sale:</span>{" "}
                  <span className="text-green-700">{saleMatchedCount} matched</span>
                  {saleMismatchCount > 0 && <span className="text-red-600 ml-2">{saleMismatchCount} mismatch</span>}
                </div>
                <div className="text-xs bg-green-50 border border-green-200 rounded-lg px-3 py-1.5">
                  <span className="font-semibold text-green-800">Return:</span>{" "}
                  <span className="text-green-700">{retMatchedCount} matched</span>
                  {retMismatchCount > 0 && <span className="text-red-600 ml-2">{retMismatchCount} mismatch</span>}
                </div>
              </div>
              <button
                onClick={handleDownload}
                className="inline-flex items-center gap-1.5 rounded-lg bg-foreground px-4 py-2 text-sm font-semibold text-background hover:bg-foreground/90 transition-colors"
              >
                ⬇ Download Diff Excel
              </button>
            </div>

            <div className="flex gap-2 border-b">
              <button
                onClick={() => setActiveTab("sale")}
                className={`px-4 py-2 text-sm font-semibold border-b-2 transition-colors ${
                  activeTab === "sale" ? "border-foreground text-foreground" : "border-transparent text-muted-foreground hover:text-foreground"
                }`}
              >
                Sale ({result.saleRows.length})
              </button>
              <button
                onClick={() => setActiveTab("sale_return")}
                className={`px-4 py-2 text-sm font-semibold border-b-2 transition-colors ${
                  activeTab === "sale_return" ? "border-foreground text-foreground" : "border-transparent text-muted-foreground hover:text-foreground"
                }`}
              >
                Sale Return ({result.saleRetRows.length})
              </button>
            </div>

            {activeTab === "sale" ? (
              <MatchTable rows={result.saleRows} filter={saleFilter} setFilter={setSaleFilter} />
            ) : (
              <MatchTable rows={result.saleRetRows} filter={retFilter} setFilter={setRetFilter} />
            )}
          </div>
        )}
      </div>
    </div>
  );
}
