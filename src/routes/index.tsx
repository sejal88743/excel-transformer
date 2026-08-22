import { createFileRoute, Link } from "@tanstack/react-router";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  convertSaleData,
  buildOutputWorkbooks,
  parseSalesRegister,
  parseSalesReturnRegister,
  validateSaleDataFile,
  validateSalesRegisterFile,
  todayFolderName,
  type ConvertStats,
  type RegisterDiscount,
} from "@/lib/sale-converter";
import {
  convertSaleReturn,
  buildSRWorkbooks,
  validateSaleReturnFile,
  type SRConvertStats,
} from "@/lib/sale-return-converter";
import { convertPurchase, buildPurchaseWorkbook, type PurConvertStats } from "@/lib/purchase-converter";
import { downloadBuffer, downloadMultipleBuffers, downloadZipBundle } from "@/lib/template-filler";
import { datasetStore } from "@/lib/dataset-store";
import {
  mergeFromPurchaseRows,
  importPurchaseFileToMaster,
  buildItemMasterWorkbook,
  getMasterCount,
  clearMaster,
  parseExistingItemMaster,
  getNewItemsSummary,
} from "@/lib/item-master";
import { saveSaleToCloud, saveSaleReturnToCloud, savePurchaseToCloud } from "@/lib/cloud-saver";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Excel Converter — Privacy-First" },
      { name: "description", content: "Convert Sale, Sale Return & Purchase Excel files to import-ready format. 100% in your browser." },
    ],
  }),
  component: Index,
});

type Stage = "idle" | "processing" | "ready" | "downloaded" | "error";

interface SaleState {
  stage: Stage;
  stats: ConvertStats | null;
  error: string | null;
  billWiseFileName: string;
  regFileName: string;
  srFileName: string;
}
interface SRState    { stage: Stage; stats: SRConvertStats | null;  error: string | null; srFileName: string; regFileName: string; }
interface PurState   { stage: Stage; stats: PurConvertStats | null; error: string | null; fileName: string; }

type DirHandle = FileSystemDirectoryHandle | null;

function Index() {
  const [folderDate, setFolderDate] = useState<string>(() => {
    try { return localStorage.getItem("xlsxConverterDate") || todayFolderName(); } catch { return todayFolderName(); }
  });
  const [dirHandle, setDirHandle] = useState<DirHandle>(null);
  const [autoStatus, setAutoStatus] = useState<string>("");

  const [sale, setSale] = useState<SaleState>({ stage: "idle", stats: null, error: null, billWiseFileName: "", regFileName: "", srFileName: "" });
  const [sr,   setSR]   = useState<SRState>  ({ stage: "idle", stats: null, error: null, srFileName: "", regFileName: "" });
  const [pur,  setPur]  = useState<PurState>  ({ stage: "idle", stats: null, error: null, fileName: "" });

  const saleBlob       = useRef<ArrayBuffer[] | null>(null);
  const saleBuf        = useRef<ArrayBuffer | null>(null);   // Bill Wise Sales raw buffer
  const saleRegBuf     = useRef<ArrayBuffer | null>(null);   // Sales Register raw buffer
  const saleSRBuf      = useRef<ArrayBuffer | null>(null);   // Sale Return optional raw buffer
  const saleRegDiscMap = useRef<Map<string, RegisterDiscount> | null>(null);

  const srBlob      = useRef<ArrayBuffer[] | null>(null);
  const srBuf       = useRef<ArrayBuffer | null>(null);   // SR file raw buffer
  const srRegBuf    = useRef<ArrayBuffer | null>(null);   // SR's own Sales Register raw buffer
  const srRegDiscMap = useRef<Map<string, RegisterDiscount> | null>(null);
  const purBlob     = useRef<ArrayBuffer | null>(null);

  const billWiseInput = useRef<HTMLInputElement>(null);
  const regInput      = useRef<HTMLInputElement>(null);
  const saleSRInput   = useRef<HTMLInputElement>(null);
  const srInput       = useRef<HTMLInputElement>(null);
  const srRegInput    = useRef<HTMLInputElement>(null);
  const purInput      = useRef<HTMLInputElement>(null);

  // ─── Existing Item Master upload & diffing ──────────────────────────────
  const [existingMaster, setExistingMaster] = useState<{
    fileName: string;
    existingNames: Set<string> | null;
    totalInFile: number;
    status: string | null;
  }>({
    fileName: "",
    existingNames: null,
    totalInFile: 0,
    status: null,
  });
  const existingMasterInput = useRef<HTMLInputElement>(null);

  const handleExistingMasterUpload = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const { existingNames, totalFound } = parseExistingItemMaster(buf);
      setExistingMaster({
        fileName: file.name,
        existingNames,
        totalInFile: totalFound,
        status: `Loaded ${totalFound} unique item name(s) from "${file.name}"`,
      });
    } catch (e) {
      setExistingMaster({
        fileName: "",
        existingNames: null,
        totalInFile: 0,
        status: `Error parsing file: ${e instanceof Error ? e.message : String(e)}`,
      });
    }
  }, []);

  const clearExistingMaster = useCallback(() => {
    setExistingMaster({
      fileName: "",
      existingNames: null,
      totalInFile: 0,
      status: null,
    });
    if (existingMasterInput.current) existingMasterInput.current.value = "";
  }, []);

  // ─── Sale conversion — runs only when ALL THREE files are loaded ──────────
  const triggerSaleConversionIfReady = useCallback(async (
    customBillWiseName?: string,
    customRegFileName?: string,
    customSRFileName?: string
  ) => {
    const billWiseName = customBillWiseName !== undefined ? customBillWiseName : sale.billWiseFileName;
    const regName = customRegFileName !== undefined ? customRegFileName : sale.regFileName;
    const srName = customSRFileName !== undefined ? customSRFileName : sale.srFileName;

    if (saleBuf.current && saleRegDiscMap.current && saleSRBuf.current) {
      setSale((s) => ({ ...s, stage: "processing", error: null }));
      try {
        await new Promise((r) => setTimeout(r, 30));
        const returnRegMap = saleSRBuf.current
          ? parseSalesReturnRegister(saleSRBuf.current)
          : (saleRegBuf.current ? parseSalesReturnRegister(saleRegBuf.current) : null);
        const { rows, stats } = convertSaleData(
          saleBuf.current,
          saleRegDiscMap.current,
          saleSRBuf.current,
          returnRegMap
        );
        datasetStore.setSale(rows, billWiseName);
        saleBlob.current = buildOutputWorkbooks(rows);
        setSale({
          stage: "ready",
          stats,
          error: null,
          billWiseFileName: billWiseName,
          regFileName: regName,
          srFileName: srName,
        });
        saveSaleToCloud(rows).catch((err) => console.error("Sale cloud save failed", err));
      } catch (e) {
        setSale((s) => ({ ...s, stage: "error", error: e instanceof Error ? e.message : String(e) }));
      }
    } else {
      setSale((s) => ({
        ...s,
        stage: "idle",
        stats: null,
        billWiseFileName: billWiseName,
        regFileName: regName,
        srFileName: srName,
      }));
    }
  }, [sale.billWiseFileName, sale.regFileName, sale.srFileName]);

  // ─── SR conversion — runs when SR file is loaded (register optional) ──────
  const triggerSRConversionIfReady = useCallback(async (
    customSRName?: string,
    customRegName?: string
  ) => {
    const srName = customSRName !== undefined ? customSRName : sr.srFileName;
    const regName = customRegName !== undefined ? customRegName : sr.regFileName;

    if (srBuf.current) {
      setSR((s) => ({ ...s, stage: "processing", error: null }));
      try {
        await new Promise((r) => setTimeout(r, 30));
        const discMap = srRegDiscMap.current ?? saleRegDiscMap.current;
        const mappedRegName = regName || sale.regFileName || "";
        const { rows, stats } = convertSaleReturn(srBuf.current, discMap);
        datasetStore.setReturn(rows, srName);
        srBlob.current = buildSRWorkbooks(rows);
        setSR({
          stage: "ready",
          stats,
          error: null,
          srFileName: srName,
          regFileName: mappedRegName,
        });
        saveSaleReturnToCloud(rows).catch((err) => console.error("SR cloud save failed", err));
      } catch (e) {
        setSR((s) => ({ ...s, stage: "error", error: e instanceof Error ? e.message : String(e) }));
      }
    } else {
      setSR((s) => ({
        ...s,
        stage: "idle",
        stats: null,
        srFileName: srName,
        regFileName: regName || s.regFileName,
      }));
    }
  }, [sr.srFileName, sr.regFileName, sale.regFileName]);

  // ─── Bill Wise Sales upload ────────────────────────────────────────────────
  const handleBillWiseSalesUpload = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const validation = validateSaleDataFile(buf);
      if (!validation.valid) {
        setSale((s) => ({
          ...s,
          stage: "error",
          error: validation.error || "Bill Wise Sales file invalid hai.",
          billWiseFileName: file.name,
        }));
        return;
      }
      saleBuf.current = buf;
      await triggerSaleConversionIfReady(file.name, undefined, undefined);
    } catch (e) {
      setSale((s) => ({ ...s, error: e instanceof Error ? e.message : String(e) }));
    }
  }, [triggerSaleConversionIfReady]);

  // ─── Sales Register upload ────────────────────────────────────────────────
  const handleSalesRegisterUpload = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const validation = validateSalesRegisterFile(buf);
      if (!validation.valid) {
        setSale((s) => ({
          ...s,
          stage: "error",
          error: validation.error || "Sales Register file invalid hai.",
          regFileName: file.name,
        }));
        return;
      }
      saleRegBuf.current = buf;
      const { discountMap, rows: regRows } = parseSalesRegister(buf);
      saleRegDiscMap.current = discountMap;
      datasetStore.setRegister(regRows, file.name);

      await triggerSaleConversionIfReady(undefined, file.name, undefined);
      await triggerSRConversionIfReady(undefined, file.name);
    } catch (e) {
      setSale((s) => ({ ...s, error: `Sales Register: ${e instanceof Error ? e.message : String(e)}` }));
    }
  }, [triggerSaleConversionIfReady, triggerSRConversionIfReady]);

  // ─── Sale Return upload (Linked for both Sale Data and Sale Return cards) ──
  const handleSaleSRUpload = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const validation = validateSaleReturnFile(buf);
      if (!validation.valid) {
        setSale((s) => ({
          ...s,
          stage: "error",
          error: validation.error || "Sale Return file invalid hai.",
          srFileName: file.name,
        }));
        return;
      }
      saleSRBuf.current = buf;
      srBuf.current = buf;

      await triggerSaleConversionIfReady(undefined, undefined, file.name);
      await triggerSRConversionIfReady(file.name, undefined);
    } catch (e) {
      setSale((s) => ({ ...s, error: `Sale Return file: ${e instanceof Error ? e.message : String(e)}` }));
    }
  }, [triggerSaleConversionIfReady, triggerSRConversionIfReady]);

  const handleSRUpload = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const validation = validateSaleReturnFile(buf);
      if (!validation.valid) {
        setSR((s) => ({
          ...s,
          stage: "error",
          error: validation.error || "Sale Return file invalid hai.",
          srFileName: file.name,
        }));
        return;
      }
      saleSRBuf.current = buf;
      srBuf.current = buf;

      await triggerSaleConversionIfReady(undefined, undefined, file.name);
      await triggerSRConversionIfReady(file.name, undefined);
    } catch (e) {
      setSR((s) => ({ ...s, error: `Sale Return: ${e instanceof Error ? e.message : String(e)}` }));
    }
  }, [triggerSaleConversionIfReady, triggerSRConversionIfReady]);

  const removeSaleSRFile = useCallback(async () => {
    saleSRBuf.current = null;
    srBuf.current = null;
    if (saleSRInput.current) saleSRInput.current.value = "";
    if (srInput.current) srInput.current.value = "";

    await triggerSaleConversionIfReady(undefined, undefined, "");
    await triggerSRConversionIfReady("", undefined);
  }, [triggerSaleConversionIfReady, triggerSRConversionIfReady]);

  // ─── Sales Return Register upload ─────────────────────────────────────────
  const handleSRRegisterUpload = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      srRegBuf.current = buf;
      const { discountMap } = parseSalesRegister(buf, true); // forReturn=true
      srRegDiscMap.current = discountMap;

      await triggerSRConversionIfReady(undefined, file.name);
    } catch (e) {
      setSR((s) => ({ ...s, error: `Sales Register: ${e instanceof Error ? e.message : String(e)}` }));
    }
  }, [triggerSRConversionIfReady]);

  const [itemMaster, setItemMaster] = useState<{ total: number; lastAdded: number; lastNames: string[] }>(() => ({
    total: typeof window !== "undefined" ? getMasterCount() : 0,
    lastAdded: 0,
    lastNames: [],
  }));

  const masterPurInput = useRef<HTMLInputElement>(null);
  const [masterPurStatus, setMasterPurStatus] = useState<string>("");

  const handleMasterPurUpload = useCallback(async (file: File) => {
    try {
      setMasterPurStatus(`Processing "${file.name}"...`);
      const buf = await file.arrayBuffer();
      const res = importPurchaseFileToMaster(buf);
      setItemMaster({ total: res.total, lastAdded: res.added, lastNames: res.newItems.slice(0, 20) });
      setMasterPurStatus(`✓ Purchase file imported: ${res.added} new items added, ${res.updated} updated. Total in DB: ${res.total}`);
    } catch (e) {
      setMasterPurStatus(`❌ Import failed: ${e instanceof Error ? e.message : String(e)}`);
    }
  }, []);

  const handlePurUpload = useCallback(async (file: File) => {
    setPur({ stage: "processing", stats: null, error: null, fileName: file.name });
    try {
      const buf = await file.arrayBuffer();
      await new Promise((r) => setTimeout(r, 30));
      const { rows, stats } = convertPurchase(buf);
      datasetStore.setPurchase(rows, file.name);
      purBlob.current = buildPurchaseWorkbook(rows);
      const merge = mergeFromPurchaseRows(rows);
      setItemMaster({ total: merge.total, lastAdded: merge.added, lastNames: merge.newItems.slice(0, 20) });
      setPur({ stage: "ready", stats, error: null, fileName: file.name });
      savePurchaseToCloud(rows).catch((err) => console.error("Purchase cloud save failed", err));
    } catch (e) {
      setPur({ stage: "error", stats: null, error: e instanceof Error ? e.message : String(e), fileName: file.name });
    }
  }, []);

  const onItemMasterDownload = useCallback(() => {
    const { buffer } = buildItemMasterWorkbook(existingMaster.existingNames ?? undefined);
    const prefix = existingMaster.existingNames ? "New_ItemMaster" : "ItemMaster";
    downloadBuffer(buffer, `${prefix}_${folderDate}`);
  }, [folderDate, existingMaster.existingNames]);

  const onItemMasterClear = useCallback(() => {
    if (!confirm("Item Master database delete karna hai? Sabhi saved items chale jayenge.")) return;
    clearMaster();
    setItemMaster({ total: 0, lastAdded: 0, lastNames: [] });
  }, []);

  const newItemsSummary = getNewItemsSummary(existingMaster.existingNames ?? undefined);

  const dayFolder = (date: string) => {
    const parts = date.split("-");
    return parts[parts.length - 1]?.replace(/^0/, "") ?? date;
  };

  const autoScanAndProcess = useCallback(async (
    parentHandle: FileSystemDirectoryHandle,
    date: string
  ) => {
    setAutoStatus("");
    const folderName = date.split("-").pop()?.replace(/^0/, "") ?? date;
    let subDir: FileSystemDirectoryHandle | null = null;
    try {
      subDir = await parentHandle.getDirectoryHandle(folderName);
    } catch {
      setAutoStatus(`📁 Folder "${folderName}" nahi mila — files manually chunein.`);
      return;
    }

    const found: { name: string; file: File }[] = [];
    for await (const [name, handle] of (subDir as unknown as { entries: () => AsyncIterable<[string, FileSystemHandle]> }).entries()) {
      if (handle.kind === "file" && (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls"))) {
        const file = await (handle as FileSystemFileHandle).getFile();
        found.push({ name, file });
      }
    }

    if (found.length === 0) {
      setAutoStatus(`📂 "${date}" folder mila lekin koi .xlsx file nahi — files manually chunein.`);
      return;
    }

    const names = found.map((f) => f.name).join(", ");
    setAutoStatus(`✅ "${folderName}" folder mein ${found.length} file(s) mili: ${names}`);

    // First pass: load registers so SR conversion gets Ushop data immediately
    for (const item of found) {
      const n = item.name.toLowerCase();
      if (n.includes("sales_register") || n.includes("salesregister") || n.includes("sales register")) {
        await handleSalesRegisterUpload(item.file);
        await handleSRRegisterUpload(item.file);
      }
    }
    // Second pass: load all other files
    for (const item of found) {
      const n = item.name.toLowerCase();
      if (n.includes("salesreturn") || n.includes("sales_return") || n.includes("salesreturndetail") || n.includes("return")) {
        await handleSRUpload(item.file);
        await handleSaleSRUpload(item.file);
      } else if (n.includes("purchase") || n.includes("product_wise_purchase") || n.includes("productwisepurchase")) {
        await handlePurUpload(item.file);
      } else if (n.includes("sales_register") || n.includes("salesregister") || n.includes("sales register")) {
        // Already processed in first pass — skip to avoid using Sales Register as Bill Wise Sales
      } else {
        await handleBillWiseSalesUpload(item.file);
      }
    }
  }, [handleBillWiseSalesUpload, handleSalesRegisterUpload, handleSRRegisterUpload, handleSRUpload, handleSaleSRUpload, handlePurUpload]);

  const saveToFolder = async (buf: ArrayBuffer, name: string, subfolderName: string) => {
    if (!dirHandle) return false;
    try {
      const subDir = await (dirHandle as FileSystemDirectoryHandle).getDirectoryHandle(subfolderName, { create: true });
      const fileHandle = await subDir.getFileHandle(name, { create: true });
      const writable = await (fileHandle as FileSystemFileHandle & { createWritable: () => Promise<FileSystemWritableFileStream> }).createWritable();
      await writable.write(buf);
      await writable.close();
      return true;
    } catch { return false; }
  };

  const onSaleDownload = useCallback(async () => {
    const bufs = saleBlob.current; if (!bufs || bufs.length === 0) return;
    if (dirHandle) {
      for (let i = 0; i < bufs.length; i++) {
        const suffix = bufs.length > 1 ? `_Part${i + 1}` : "";
        const ok = await saveToFolder(bufs[i], `Converted_Sale_${folderDate}${suffix}.xlsx`, folderDate);
        if (!ok) return;
      }
    } else if (bufs.length > 1) {
      // Multiple parts (e.g. 25+ files) - pack into 1-click ZIP so browser never drops any file
      const files = bufs.map((b, i) => ({
        buf: b,
        fileName: `Converted_Sale_${folderDate}_Part${i + 1}.xlsx`,
      }));
      await downloadZipBundle(files, `Converted_Sale_${folderDate}_All_${bufs.length}_Parts.zip`);
    } else {
      downloadBuffer(bufs[0], `Converted_Sale_${folderDate}`);
    }
  }, [folderDate, dirHandle]);

  const onSaleDownloadAllSeparate = useCallback(async () => {
    const bufs = saleBlob.current; if (!bufs || bufs.length === 0) return;
    const files = bufs.map((b, i) => ({
      buf: b,
      fileName: `Converted_Sale_${folderDate}${bufs.length > 1 ? `_Part${i + 1}` : ""}`,
    }));
    await downloadMultipleBuffers(files, 350);
  }, [folderDate]);

  const onSaleDownloadPart = useCallback((idx: number) => {
    const bufs = saleBlob.current; if (!bufs || !bufs[idx]) return;
    const suffix = bufs.length > 1 ? `_Part${idx + 1}` : "";
    downloadBuffer(bufs[idx], `Converted_Sale_${folderDate}${suffix}`);
  }, [folderDate]);

  const onSRDownload = useCallback(async () => {
    const bufs = srBlob.current; if (!bufs || bufs.length === 0) return;
    if (dirHandle) {
      for (let i = 0; i < bufs.length; i++) {
        const suffix = bufs.length > 1 ? `_Part${i + 1}` : "";
        const ok = await saveToFolder(bufs[i], `Sale_Return_${folderDate}${suffix}.xlsx`, folderDate);
        if (!ok) return;
      }
    } else if (bufs.length > 1) {
      // Multiple parts - pack into 1-click ZIP
      const files = bufs.map((b, i) => ({
        buf: b,
        fileName: `Sale_Return_${folderDate}_Part${i + 1}.xlsx`,
      }));
      await downloadZipBundle(files, `Sale_Return_${folderDate}_All_${bufs.length}_Parts.zip`);
    } else {
      downloadBuffer(bufs[0], `Sale_Return_${folderDate}`);
    }
  }, [folderDate, dirHandle]);

  const onSRDownloadAllSeparate = useCallback(async () => {
    const bufs = srBlob.current; if (!bufs || bufs.length === 0) return;
    const files = bufs.map((b, i) => ({
      buf: b,
      fileName: `Sale_Return_${folderDate}${bufs.length > 1 ? `_Part${i + 1}` : ""}`,
    }));
    await downloadMultipleBuffers(files, 350);
  }, [folderDate]);

  const onSRDownloadPart = useCallback((idx: number) => {
    const bufs = srBlob.current; if (!bufs || !bufs[idx]) return;
    const suffix = bufs.length > 1 ? `_Part${idx + 1}` : "";
    downloadBuffer(bufs[idx], `Sale_Return_${folderDate}${suffix}`);
  }, [folderDate]);

  const onPurDownload = useCallback(async () => {
    const buf = purBlob.current; if (!buf) return;
    if (dirHandle) {
      const ok = await saveToFolder(buf, `Purchase_${folderDate}.xlsx`, folderDate);
      if (!ok) return;
    } else { downloadBuffer(buf, `Purchase_${folderDate}`); }
  }, [folderDate, dirHandle]);

  const resetSale = () => {
    saleBlob.current = null;
    saleBuf.current = null;
    saleRegBuf.current = null;
    saleSRBuf.current = null;
    saleRegDiscMap.current = null;
    setSale({ stage: "idle", stats: null, error: null, billWiseFileName: "", regFileName: "", srFileName: "" });
    if (billWiseInput.current) billWiseInput.current.value = "";
    if (regInput.current) regInput.current.value = "";
    if (saleSRInput.current) saleSRInput.current.value = "";
  };
  const resetSR  = () => {
    srBlob.current = null; srBuf.current = null; srRegBuf.current = null; srRegDiscMap.current = null;
    setSR({ stage: "idle", stats: null, error: null, srFileName: "", regFileName: "" });
    if (srInput.current) srInput.current.value = "";
    if (srRegInput.current) srRegInput.current.value = "";
  };
  const resetPur = () => { purBlob.current = null; setPur ({ stage: "idle", stats: null, error: null, fileName: "" }); if (purInput.current) purInput.current.value = ""; };

  const pickFolder = async () => {
    if (!("showDirectoryPicker" in window)) { alert("Aapka browser folder selection support nahi karta. Chrome ya Edge use karein."); return; }
    try {
      const handle = await (window as unknown as { showDirectoryPicker: (opts?: object) => Promise<FileSystemDirectoryHandle> }).showDirectoryPicker({ mode: "readwrite" });
      setDirHandle(handle);
      await autoScanAndProcess(handle, folderDate);
    } catch { /* user cancelled */ }
  };

  const handleDateChange = async (v: string) => {
    setFolderDate(v);
    try { localStorage.setItem("xlsxConverterDate", v); } catch {}
    resetSale(); resetSR(); resetPur();
    setAutoStatus("");
    if (dirHandle) await autoScanAndProcess(dirHandle, v);
  };

  const dirName = dirHandle ? (dirHandle as FileSystemDirectoryHandle & { name: string }).name : null;

  return (
    <div className="min-h-screen bg-background text-foreground">
      <div className="mx-auto max-w-6xl px-4 py-3">
        <header className="mb-3 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-3">
            <div className="inline-flex items-center gap-1.5 rounded-full border border-border bg-muted/40 px-2.5 py-0.5 text-xs text-muted-foreground">
              <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
              100% browser · no upload · no storage
            </div>
            <h1 className="text-lg font-semibold tracking-tight">Excel Converter</h1>
          </div>
        </header>

        <section className="mb-3 grid gap-3 sm:grid-cols-2">
          <div className="block">
            <span className="text-sm font-medium">Date chunein</span>
            <div className="mt-1 flex items-center gap-2">
              <input
                type="date"
                className="flex-1 rounded-md border border-input bg-background px-3 py-2 text-sm"
                value={folderDate}
                onChange={(e) => handleDateChange(e.target.value || todayFolderName())}
              />
              <button
                type="button"
                onClick={pickFolder}
                title="Is date ke naam wale subfolder se files auto-load hongi"
                className={`inline-flex items-center gap-1.5 rounded-md border px-3 py-2 text-sm font-medium transition whitespace-nowrap ${
                  dirHandle ? "border-emerald-500/40 bg-emerald-500/10 text-emerald-700 hover:bg-emerald-500/20" : "border-input bg-background text-foreground hover:bg-accent"
                }`}
              >
                <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" />
                </svg>
                {dirName ?? "📁 Folder Chunein"}
              </button>
            </div>
            {autoStatus && <p className="mt-2 text-xs text-blue-600 leading-relaxed">{autoStatus}</p>}
            {dirName && <p className="mt-1 text-xs text-emerald-600">Scan: <span className="font-mono">{dirName}/{dayFolder(folderDate)}/</span></p>}
          </div>

          <div className="flex flex-col justify-end gap-1 text-xs text-muted-foreground">
            <code className="truncate rounded-md border border-input bg-muted/40 px-3 py-2">Converted_Sale_{folderDate}.xlsx</code>
            <code className="truncate rounded-md border border-input bg-muted/40 px-3 py-2">Sale_Return_{folderDate}.xlsx</code>
            <code className="truncate rounded-md border border-input bg-muted/40 px-3 py-2">Purchase_{folderDate}.xlsx</code>
          </div>
        </section>

        <div className="grid gap-4 lg:grid-cols-3">
          {/* ── Sale Data card — two file uploads ── */}
          <div className="rounded-2xl border border-border bg-card p-5">
            <div className="mb-4">
              <h2 className="text-lg font-semibold">Sale Data</h2>
              <p className="text-xs text-muted-foreground">Bill Wise Sales + Sales Register → Sale Item import</p>
            </div>

            {/* File slot 1 — Bill Wise Sales (required) */}
            <FileSlot
              label="Bill Wise Sales"
              required
              hint="LeverEDGE_…_Bill Wise Sales_…"
              fileName={sale.billWiseFileName}
              inputRef={billWiseInput}
              disabled={sale.stage === "processing"}
              onUpload={handleBillWiseSalesUpload}
            />

            {/* File slot 2 — Sales Register (required, for Ushop/Shikhar discounts) */}
            <div className="mt-3">
              <FileSlot
                label="Sales Register"
                required
                hint="LeverEDGE_…_Sales_Register_…"
                fileName={sale.regFileName}
                inputRef={regInput}
                disabled={sale.stage === "processing"}
                onUpload={handleSalesRegisterUpload}
              />
            </div>

            {/* File slot 3 — Sale Return (required, to merge missing bills into sale data) */}
            <div className="mt-3">
              <FileSlot
                label="Sale Return File"
                required
                hint="LeverEDGE_…_SalesReturn_… (Isme se missing bills merge honge)"
                fileName={sale.srFileName}
                inputRef={saleSRInput}
                disabled={sale.stage === "processing"}
                onUpload={handleSaleSRUpload}
                onClear={removeSaleSRFile}
              />
            </div>

            {/* Status area */}
            <div className="mt-4">
              {sale.stage === "processing" && (
                <div className="flex items-center gap-2 text-sm text-muted-foreground">
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-primary border-t-transparent" />
                  Converting…
                </div>
              )}

              {/* Waiting state — checklist of required files */}
              {sale.stage === "idle" && (sale.billWiseFileName || sale.regFileName || sale.srFileName) && !sale.error && (
                <div className="rounded-md border border-amber-400/40 bg-amber-50/60 px-3 py-2.5 text-xs text-amber-800">
                  <p className="font-semibold mb-1">Mapping start karne ke liye sabhi 3 files select hona zaroori hai:</p>
                  <ul className="list-none space-y-1 mt-1">
                    <li className="flex items-center gap-1.5">
                      {sale.billWiseFileName ? (
                        <span className="text-emerald-600 font-medium">✓ Bill Wise Sales loaded</span>
                      ) : (
                        <span className="text-amber-700">➜ Bill Wise Sales chunein</span>
                      )}
                    </li>
                    <li className="flex items-center gap-1.5">
                      {sale.regFileName ? (
                        <span className="text-emerald-600 font-medium">✓ Sales Register loaded</span>
                      ) : (
                        <span className="text-amber-700">➜ Sales Register chunein</span>
                      )}
                    </li>
                    <li className="flex items-center gap-1.5">
                      {sale.srFileName ? (
                        <span className="text-emerald-600 font-medium">✓ Sale Return loaded</span>
                      ) : (
                        <span className="text-amber-700">➜ Sale Return chunein</span>
                      )}
                    </li>
                  </ul>
                </div>
              )}

              {sale.stage === "ready" && (
                <div>
                  <p className="text-sm font-medium text-emerald-600">✅ Conversion ready</p>
                  {sale.stats && (
                    <>
                      <div className="mt-3 grid grid-cols-2 gap-2">
                        <Stat label="Source rows" value={sale.stats.sourceRows} />
                        <Stat label="Export rows" value={sale.stats.exportRows} accent />
                        <Stat label="Removed (Net&lt;0)" value={sale.stats.removedNegative} />
                        <Stat label="Discount bills" value={sale.stats.discountBillsMatched} accent />
                        {(sale.stats.srMissingBillsAdded ?? 0) > 0 && (
                          <Stat label="Added from SR" value={`${sale.stats.srMissingBillsAdded} bills (${sale.stats.srMissingRowsAdded} rows)`} accent />
                        )}
                        {(sale.stats.zeroQtyRebuilt ?? 0) > 0 && (
                          <Stat label="Zero-Qty Rebuilt" value={`${sale.stats.zeroQtyRebuilt} rows`} accent />
                        )}
                      </div>

                      {(sale.stats.zeroQtyRebuilt ?? 0) > 0 && (
                        <div className="mt-2 rounded-md border border-emerald-500/30 bg-emerald-500/10 p-2 text-xs text-emerald-700 dark:text-emerald-300">
                          ✓ {sale.stats.zeroQtyRebuilt} zero-unit returned item(s) Sales Return Register se value & rate calculate karke rebuild kiye gaye.
                        </div>
                      )}

                      {(sale.stats.srMissingBillsAdded ?? 0) > 0 && (
                        <div className="mt-2 rounded-md border border-emerald-500/30 bg-emerald-500/10 p-2 text-xs text-emerald-700 dark:text-emerald-300">
                          ℹ {sale.stats.srMissingBillsAdded} missing bill(s) ({sale.stats.srMissingRowsAdded} items) Sale Return file se Sale Data me add kiye gaye.
                        </div>
                      )}

                      {/* Discount bill list */}
                      {sale.stats.discountBills && sale.stats.discountBills.length > 0 && (
                        <div className="mt-3">
                          <p className="text-xs font-medium text-muted-foreground mb-1">
                            Discount Breakdown (USHOP DIS = BTPR SchDisc + Ushop Redemption | SHIKHAR DIS):
                          </p>
                          <div className="rounded-md border border-border bg-muted/30 max-h-44 overflow-y-auto">
                            <table className="w-full text-xs">
                              <thead className="sticky top-0 bg-muted/80">
                                <tr>
                                  <th className="px-2 py-1 text-left font-medium text-muted-foreground">Bill No</th>
                                  <th className="px-2 py-1 text-right font-medium text-muted-foreground">BTPR</th>
                                  <th className="px-2 py-1 text-right font-medium text-muted-foreground">Ushop</th>
                                  <th className="px-2 py-1 text-right font-semibold text-emerald-600 dark:text-emerald-400">USHOP DIS (A/L 1)</th>
                                  <th className="px-2 py-1 text-right font-medium text-muted-foreground">SHIKHAR DIS (A/L 2)</th>
                                </tr>
                              </thead>
                              <tbody>
                                {sale.stats.discountBills.map((d, i) => {
                                  const add1 = d.totalAddLess1 ?? ((d.btpr || 0) + (d.ushop || 0));
                                  return (
                                    <tr key={i} className="border-t border-border/50">
                                      <td className="px-2 py-0.5 font-mono">{d.bill}</td>
                                      <td className="px-2 py-0.5 text-right">{(d.btpr ?? 0) > 0 ? (d.btpr!).toFixed(2) : "—"}</td>
                                      <td className="px-2 py-0.5 text-right">{d.ushop > 0 ? d.ushop.toFixed(2) : "—"}</td>
                                      <td className="px-2 py-0.5 text-right font-semibold text-emerald-600 dark:text-emerald-400">
                                        {add1 > 0 ? add1.toFixed(2) : "—"}
                                      </td>
                                      <td className="px-2 py-0.5 text-right">{d.shikhar > 0 ? d.shikhar.toFixed(2) : "—"}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {sale.stats.discountBills && sale.stats.discountBills.length === 0 && (
                        <p className="mt-2 text-xs text-muted-foreground">⚠ Kisi bhi bill mein discount nahi mila.</p>
                      )}
                    </>
                  )}
                  {(() => {
                    const partCount = saleBlob.current?.length || 1;
                    return (
                      <div className="mt-4 space-y-2">
                        <div className="flex flex-wrap gap-2">
                          <button
                            type="button"
                            onClick={onSaleDownload}
                            className="inline-flex items-center justify-center gap-1.5 rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 shadow-sm"
                          >
                            {partCount > 1 ? `📦 Download All in 1 ZIP (${partCount} Files)` : "⬇ Download Converted File"}
                          </button>
                          {partCount > 1 && (
                            <button
                              type="button"
                              onClick={onSaleDownloadAllSeparate}
                              className="inline-flex items-center justify-center gap-1.5 rounded-md border border-primary/40 bg-primary/10 px-3 py-2 text-xs font-medium text-primary hover:bg-primary/20"
                              title="Downloads each part .xlsx with automated delay"
                            >
                              ⬇ Download {partCount} Separate Files
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={resetSale}
                            className="inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent"
                          >
                            Reset
                          </button>
                        </div>

                        {partCount > 1 && (
                          <div className="pt-1.5">
                            <p className="text-[11px] font-medium text-muted-foreground mb-1">
                              Ya koi bhi single part download karein ({partCount} parts):
                            </p>
                            <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto p-1.5 bg-muted/20 rounded border border-border/50">
                              {Array.from({ length: partCount }, (_, idx) => (
                                <button
                                  key={idx}
                                  type="button"
                                  onClick={() => onSaleDownloadPart(idx)}
                                  className="px-2 py-0.5 text-xs font-medium rounded border border-input bg-background hover:bg-accent transition-colors"
                                >
                                  Part {idx + 1}
                                </button>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              )}

              {sale.stage === "downloaded" && (
                <div>
                  <p className="text-sm font-medium text-emerald-600">Downloaded. Memory cleared.</p>
                  <button type="button" onClick={resetSale}
                    className="mt-3 inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent">
                    Convert another file
                  </button>
                </div>
              )}

              {sale.error && (
                <div className="rounded-md border border-destructive/30 bg-destructive/5 p-3 text-xs text-destructive">
                  <p className="font-medium">Conversion failed</p>
                  <p className="mt-1">{sale.error}</p>
                  <button onClick={resetSale} className="mt-2 underline">Try again</button>
                </div>
              )}
            </div>
          </div>

          {/* ── Sale Return card — two file uploads ── */}
          <div className="rounded-2xl border border-border bg-card p-5">
            <div className="mb-4">
              <h2 className="text-lg font-semibold">Sale Return</h2>
              <p className="text-xs text-muted-foreground">LeverEdge SR + Sales Register → Sale Return Accounting import</p>
            </div>

            {/* File slot 1 — SR file (required) */}
            <FileSlot
              label="Sale Return File"
              required
              hint="LeverEDGE_…_SalesReturn_…"
              fileName={sr.srFileName}
              inputRef={srInput}
              disabled={sr.stage === "processing"}
              onUpload={handleSRUpload}
              onClear={removeSaleSRFile}
            />

            {/* File slot 2 — Sales Register (for Ushop discounts) */}
            <div className="mt-3">
              <FileSlot
                label="Sales Register"
                hint="LeverEDGE_…_Sales_Register_… (Ushop amount ke liye)"
                fileName={sr.regFileName}
                inputRef={srRegInput}
                disabled={sr.stage === "processing"}
                onUpload={handleSRRegisterUpload}
              />
            </div>

            {/* Status area */}
            <div className="mt-4">
              {sr.stage === "processing" && (
                <div className="flex items-center gap-2 text-sm text-muted-foreground">
                  <div className="h-4 w-4 animate-spin rounded-full border-2 border-primary border-t-transparent" />
                  Converting…
                </div>
              )}

              {sr.stage === "ready" && (
                <div>
                  <p className="text-sm font-medium text-emerald-600">✅ Conversion ready</p>
                  {sr.stats && (
                    <>
                      <div className="mt-3 grid grid-cols-2 gap-2">
                        <Stat label="Source rows" value={sr.stats.sourceRows} />
                        <Stat label="Export rows" value={sr.stats.exportRows} accent />
                        <Stat label="Invalid" value={sr.stats.rejectedInvalid} />
                        <Stat label="Discount bills" value={sr.stats.discountBillsMatched} accent />
                      </div>

                      {/* Discount bill list */}
                      {sr.stats.discountBills && sr.stats.discountBills.length > 0 && (
                        <div className="mt-3">
                          <p className="text-xs font-medium text-muted-foreground mb-1">
                            Discount Breakdown (USHOP DIS = BTPR SchDisc + Ushop Redemption | SHIKHAR DIS):
                          </p>
                          <div className="rounded-md border border-border bg-muted/30 max-h-44 overflow-y-auto">
                            <table className="w-full text-xs">
                              <thead className="sticky top-0 bg-muted/80">
                                <tr>
                                  <th className="px-2 py-1 text-left font-medium text-muted-foreground">Bill No</th>
                                  <th className="px-2 py-1 text-right font-medium text-muted-foreground">BTPR</th>
                                  <th className="px-2 py-1 text-right font-medium text-muted-foreground">Ushop</th>
                                  <th className="px-2 py-1 text-right font-semibold text-emerald-600 dark:text-emerald-400">USHOP DIS (A/L 1)</th>
                                  <th className="px-2 py-1 text-right font-medium text-muted-foreground">SHIKHAR DIS (A/L 2)</th>
                                </tr>
                              </thead>
                              <tbody>
                                {sr.stats.discountBills.map((d, i) => {
                                  const add1 = d.totalAddLess1 ?? ((d.btpr || 0) + (d.ushop || 0));
                                  return (
                                    <tr key={i} className="border-t border-border/50">
                                      <td className="px-2 py-0.5 font-mono">{d.bill}</td>
                                      <td className="px-2 py-0.5 text-right">{(d.btpr ?? 0) > 0 ? (d.btpr!).toFixed(2) : "—"}</td>
                                      <td className="px-2 py-0.5 text-right">{d.ushop > 0 ? d.ushop.toFixed(2) : "—"}</td>
                                      <td className="px-2 py-0.5 text-right font-semibold text-emerald-600 dark:text-emerald-400">
                                        {add1 > 0 ? add1.toFixed(2) : "—"}
                                      </td>
                                      <td className="px-2 py-0.5 text-right">{d.shikhar > 0 ? d.shikhar.toFixed(2) : "—"}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {sr.stats.discountBills && sr.stats.discountBills.length === 0 && (
                        <p className="mt-2 text-xs text-muted-foreground">⚠ Kisi bhi bill mein discount nahi mila.</p>
                      )}
                    </>
                  )}
                  {(() => {
                    const partCount = srBlob.current?.length || 1;
                    return (
                      <div className="mt-4 space-y-2">
                        <div className="flex flex-wrap gap-2">
                          <button
                            type="button"
                            onClick={onSRDownload}
                            className="inline-flex items-center justify-center gap-1.5 rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 shadow-sm"
                          >
                            {partCount > 1 ? `📦 Download All in 1 ZIP (${partCount} Files)` : "⬇ Download Converted File"}
                          </button>
                          {partCount > 1 && (
                            <button
                              type="button"
                              onClick={onSRDownloadAllSeparate}
                              className="inline-flex items-center justify-center gap-1.5 rounded-md border border-primary/40 bg-primary/10 px-3 py-2 text-xs font-medium text-primary hover:bg-primary/20"
                              title="Downloads each part .xlsx with automated delay"
                            >
                              ⬇ Download {partCount} Separate Files
                            </button>
                          )}
                          <button
                            type="button"
                            onClick={resetSR}
                            className="inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent"
                          >
                            Reset
                          </button>
                        </div>

                        {partCount > 1 && (
                          <div className="pt-1.5">
                            <p className="text-[11px] font-medium text-muted-foreground mb-1">
                              Ya koi bhi single part download karein ({partCount} parts):
                            </p>
                            <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto p-1.5 bg-muted/20 rounded border border-border/50">
                              {Array.from({ length: partCount }, (_, idx) => (
                                <button
                                  key={idx}
                                  type="button"
                                  onClick={() => onSRDownloadPart(idx)}
                                  className="px-2 py-0.5 text-xs font-medium rounded border border-input bg-background hover:bg-accent transition-colors"
                                >
                                  Part {idx + 1}
                                </button>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </div>
              )}

              {sr.stage === "downloaded" && (
                <div>
                  <p className="text-sm font-medium text-emerald-600">Downloaded. Memory cleared.</p>
                  <button type="button" onClick={resetSR}
                    className="mt-3 inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent">
                    Convert another file
                  </button>
                </div>
              )}

              {sr.error && (
                <div className="rounded-md border border-destructive/30 bg-destructive/5 p-3 text-xs text-destructive">
                  <p className="font-medium">Conversion failed</p>
                  <p className="mt-1">{sr.error}</p>
                  <button onClick={resetSR} className="mt-2 underline">Try again</button>
                </div>
              )}
            </div>
          </div>


          <UploadCard
            title="Purchase"
            subtitle="LeverEdge Product Wise Purchase → Purchase Item import"
            inputRef={purInput}
            stage={pur.stage} fileName={pur.fileName} error={pur.error}
            onUpload={handlePurUpload} onDownload={onPurDownload} onReset={resetPur}
          >
            {pur.stats && (
              <div className="grid grid-cols-2 gap-2">
                <Stat label="Source rows" value={pur.stats.sourceRows} />
                <Stat label="Export rows" value={pur.stats.exportRows} accent />
                <Stat label="Invalid" value={pur.stats.rejectedInvalid} />
              </div>
            )}
          </UploadCard>
        </div>

        <section className="mt-4 rounded-2xl border border-border bg-card p-5">
          <div className="flex flex-wrap items-start justify-between gap-3">
            <div>
              <h2 className="text-lg font-semibold">📦 Item Master Database & Differential Export</h2>
              <p className="text-xs text-muted-foreground">
                Item Master ka saara data <strong>Purchase File</strong> (LeverEdge Product Wise Purchase) se liya jaata hai (HSN Code, GST %, MRP, Purchase Price/Rate) aur bina kisi duplicate item ke save hota hai.
              </p>
            </div>
            <div className="flex flex-wrap gap-2">
              <button
                type="button"
                onClick={onItemMasterDownload}
                disabled={newItemsSummary.newCount === 0}
                className="inline-flex items-center justify-center rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 disabled:opacity-50"
              >
                ⬇ Download {existingMaster.existingNames ? `New Items Master (${newItemsSummary.newCount})` : `Item Master (${itemMaster.total})`}
              </button>
              <button
                type="button"
                onClick={onItemMasterClear}
                disabled={itemMaster.total === 0}
                className="inline-flex items-center justify-center rounded-md border border-destructive/40 bg-destructive/5 px-3 py-1.5 text-xs font-medium text-destructive hover:bg-destructive/10 disabled:opacity-50"
              >
                Clear DB
              </button>
            </div>
          </div>

          {/* Direct Purchase File Feed Slot */}
          <div className="mt-4 rounded-xl border border-emerald-500/30 bg-emerald-500/5 p-4">
            <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
              <div>
                <h3 className="text-sm font-medium text-emerald-950 dark:text-emerald-200">
                  📥 Feed Item Master from Purchase File
                </h3>
                <p className="text-xs text-emerald-800/80 dark:text-emerald-300/80">
                  Product Wise Purchase Excel upload karein. Sabhi items ke HSN, GST %, MRP aur Purchase Rate bina kisi duplicate ke database me update ho jayenge.
                </p>
              </div>
            </div>

            <FileSlot
              label="Purchase File (Product Wise Purchase / Item Import)"
              hint="LeverEDGE_Product_Wise_Purchase.xlsx"
              fileName=""
              inputRef={masterPurInput}
              onUpload={handleMasterPurUpload}
              accept=".xlsx,.xls,.csv"
            />

            {masterPurStatus && (
              <p className="mt-2 text-xs font-medium text-emerald-700 dark:text-emerald-300">
                {masterPurStatus}
              </p>
            )}
          </div>

          {/* Upload Existing Item Master Slot */}
          <div className="mt-4 rounded-xl border border-border/80 bg-muted/20 p-4">
            <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
              <div>
                <h3 className="text-sm font-medium">Upload Existing Item Master (Filter / Match)</h3>
                <p className="text-xs text-muted-foreground">
                  Apne Tally ya Accounting software ki current Item Master file upload karein. Jo items isme honge, wo download se exclude ho jayenge aur sirf <strong>NEW items</strong> hi download honge.
                </p>
              </div>
              {existingMaster.fileName && (
                <button
                  type="button"
                  onClick={clearExistingMaster}
                  className="text-xs text-destructive hover:underline"
                >
                  Remove Upload Filter
                </button>
              )}
            </div>

            <FileSlot
              label="Existing Item Master (Excel / CSV)"
              hint="Tally / Existing_Item_Master.xlsx / .csv"
              fileName={existingMaster.fileName}
              inputRef={existingMasterInput}
              onUpload={handleExistingMasterUpload}
              accept=".xlsx,.xls,.csv"
            />

            {existingMaster.status && (
              <p className="mt-2 text-xs font-medium text-emerald-600 dark:text-emerald-400">
                ✓ {existingMaster.status}
              </p>
            )}
          </div>

          {/* Stat indicators */}
          <div className="mt-4 grid grid-cols-2 gap-2 sm:grid-cols-4">
            <Stat label="Total items in DB" value={itemMaster.total} />
            <Stat label="Added this session" value={itemMaster.lastAdded} />
            <Stat
              label="Existing in Uploaded Master"
              value={existingMaster.existingNames ? newItemsSummary.existingCount : "— (No filter)"}
            />
            <Stat
              label="New items to export"
              value={newItemsSummary.newCount}
              accent
            />
          </div>

          {/* List of items that will be downloaded */}
          {newItemsSummary.newItems.length > 0 && (
            <div className="mt-4">
              <div className="flex items-center justify-between mb-1.5">
                <p className="text-xs font-medium text-muted-foreground">
                  {existingMaster.existingNames
                    ? `New Items ready for export (${newItemsSummary.newCount} items):`
                    : `All Unique Items in Database (${itemMaster.total} items):`}
                </p>
                <span className="text-[11px] text-muted-foreground">Strictly deduplicated by item name</span>
              </div>
              <div className="rounded-md border border-border bg-muted/30 max-h-52 overflow-y-auto">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-muted/80 backdrop-blur-xs">
                    <tr>
                      <th className="px-2.5 py-1.5 text-left font-medium text-muted-foreground">#</th>
                      <th className="px-2.5 py-1.5 text-left font-medium text-muted-foreground">Item Name</th>
                      <th className="px-2.5 py-1.5 text-left font-medium text-muted-foreground">HSN Code</th>
                      <th className="px-2.5 py-1.5 text-right font-medium text-muted-foreground">GST %</th>
                      <th className="px-2.5 py-1.5 text-right font-medium text-muted-foreground">MRP</th>
                      <th className="px-2.5 py-1.5 text-right font-medium text-muted-foreground">Purchase Rate</th>
                    </tr>
                  </thead>
                  <tbody>
                    {newItemsSummary.newItems.slice(0, 50).map((it, i) => (
                      <tr key={i} className="border-t border-border/50 hover:bg-muted/40">
                        <td className="px-2.5 py-1 text-muted-foreground">{i + 1}</td>
                        <td className="px-2.5 py-1 font-medium">{it.name}</td>
                        <td className="px-2.5 py-1 font-mono text-emerald-600 dark:text-emerald-400">{it.hsn || "—"}</td>
                        <td className="px-2.5 py-1 text-right">{it.gstRate ? `${it.gstRate}%` : "—"}</td>
                        <td className="px-2.5 py-1 text-right">{it.mrp ? `₹${it.mrp}` : "—"}</td>
                        <td className="px-2.5 py-1 text-right">{it.purchasePrice ? `₹${it.purchasePrice}` : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {newItemsSummary.newItems.length > 50 && (
                  <div className="p-2 text-center text-xs text-muted-foreground border-t border-border/50">
                    …and {newItemsSummary.newItems.length - 50} more items will be included in download
                  </div>
                )}
              </div>
            </div>
          )}

          {existingMaster.existingNames && newItemsSummary.newCount === 0 && itemMaster.total > 0 && (
            <div className="mt-3 rounded-md border border-emerald-500/30 bg-emerald-500/10 p-3 text-xs text-emerald-700 dark:text-emerald-300">
              🎉 All {itemMaster.total} items in database already exist in the uploaded Master file. No new items need to be imported!
            </div>
          )}
        </section>

        <footer className="mt-4 border-t border-border pt-3 text-xs text-muted-foreground">
          <p className="font-medium text-foreground">Pipelines</p>
          <ul className="mt-2 grid gap-1 sm:grid-cols-3">
            <li>• Sale: Net&lt;0 removed, HSN cleaned, Discount→T, Ushop/Shikhar discount (1st row/bill)</li>
            <li>• Return: SR→CN No, GST auto-detect, Discount total, Classification</li>
            <li>• Purchase: Rate=Price÷UPC, GST=CGST+SGST+IGST, Voucher auto-seq</li>
          </ul>
        </footer>
      </div>
    </div>
  );
}

// ── FileSlot: single drag-drop / click upload zone ────────────────────────
function FileSlot({
  label,
  hint,
  required,
  fileName,
  inputRef,
  disabled,
  onUpload,
  onClear,
  accept = ".xlsx,.xls",
}: {
  label: string;
  hint?: string;
  required?: boolean;
  fileName: string;
  inputRef: React.RefObject<HTMLInputElement | null>;
  disabled?: boolean;
  onUpload: (f: File) => void;
  onClear?: () => void;
  accept?: string;
}) {
  return (
    <div
      className={`rounded-xl border-2 border-dashed px-4 py-3 transition ${
        disabled ? "opacity-50" : "hover:border-primary/50"
      } ${fileName ? "border-emerald-400/50 bg-emerald-500/5" : "border-border"}`}
      onDragOver={(e) => { if (!disabled) e.preventDefault(); }}
      onDrop={(e) => {
        e.preventDefault();
        if (disabled) return;
        const f = e.dataTransfer.files?.[0];
        if (f) onUpload(f);
      }}
    >
      <div className="flex items-center justify-between gap-2">
        <div className="min-w-0">
          <p className="text-xs font-medium">
            {label}
            {required && <span className="ml-1 text-destructive">*</span>}
          </p>
          {fileName ? (
            <p className="mt-0.5 truncate text-xs text-emerald-600 font-mono">{fileName}</p>
          ) : (
            hint && <p className="mt-0.5 truncate text-xs text-muted-foreground">{hint}</p>
          )}
        </div>
        <div className="flex items-center gap-1.5 shrink-0">
          {fileName && onClear && (
            <button
              type="button"
              disabled={disabled}
              onClick={onClear}
              className="inline-flex items-center justify-center rounded-md border border-input bg-background px-2 py-1.5 text-xs font-medium text-muted-foreground hover:bg-destructive/10 hover:text-destructive disabled:opacity-50"
              title="Remove file"
            >
              Clear
            </button>
          )}
          <button
            type="button"
            disabled={disabled}
            onClick={() => inputRef.current?.click()}
            className="inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent disabled:opacity-50"
          >
            {fileName ? "Change" : "Choose"}
          </button>
        </div>
        <input
          ref={inputRef}
          type="file"
          accept={accept}
          className="hidden"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) onUpload(f); }}
        />
      </div>
    </div>
  );
}

// ── UploadCard: used by Sale Return & Purchase ────────────────────────────
function UploadCard({
  title, subtitle, inputRef, stage, fileName, error, onUpload, onDownload, onReset, children,
}: {
  title: string; subtitle: string;
  inputRef: React.RefObject<HTMLInputElement | null>;
  stage: Stage; fileName: string; error: string | null;
  onUpload: (f: File) => void; onDownload: () => void; onReset: () => void;
  children?: React.ReactNode;
}) {
  return (
    <div className="rounded-2xl border border-border bg-card p-5">
      <div className="mb-4">
        <h2 className="text-lg font-semibold">{title}</h2>
        <p className="text-xs text-muted-foreground">{subtitle}</p>
      </div>
      <div
        className={`rounded-xl border-2 border-dashed p-8 text-center transition ${stage === "processing" ? "border-primary/40" : "border-border"}`}
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) onUpload(f); }}
      >
        {stage === "idle" || stage === "error" ? (
          <>
            <p className="text-sm text-muted-foreground">Drop .xlsx here, or</p>
            <button type="button" onClick={() => inputRef.current?.click()}
              className="mt-3 inline-flex items-center justify-center rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground transition hover:bg-primary/90">
              Choose File
            </button>
            <input ref={inputRef} type="file" accept=".xlsx,.xls" className="hidden"
              onChange={(e) => { const f = e.target.files?.[0]; if (f) onUpload(f); }} />
          </>
        ) : stage === "processing" ? (
          <div className="text-sm text-muted-foreground">
            <div className="mx-auto mb-3 h-6 w-6 animate-spin rounded-full border-2 border-primary border-t-transparent" />
            Converting <span className="font-mono">{fileName}</span>…
          </div>
        ) : stage === "ready" ? (
          <div>
            <p className="text-sm font-medium text-emerald-600">✅ Conversion ready</p>
            <p className="mt-1 text-xs text-muted-foreground font-mono">{fileName}</p>
            <button type="button" onClick={onDownload}
              className="mt-4 inline-flex items-center justify-center rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90">
              ⬇ Download
            </button>
            <button type="button" onClick={onReset}
              className="mt-3 ml-2 inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent">
              Reset
            </button>
          </div>
        ) : (
          <div>
            <p className="text-sm font-medium text-emerald-600">Downloaded. Memory cleared.</p>
            <button type="button" onClick={onReset}
              className="mt-3 inline-flex items-center justify-center rounded-md border border-input bg-background px-3 py-1.5 text-xs font-medium hover:bg-accent">
              Convert another file
            </button>
          </div>
        )}
      </div>
      {error && (
        <div className="mt-4 rounded-md border border-destructive/30 bg-destructive/5 p-3 text-xs text-destructive">
          <p className="font-medium">Conversion failed</p>
          <p className="mt-1">{error}</p>
          <button onClick={onReset} className="mt-2 underline">Try again</button>
        </div>
      )}
      {children && <div className="mt-5">{children}</div>}
    </div>
  );
}

function Stat({ label, value, accent }: { label: string; value: number | string; accent?: boolean }) {
  return (
    <div className={`rounded-lg border p-3 ${accent ? "border-primary/30 bg-primary/5" : "border-border bg-muted/30"}`}>
      <div className="text-xs text-muted-foreground">{label}</div>
      <div className="mt-1 text-xl font-semibold tabular-nums">
        {typeof value === "number" ? (value ?? 0).toLocaleString() : value}
      </div>
    </div>
  );
}
