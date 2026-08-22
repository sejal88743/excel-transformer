// In-memory dataset store. Wiped on page refresh. Used by /view page.
export type Row = Record<string, unknown>;

type Listener = () => void;
const listeners = new Set<Listener>();

interface Dataset {
  rows: Row[];
  fileName: string;
  uploadedAt: number;
}

const state = {
  sale: null as Dataset | null,
  return: null as Dataset | null,
  purchase: null as Dataset | null,
  register: null as Dataset | null,
  version: 0,
};

function emit() { state.version++; listeners.forEach((l) => l()); }

export const datasetStore = {
  getSale: () => state.sale,
  getReturn: () => state.return,
  getPurchase: () => state.purchase,
  getRegister: () => state.register,
  getVersion: () => state.version,
  setSale(rows: Row[], fileName: string) {
    state.sale = { rows: rows.map((r) => ({ ...r })), fileName, uploadedAt: Date.now() };
    emit();
  },
  setReturn(rows: Row[], fileName: string) {
    state.return = { rows: rows.map((r) => ({ ...r })), fileName, uploadedAt: Date.now() };
    emit();
  },
  setPurchase(rows: Row[], fileName: string) {
    state.purchase = { rows: rows.map((r) => ({ ...r })), fileName, uploadedAt: Date.now() };
    emit();
  },
  setRegister(rows: Row[], fileName: string) {
    state.register = { rows: rows.map((r) => ({ ...r })), fileName, uploadedAt: Date.now() };
    emit();
  },
  clearSale() { state.sale = null; emit(); },
  clearReturn() { state.return = null; emit(); },
  clearPurchase() { state.purchase = null; emit(); },
  clearRegister() { state.register = null; emit(); },
  subscribe(l: Listener) { listeners.add(l); return () => listeners.delete(l); },
};
