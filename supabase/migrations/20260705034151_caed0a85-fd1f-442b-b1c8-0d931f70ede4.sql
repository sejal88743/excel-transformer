
-- Master tables
CREATE TABLE public.items (
  name text PRIMARY KEY,
  hsn text,
  gst_rate numeric,
  mrp numeric,
  purchase_price numeric,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.items TO anon, authenticated;
GRANT ALL ON public.items TO service_role;
ALTER TABLE public.items ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read items"    ON public.items FOR SELECT USING (true);
CREATE POLICY "public insert items"  ON public.items FOR INSERT WITH CHECK (true);
CREATE POLICY "public update items"  ON public.items FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public delete items"  ON public.items FOR DELETE USING (true);

CREATE TABLE public.parties (
  name text PRIMARY KEY,
  gstin text,
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.parties TO anon, authenticated;
GRANT ALL ON public.parties TO service_role;
ALTER TABLE public.parties ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read parties"   ON public.parties FOR SELECT USING (true);
CREATE POLICY "public insert parties" ON public.parties FOR INSERT WITH CHECK (true);
CREATE POLICY "public update parties" ON public.parties FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public delete parties" ON public.parties FOR DELETE USING (true);

CREATE TABLE public.salespersons (
  name text PRIMARY KEY,
  created_at timestamptz NOT NULL DEFAULT now()
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.salespersons TO anon, authenticated;
GRANT ALL ON public.salespersons TO service_role;
ALTER TABLE public.salespersons ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read sp"   ON public.salespersons FOR SELECT USING (true);
CREATE POLICY "public insert sp" ON public.salespersons FOR INSERT WITH CHECK (true);
CREATE POLICY "public update sp" ON public.salespersons FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public delete sp" ON public.salespersons FOR DELETE USING (true);

-- Sale lines (positive/negative both stored as-is)
CREATE TABLE public.sale_lines (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  bill_no text NOT NULL,
  bill_date text NOT NULL,
  item_name text NOT NULL,
  party_name text,
  salesperson text,
  qty numeric,
  rate numeric,
  mrp numeric,
  gst_pct numeric,
  amount numeric,
  data jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  CONSTRAINT sale_lines_uniq UNIQUE (bill_no, item_name, bill_date)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.sale_lines TO anon, authenticated;
GRANT ALL ON public.sale_lines TO service_role;
ALTER TABLE public.sale_lines ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read sale"   ON public.sale_lines FOR SELECT USING (true);
CREATE POLICY "public insert sale" ON public.sale_lines FOR INSERT WITH CHECK (true);
CREATE POLICY "public update sale" ON public.sale_lines FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public delete sale" ON public.sale_lines FOR DELETE USING (true);

CREATE TABLE public.sale_return_lines (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  bill_no text NOT NULL,
  bill_date text NOT NULL,
  item_name text NOT NULL,
  party_name text,
  salesperson text,
  qty numeric,
  rate numeric,
  mrp numeric,
  gst_pct numeric,
  amount numeric,
  data jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  CONSTRAINT sale_return_lines_uniq UNIQUE (bill_no, item_name, bill_date)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.sale_return_lines TO anon, authenticated;
GRANT ALL ON public.sale_return_lines TO service_role;
ALTER TABLE public.sale_return_lines ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read sr"   ON public.sale_return_lines FOR SELECT USING (true);
CREATE POLICY "public insert sr" ON public.sale_return_lines FOR INSERT WITH CHECK (true);
CREATE POLICY "public update sr" ON public.sale_return_lines FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public delete sr" ON public.sale_return_lines FOR DELETE USING (true);

CREATE TABLE public.purchase_lines (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  voucher_no text NOT NULL,
  invoice_no text NOT NULL,
  invoice_date text NOT NULL,
  supplier_name text,
  item_name text NOT NULL,
  qty numeric,
  rate numeric,
  mrp numeric,
  gst_pct numeric,
  data jsonb NOT NULL,
  created_at timestamptz NOT NULL DEFAULT now(),
  CONSTRAINT purchase_lines_uniq UNIQUE (invoice_no, item_name, invoice_date)
);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.purchase_lines TO anon, authenticated;
GRANT ALL ON public.purchase_lines TO service_role;
ALTER TABLE public.purchase_lines ENABLE ROW LEVEL SECURITY;
CREATE POLICY "public read pur"   ON public.purchase_lines FOR SELECT USING (true);
CREATE POLICY "public insert pur" ON public.purchase_lines FOR INSERT WITH CHECK (true);
CREATE POLICY "public update pur" ON public.purchase_lines FOR UPDATE USING (true) WITH CHECK (true);
CREATE POLICY "public delete pur" ON public.purchase_lines FOR DELETE USING (true);

CREATE INDEX idx_sale_lines_bill ON public.sale_lines(bill_no);
CREATE INDEX idx_sr_lines_bill ON public.sale_return_lines(bill_no);
CREATE INDEX idx_pur_lines_inv ON public.purchase_lines(invoice_no);
