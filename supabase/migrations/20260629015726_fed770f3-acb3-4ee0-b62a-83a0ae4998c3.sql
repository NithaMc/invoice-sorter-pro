
CREATE TABLE public.sorter_invoices (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  invoice text NOT NULL,
  weight text,
  place text,
  created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX sorter_invoices_invoice_lower_idx ON public.sorter_invoices (lower(invoice));
GRANT SELECT, INSERT, UPDATE, DELETE ON public.sorter_invoices TO anon, authenticated;
GRANT ALL ON public.sorter_invoices TO service_role;
ALTER TABLE public.sorter_invoices ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Anyone view sorter_invoices" ON public.sorter_invoices FOR SELECT USING (true);
CREATE POLICY "Anyone insert sorter_invoices" ON public.sorter_invoices FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone delete sorter_invoices" ON public.sorter_invoices FOR DELETE USING (true);

CREATE TABLE public.sorter_scans (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  invoice text NOT NULL,
  status text NOT NULL,
  weight text,
  place text,
  source text NOT NULL DEFAULT 'scan',
  created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX sorter_scans_created_idx ON public.sorter_scans (created_at DESC);
GRANT SELECT, INSERT, UPDATE, DELETE ON public.sorter_scans TO anon, authenticated;
GRANT ALL ON public.sorter_scans TO service_role;
ALTER TABLE public.sorter_scans ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Anyone view sorter_scans" ON public.sorter_scans FOR SELECT USING (true);
CREATE POLICY "Anyone insert sorter_scans" ON public.sorter_scans FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone delete sorter_scans" ON public.sorter_scans FOR DELETE USING (true);

ALTER PUBLICATION supabase_realtime ADD TABLE public.sorter_invoices;
ALTER PUBLICATION supabase_realtime ADD TABLE public.sorter_scans;
