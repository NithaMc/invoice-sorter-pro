
CREATE TABLE public.shared_items (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  text TEXT,
  file_url TEXT,
  file_name TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

GRANT SELECT, INSERT, DELETE ON public.shared_items TO anon, authenticated;
GRANT ALL ON public.shared_items TO service_role;

ALTER TABLE public.shared_items ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Anyone can view shared items" ON public.shared_items FOR SELECT USING (true);
CREATE POLICY "Anyone can insert shared items" ON public.shared_items FOR INSERT WITH CHECK (true);
CREATE POLICY "Anyone can delete shared items" ON public.shared_items FOR DELETE USING (true);

ALTER PUBLICATION supabase_realtime ADD TABLE public.shared_items;
ALTER TABLE public.shared_items REPLICA IDENTITY FULL;
