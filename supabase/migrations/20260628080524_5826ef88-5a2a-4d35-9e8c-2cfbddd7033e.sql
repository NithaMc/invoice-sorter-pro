
CREATE POLICY "Public read shared-files" ON storage.objects FOR SELECT USING (bucket_id = 'shared-files');
CREATE POLICY "Public upload shared-files" ON storage.objects FOR INSERT WITH CHECK (bucket_id = 'shared-files');
