import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useRef, useState } from "react";
import { supabase } from "@/integrations/supabase/client";
import { Toaster, toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import { Card } from "@/components/ui/card";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import { Trash2, Upload, FileIcon, Loader2, Wifi } from "lucide-react";

export const Route = createFileRoute("/share")({
  head: () => ({
    meta: [
      { title: "Live Share — Real-time text & file sharing" },
      { name: "description", content: "Share text and files instantly across devices. Live updates with no refresh." },
    ],
  }),
  component: SharePage,
});

type Item = {
  id: string;
  text: string | null;
  file_url: string | null;
  file_name: string | null;
  created_at: string;
};

function SharePage() {
  const [items, setItems] = useState<Item[]>([]);
  const [text, setText] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [tab, setTab] = useState("display");
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      const { data, error } = await supabase
        .from("shared_items")
        .select("*")
        .order("created_at", { ascending: false })
        .limit(200);
      if (!cancelled && !error && data) setItems(data as Item[]);
    })();

    const channel = supabase
      .channel("shared_items_changes")
      .on(
        "postgres_changes",
        { event: "INSERT", schema: "public", table: "shared_items" },
        (payload) => setItems((prev) => [payload.new as Item, ...prev]),
      )
      .on(
        "postgres_changes",
        { event: "DELETE", schema: "public", table: "shared_items" },
        (payload) => setItems((prev) => prev.filter((i) => i.id !== (payload.old as Item).id)),
      )
      .subscribe();

    return () => {
      cancelled = true;
      supabase.removeChannel(channel);
    };
  }, []);

  const submit = async () => {
    if (!text.trim() && !file) {
      toast.error("Add text or pick a file");
      return;
    }
    setSubmitting(true);
    try {
      let file_url: string | null = null;
      let file_name: string | null = null;
      if (file) {
        const path = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}-${file.name}`;
        const up = await supabase.storage.from("shared-files").upload(path, file, {
          cacheControl: "3600",
          upsert: false,
        });
        if (up.error) throw up.error;
        const signed = await supabase.storage
          .from("shared-files")
          .createSignedUrl(path, 60 * 60 * 24 * 365 * 5);
        if (signed.error) throw signed.error;
        file_url = signed.data.signedUrl;
        file_name = file.name;
      }
      const { error } = await supabase.from("shared_items").insert({
        text: text.trim() || null,
        file_url,
        file_name,
      });
      if (error) throw error;
      setText("");
      setFile(null);
      if (fileInputRef.current) fileInputRef.current.value = "";
      toast.success("Shared!");
      setTab("display");
    } catch (e) {
      toast.error("Failed: " + (e as Error).message);
    } finally {
      setSubmitting(false);
    }
  };

  const remove = async (id: string) => {
    const { error } = await supabase.from("shared_items").delete().eq("id", id);
    if (error) toast.error(error.message);
  };

  const isImage = (name?: string | null) =>
    !!name && /\.(png|jpe?g|gif|webp|avif|svg|bmp)$/i.test(name);

  return (
    <div className="min-h-screen bg-background pb-16">
      <Toaster position="top-center" richColors />
      <header className="sticky top-0 z-10 border-b bg-background/95 backdrop-blur px-4 py-3">
        <div className="max-w-2xl mx-auto flex items-center justify-between">
          <div>
            <h1 className="text-lg font-semibold">Live Share</h1>
            <p className="text-xs text-muted-foreground">Real-time text & file sharing across devices</p>
          </div>
          <span className="inline-flex items-center gap-1 text-xs text-green-600">
            <Wifi className="size-3" /> Live
          </span>
        </div>
      </header>

      <main className="max-w-2xl mx-auto px-4 py-4">
        <Tabs value={tab} onValueChange={setTab}>
          <TabsList className="grid grid-cols-2 w-full">
            <TabsTrigger value="upload">Upload</TabsTrigger>
            <TabsTrigger value="display">Display ({items.length})</TabsTrigger>
          </TabsList>

          <TabsContent value="upload" className="mt-4">
            <Card className="p-4 space-y-3">
              <Textarea
                placeholder="Type a message to share..."
                value={text}
                onChange={(e) => setText(e.target.value)}
                rows={5}
              />
              <div className="flex items-center gap-2">
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  onChange={(e) => setFile(e.target.files?.[0] ?? null)}
                />
                <Button
                  type="button"
                  variant="outline"
                  onClick={() => fileInputRef.current?.click()}
                  disabled={submitting}
                >
                  <Upload /> {file ? "Change file" : "Attach file"}
                </Button>
                {file && (
                  <span className="text-xs text-muted-foreground truncate">{file.name}</span>
                )}
              </div>
              <Button onClick={submit} disabled={submitting} className="w-full">
                {submitting ? <Loader2 className="animate-spin" /> : null}
                {submitting ? "Sharing..." : "Share"}
              </Button>
              <p className="text-[11px] text-muted-foreground">
                Public feed — anyone with the link sees everything you post.
              </p>
            </Card>
          </TabsContent>

          <TabsContent value="display" className="mt-4 space-y-3">
            {items.length === 0 && (
              <Card className="p-8 text-center text-sm text-muted-foreground">
                Nothing shared yet. Switch to Upload to post something.
              </Card>
            )}
            {items.map((item) => (
              <Card key={item.id} className="p-4 space-y-2">
                <div className="flex items-start justify-between gap-2">
                  <span className="text-[11px] text-muted-foreground">
                    {new Date(item.created_at).toLocaleString()}
                  </span>
                  <Button size="icon" variant="ghost" onClick={() => remove(item.id)} aria-label="Delete">
                    <Trash2 className="size-4" />
                  </Button>
                </div>
                {item.text && (
                  <p className="whitespace-pre-wrap break-words text-sm">{item.text}</p>
                )}
                {item.file_url && (
                  <div className="pt-1">
                    {isImage(item.file_name) ? (
                      <a href={item.file_url} target="_blank" rel="noreferrer">
                        <img
                          src={item.file_url}
                          alt={item.file_name ?? "shared image"}
                          className="rounded-md max-h-80 border"
                        />
                      </a>
                    ) : (
                      <a
                        href={item.file_url}
                        target="_blank"
                        rel="noreferrer"
                        download={item.file_name ?? undefined}
                        className="inline-flex items-center gap-2 text-sm text-primary hover:underline"
                      >
                        <FileIcon className="size-4" /> {item.file_name ?? "Download file"}
                      </a>
                    )}
                  </div>
                )}
              </Card>
            ))}
          </TabsContent>
        </Tabs>
      </main>
    </div>
  );
}
