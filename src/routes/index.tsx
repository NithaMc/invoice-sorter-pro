import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Html5Qrcode } from "html5-qrcode";
import { Toaster, toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { supabase } from "@/integrations/supabase/client";
import { Wifi, Search, Trash2, Check, X } from "lucide-react";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Invoice Box Sorter" },
      { name: "description", content: "Scan barcodes to sort arrived invoice boxes against your shipment list. Synced live across every device." },
    ],
  }),
  component: Index,
});

type Invoice = {
  id: string;
  invoice: string;
  weight?: string | null;
  place?: string | null;
};

type ScanRecord = {
  id: string;
  invoice: string;
  status: "matched" | "duplicate" | "unknown" | "manual";
  weight?: string | null;
  place?: string | null;
  source: string;
  created_at: string;
};

function Index() {
  const [invoices, setInvoices] = useState<Invoice[]>([]);
  const [scans, setScans] = useState<ScanRecord[]>([]);
  const [manualInput, setManualInput] = useState("");
  const [scanning, setScanning] = useState(false);
  const [lastResult, setLastResult] = useState<ScanRecord | null>(null);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState<"all" | "received" | "pending" | "duplicate" | "unknown" | "verify">("all");
  const [search, setSearch] = useState("");
  const [verified, setVerified] = useState<Set<string>>(() => {
    if (typeof window === "undefined") return new Set();
    try { return new Set(JSON.parse(localStorage.getItem("verifiedScans") || "[]")); } catch { return new Set(); }
  });
  const saveVerified = (next: Set<string>) => {
    setVerified(new Set(next));
    if (typeof window !== "undefined") localStorage.setItem("verifiedScans", JSON.stringify([...next]));
  };
  const scannerRef = useRef<Html5Qrcode | null>(null);
  const lastScanRef = useRef<{ code: string; at: number }>({ code: "", at: 0 });
  const invoicesRef = useRef<Invoice[]>([]);
  const scansRef = useRef<ScanRecord[]>([]);

  useEffect(() => { invoicesRef.current = invoices; }, [invoices]);
  useEffect(() => { scansRef.current = scans; }, [scans]);

  // Initial load + realtime
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const [inv, sc] = await Promise.all([
        supabase.from("sorter_invoices").select("*").order("created_at", { ascending: true }),
        supabase.from("sorter_scans").select("*").order("created_at", { ascending: false }).limit(500),
      ]);
      if (cancelled) return;
      if (inv.data) setInvoices(inv.data as Invoice[]);
      if (sc.data) setScans(sc.data as ScanRecord[]);
      setLoading(false);
    })();

    const ch = supabase
      .channel("sorter_sync")
      .on("postgres_changes", { event: "INSERT", schema: "public", table: "sorter_invoices" },
        (p) => setInvoices((prev) => prev.some(i => i.id === (p.new as Invoice).id) ? prev : [...prev, p.new as Invoice]))
      .on("postgres_changes", { event: "DELETE", schema: "public", table: "sorter_invoices" },
        (p) => setInvoices((prev) => prev.filter(i => i.id !== (p.old as Invoice).id)))
      .on("postgres_changes", { event: "INSERT", schema: "public", table: "sorter_scans" },
        (p) => {
          const rec = p.new as ScanRecord;
          setScans((prev) => prev.some(s => s.id === rec.id) ? prev : [rec, ...prev]);
          setLastResult(rec);
        })
      .on("postgres_changes", { event: "DELETE", schema: "public", table: "sorter_scans" },
        (p) => setScans((prev) => prev.filter(s => s.id !== (p.old as ScanRecord).id)))
      .subscribe();

    return () => { cancelled = true; supabase.removeChannel(ch); };
  }, []);

  const handleExcel = async (file: File) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
    const parsed = rows.map((r) => {
      const keys = Object.keys(r);
      const find = (...names: string[]) =>
        keys.find((k) => names.some((n) => k.toLowerCase().replace(/[\s_]/g, "").includes(n)));
      const invKey = find("invoice", "inv", "bill", "awb", "docket") || keys[0];
      const wKey = find("weight", "kg", "wt");
      const pKey = find("place", "destination", "delivery", "city", "location", "address");
      return {
        invoice: String(r[invKey] ?? "").trim(),
        weight: wKey ? String(r[wKey] ?? "").trim() : null,
        place: pKey ? String(r[pKey] ?? "").trim() : null,
      };
    }).filter((r) => r.invoice);

    if (!parsed.length) { toast.error("No invoices found in file"); return; }

    // Replace shared list
    await supabase.from("sorter_scans").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    await supabase.from("sorter_invoices").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    const { error } = await supabase.from("sorter_invoices").insert(parsed);
    if (error) { toast.error("Upload failed: " + error.message); return; }
    toast.success(`Loaded ${parsed.length} invoices for everyone`);
  };

  const processCode = async (raw: string, source: "scan" | "manual") => {
    const code = raw.trim();
    if (!code) return;
    const now = Date.now();
    if (source === "scan" && lastScanRef.current.code === code && now - lastScanRef.current.at < 1500) return;
    lastScanRef.current = { code, at: now };

    const match = invoicesRef.current.find((i) => i.invoice.toLowerCase() === code.toLowerCase());
    const already = scansRef.current.find((s) => s.invoice.toLowerCase() === code.toLowerCase() && (s.status === "matched" || s.status === "manual"));

    let status: ScanRecord["status"];
    if (already) {
      status = "duplicate";
      toast.error(`Duplicate: ${code}`);
      if (navigator.vibrate) navigator.vibrate([100, 60, 100, 60, 200]);
    } else if (match) {
      status = source === "manual" ? "manual" : "matched";
      toast.success(`Matched: ${code}`);
      if (navigator.vibrate) navigator.vibrate(80);
    } else {
      status = "unknown";
      toast.warning(`Not in list: ${code}`);
      if (navigator.vibrate) navigator.vibrate([200, 80, 200]);
    }

    const { error } = await supabase.from("sorter_scans").insert({
      invoice: code,
      status,
      weight: match?.weight ?? null,
      place: match?.place ?? null,
      source,
    });
    if (error) toast.error(error.message);
  };

  const startScan = async () => {
    if (scanning) return;
    setScanning(true);
    try {
      const el = document.getElementById("reader");
      if (!el) throw new Error("reader element missing");
      const scanner = new Html5Qrcode("reader");
      scannerRef.current = scanner;
      await scanner.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: { width: 280, height: 160 } },
        (decoded) => { void processCode(decoded, "scan"); },
        () => {},
      );
    } catch (e) {
      toast.error("Camera failed: " + (e as Error).message);
      setScanning(false);
    }
  };

  const stopScan = async () => {
    try { await scannerRef.current?.stop(); await scannerRef.current?.clear(); } catch {}
    scannerRef.current = null;
    setScanning(false);
  };

  useEffect(() => () => { void stopScan(); }, []);

  const exportResults = () => {
    const rows = invoices.map((inv) => {
      const scan = scans.find((s) => s.invoice.toLowerCase() === inv.invoice.toLowerCase() && s.status !== "duplicate");
      return {
        Invoice: inv.invoice,
        Weight: inv.weight ?? "",
        Place: inv.place ?? "",
        Status: scan ? "Received" : "Pending",
        ScannedAt: scan ? new Date(scan.created_at).toLocaleString() : "",
        Method: scan?.status === "manual" ? "Manual" : scan ? "Scan" : "",
      };
    });
    const extras = scans.filter((s) => s.status === "unknown").map((s) => ({
      Invoice: s.invoice, Weight: "", Place: "", Status: "Unknown",
      ScannedAt: new Date(s.created_at).toLocaleString(), Method: "Scan",
    }));
    const ws = XLSX.utils.json_to_sheet([...rows, ...extras]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sorting");
    XLSX.writeFile(wb, `sorting-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const resetAll = async () => {
    if (!confirm("Reset shipment list and scans for EVERYONE?")) return;
    await supabase.from("sorter_scans").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    await supabase.from("sorter_invoices").delete().neq("id", "00000000-0000-0000-0000-000000000000");
    setLastResult(null);
    toast.success("Reset");
  };

  const receivedCount = invoices.filter((inv) =>
    scans.some((s) => s.invoice.toLowerCase() === inv.invoice.toLowerCase() && s.status !== "duplicate" && s.status !== "unknown"),
  ).length;
  const pendingCount = invoices.length - receivedCount;
  const duplicates = scans.filter((s) => s.status === "duplicate").length;
  const unknowns = scans.filter((s) => s.status === "unknown").length;

  const deleteScan = async (id: string) => {
    const { error } = await supabase.from("sorter_scans").delete().eq("id", id);
    if (error) toast.error(error.message); else toast.success("Deleted");
  };

  const toggleVerify = (id: string) => {
    const next = new Set(verified);
    if (next.has(id)) next.delete(id); else next.add(id);
    saveVerified(next);
  };

  const q = search.trim().toLowerCase();
  const tabRows = useMemo(() => {
    type Row = { key: string; invoice: string; weight?: string | null; place?: string | null; status: string; scanId?: string; verified?: boolean };
    let rows: Row[] = [];
    const receivedScans = scans.filter((s) => s.status === "matched" || s.status === "manual");
    if (tab === "all") {
      rows = invoices.map((inv) => {
        const sc = receivedScans.find((s) => s.invoice.toLowerCase() === inv.invoice.toLowerCase());
        return { key: inv.id, invoice: inv.invoice, weight: inv.weight, place: inv.place, status: sc ? "received" : "pending", scanId: sc?.id, verified: sc ? verified.has(sc.id) : undefined };
      });
    } else if (tab === "received") {
      rows = invoices
        .map((inv) => {
          const sc = receivedScans.find((s) => s.invoice.toLowerCase() === inv.invoice.toLowerCase());
          return sc ? { key: inv.id, invoice: inv.invoice, weight: inv.weight, place: inv.place, status: "received", scanId: sc.id, verified: verified.has(sc.id) } : null;
        })
        .filter(Boolean) as Row[];
    } else if (tab === "pending") {
      rows = invoices
        .filter((inv) => !receivedScans.some((s) => s.invoice.toLowerCase() === inv.invoice.toLowerCase()))
        .map((inv) => ({ key: inv.id, invoice: inv.invoice, weight: inv.weight, place: inv.place, status: "pending" }));
    } else if (tab === "duplicate") {
      rows = scans.filter((s) => s.status === "duplicate").map((s) => ({ key: s.id, invoice: s.invoice, weight: s.weight, place: s.place, status: "duplicate", scanId: s.id }));
    } else if (tab === "unknown") {
      rows = scans.filter((s) => s.status === "unknown").map((s) => ({ key: s.id, invoice: s.invoice, weight: s.weight, place: s.place, status: "unknown", scanId: s.id }));
    } else if (tab === "verify") {
      rows = receivedScans.map((s) => ({ key: s.id, invoice: s.invoice, weight: s.weight, place: s.place, status: "received", scanId: s.id, verified: verified.has(s.id) }));
    }
    if (q) rows = rows.filter((r) => r.invoice.toLowerCase().includes(q) || (r.place || "").toLowerCase().includes(q));
    return rows;
  }, [tab, invoices, scans, verified, q]);


  return (
    <div className="min-h-screen bg-background pb-24">
      <Toaster position="top-center" richColors />
      <header className="sticky top-0 z-10 border-b bg-background/95 backdrop-blur px-4 py-3">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-lg font-semibold">Invoice Box Sorter</h1>
            <p className="text-xs text-muted-foreground">Shared live across all devices</p>
          </div>
          <span className="inline-flex items-center gap-1 text-xs text-green-600">
            <Wifi className="size-3" /> Live
          </span>
        </div>
      </header>

      <div className="px-4 py-4 space-y-4 max-w-xl mx-auto">
        {!loading && invoices.length === 0 && (
          <Card className="p-4 space-y-3">
            <h2 className="font-medium">1. Load shipment Excel (shared)</h2>
            <p className="text-xs text-muted-foreground">Columns: Invoice, Weight, Place. Everyone will see this list.</p>
            <Input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && handleExcel(e.target.files[0])} />
          </Card>
        )}

        {invoices.length > 0 && (
          <>
            <div className="grid grid-cols-5 gap-1.5">
              {([
                ["all", "All", invoices.length, ""],
                ["received", "Received", receivedCount, "text-green-600"],
                ["pending", "Pending", pendingCount, ""],
                ["duplicate", "Dupes", duplicates, "text-destructive"],
                ["unknown", "Unknown", unknowns, "text-amber-600"],
              ] as const).map(([k, label, count, cls]) => (
                <button
                  key={k}
                  onClick={() => setTab(k)}
                  className={`rounded-lg border p-2 text-center transition ${tab === k ? "border-primary bg-primary/10" : "bg-card hover:bg-accent"}`}
                >
                  <div className={`text-lg font-bold ${cls}`}>{count}</div>
                  <div className="text-[10px] uppercase text-muted-foreground">{label}</div>
                </button>
              ))}
            </div>

            <Card className="p-3 space-y-3">
              <div className="flex gap-2">
                {!scanning ? (
                  <Button onClick={startScan} className="flex-1">Start camera scan</Button>
                ) : (
                  <Button variant="destructive" onClick={stopScan} className="flex-1">Stop scan</Button>
                )}
              </div>
              <div id="reader" className={scanning ? "rounded-md overflow-hidden" : "hidden"} />
              <div className="flex gap-2">
                <Input
                  placeholder="Type invoice # if barcode unreadable"
                  value={manualInput}
                  onChange={(e) => setManualInput(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && manualInput.trim()) {
                      void processCode(manualInput, "manual");
                      setManualInput("");
                    }
                  }}
                />
                <Button
                  variant="secondary"
                  onClick={() => { if (manualInput.trim()) { void processCode(manualInput, "manual"); setManualInput(""); } }}
                >Mark</Button>
              </div>
            </Card>

            {lastResult && (
              <Card className={`p-4 border-2 ${
                lastResult.status === "matched" || lastResult.status === "manual" ? "border-green-500" :
                lastResult.status === "duplicate" ? "border-destructive" : "border-amber-500"
              }`}>
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <div className="text-xs text-muted-foreground">Last scan</div>
                    <div className="text-lg font-bold break-all">{lastResult.invoice}</div>
                  </div>
                  <Badge variant={
                    lastResult.status === "duplicate" ? "destructive" :
                    lastResult.status === "unknown" ? "outline" : "default"
                  }>{lastResult.status.toUpperCase()}</Badge>
                </div>
                {(lastResult.weight || lastResult.place) && (
                  <div className="mt-3 grid grid-cols-2 gap-2 text-sm">
                    {lastResult.weight && (<div><div className="text-xs text-muted-foreground">Weight</div><div className="font-medium">{lastResult.weight}</div></div>)}
                    {lastResult.place && (<div><div className="text-xs text-muted-foreground">Place</div><div className="font-medium">{lastResult.place}</div></div>)}
                  </div>
                )}
                {(lastResult.status === "matched" || lastResult.status === "manual") && (
                  <div className="mt-3 flex gap-2">
                    <Button
                      size="sm"
                      variant={verified.has(lastResult.id) ? "secondary" : "default"}
                      className="flex-1"
                      onClick={() => toggleVerify(lastResult.id)}
                    >
                      <Check className="size-4" /> {verified.has(lastResult.id) ? "Verified ✓" : "OK"}
                    </Button>
                    <Button size="sm" variant="destructive" className="flex-1" onClick={() => { void deleteScan(lastResult.id); setLastResult(null); }}>
                      <X className="size-4" /> Not OK
                    </Button>
                  </div>
                )}
              </Card>
            )}

            <Card className="p-3 space-y-3">
              <div className="flex items-center justify-between gap-2">
                <h3 className="font-medium text-sm capitalize">{tab} ({tabRows.length})</h3>
                <div className="flex gap-2">
                  <Button size="sm" variant={tab === "verify" ? "default" : "outline"} onClick={() => setTab("verify")}>Verify</Button>
                  <Button size="sm" variant="outline" onClick={exportResults}>Export</Button>
                  <Button size="sm" variant="ghost" onClick={resetAll}>Reset</Button>
                </div>
              </div>
              <div className="relative">
                <Search className="absolute left-2 top-1/2 -translate-y-1/2 size-4 text-muted-foreground" />
                <Input
                  className="pl-8"
                  placeholder="Search invoice # or place…"
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
              </div>
              <div className="max-h-96 overflow-y-auto divide-y">
                {tabRows.length === 0 && <p className="text-xs text-muted-foreground py-2">Nothing to show.</p>}
                {tabRows.map((r) => {
                  const canDelete = r.status === "duplicate" || r.status === "unknown" || tab === "verify";
                  const showVerify = tab === "verify" && r.scanId;
                  return (
                    <div key={r.key} className="py-2 flex items-center justify-between gap-2 text-sm">
                      <div className="min-w-0 flex-1">
                        <div className="font-medium truncate flex items-center gap-2">
                          {r.invoice}
                          {r.verified && <Check className="size-3 text-green-600 shrink-0" />}
                        </div>
                        <div className="text-xs text-muted-foreground truncate">
                          {r.place || "—"}{r.weight ? ` · ${r.weight}` : ""}
                        </div>
                      </div>
                      <div className="flex items-center gap-1 shrink-0">
                        <Badge variant={
                          r.status === "duplicate" ? "destructive" :
                          r.status === "unknown" ? "outline" :
                          r.status === "pending" ? "secondary" : "default"
                        }>{r.status}</Badge>
                        {showVerify && r.scanId && (
                          <Button size="icon" variant={r.verified ? "secondary" : "outline"} className="h-7 w-7" onClick={() => toggleVerify(r.scanId!)} title="Mark verified">
                            <Check className="size-3.5" />
                          </Button>
                        )}
                        {canDelete && r.scanId && (
                          <Button size="icon" variant="ghost" className="h-7 w-7" onClick={() => void deleteScan(r.scanId!)} title="Delete">
                            <Trash2 className="size-3.5 text-destructive" />
                          </Button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </Card>
          </>
        )}
      </div>
    </div>
  );
}
