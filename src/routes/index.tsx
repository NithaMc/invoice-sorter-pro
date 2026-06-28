import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { Html5Qrcode } from "html5-qrcode";
import { Toaster, toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Invoice Box Sorter" },
      { name: "description", content: "Scan barcodes to sort arrived invoice boxes against your shipment list." },
    ],
  }),
  component: Index,
});

type Invoice = {
  invoice: string;
  weight?: string;
  place?: string;
  [k: string]: string | undefined;
};

type ScanRecord = {
  invoice: string;
  status: "matched" | "duplicate" | "unknown" | "manual";
  at: number;
  weight?: string;
  place?: string;
};

const STORAGE_KEY = "invoice_sorter_v1";

function loadState(): { invoices: Invoice[]; scans: ScanRecord[] } {
  if (typeof window === "undefined") return { invoices: [], scans: [] };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  return { invoices: [], scans: [] };
}

function Index() {
  const [invoices, setInvoices] = useState<Invoice[]>([]);
  const [scans, setScans] = useState<ScanRecord[]>([]);
  const [manualInput, setManualInput] = useState("");
  const [scanning, setScanning] = useState(false);
  const [lastResult, setLastResult] = useState<ScanRecord | null>(null);
  const scannerRef = useRef<Html5Qrcode | null>(null);
  const lastScanRef = useRef<{ code: string; at: number }>({ code: "", at: 0 });

  useEffect(() => {
    const s = loadState();
    setInvoices(s.invoices);
    setScans(s.scans);
  }, []);

  useEffect(() => {
    if (typeof window === "undefined") return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ invoices, scans }));
  }, [invoices, scans]);

  const handleExcel = async (file: File) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { defval: "" });
    const parsed: Invoice[] = rows.map((r) => {
      const keys = Object.keys(r);
      const find = (...names: string[]) =>
        keys.find((k) => names.some((n) => k.toLowerCase().replace(/[\s_]/g, "").includes(n)));
      const invKey = find("invoice", "inv", "bill", "awb", "docket") || keys[0];
      const wKey = find("weight", "kg", "wt");
      const pKey = find("place", "destination", "delivery", "city", "location", "address");
      return {
        invoice: String(r[invKey] ?? "").trim(),
        weight: wKey ? String(r[wKey] ?? "").trim() : undefined,
        place: pKey ? String(r[pKey] ?? "").trim() : undefined,
      };
    }).filter((r) => r.invoice);
    setInvoices(parsed);
    setScans([]);
    toast.success(`Loaded ${parsed.length} invoices`);
  };

  const processCode = (raw: string, source: "scan" | "manual") => {
    const code = raw.trim();
    if (!code) return;
    const now = Date.now();
    if (source === "scan" && lastScanRef.current.code === code && now - lastScanRef.current.at < 1500) return;
    lastScanRef.current = { code, at: now };

    const match = invoices.find((i) => i.invoice.toLowerCase() === code.toLowerCase());
    const already = scans.find((s) => s.invoice.toLowerCase() === code.toLowerCase() && (s.status === "matched" || s.status === "manual"));

    let rec: ScanRecord;
    if (already) {
      rec = { invoice: code, status: "duplicate", at: now, weight: match?.weight, place: match?.place };
      toast.error(`Duplicate: ${code}`);
      if (navigator.vibrate) navigator.vibrate([100, 60, 100, 60, 200]);
    } else if (match) {
      rec = { invoice: code, status: source === "manual" ? "manual" : "matched", at: now, weight: match.weight, place: match.place };
      toast.success(`Matched: ${code}`);
      if (navigator.vibrate) navigator.vibrate(80);
    } else {
      rec = { invoice: code, status: "unknown", at: now };
      toast.warning(`Not in list: ${code}`);
      if (navigator.vibrate) navigator.vibrate([200, 80, 200]);
    }
    setLastResult(rec);
    setScans((prev) => [rec, ...prev]);
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
        (decoded) => processCode(decoded, "scan"),
        () => {},
      );
    } catch (e) {
      toast.error("Camera failed: " + (e as Error).message);
      setScanning(false);
    }
  };

  const stopScan = async () => {
    try {
      await scannerRef.current?.stop();
      await scannerRef.current?.clear();
    } catch {}
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
        ScannedAt: scan ? new Date(scan.at).toLocaleString() : "",
        Method: scan?.status === "manual" ? "Manual" : scan ? "Scan" : "",
      };
    });
    const extras = scans.filter((s) => s.status === "unknown").map((s) => ({
      Invoice: s.invoice, Weight: "", Place: "", Status: "Unknown", ScannedAt: new Date(s.at).toLocaleString(), Method: "Scan",
    }));
    const ws = XLSX.utils.json_to_sheet([...rows, ...extras]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sorting");
    XLSX.writeFile(wb, `sorting-${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const receivedCount = invoices.filter((inv) =>
    scans.some((s) => s.invoice.toLowerCase() === inv.invoice.toLowerCase() && s.status !== "duplicate" && s.status !== "unknown"),
  ).length;
  const pendingCount = invoices.length - receivedCount;
  const duplicates = scans.filter((s) => s.status === "duplicate").length;
  const unknowns = scans.filter((s) => s.status === "unknown").length;

  return (
    <div className="min-h-screen bg-background pb-24">
      <Toaster position="top-center" richColors />
      <header className="sticky top-0 z-10 border-b bg-background/95 backdrop-blur px-4 py-3">
        <h1 className="text-lg font-semibold">Invoice Box Sorter</h1>
        <p className="text-xs text-muted-foreground">Scan arrived boxes against your shipment list</p>
      </header>

      <div className="px-4 py-4 space-y-4 max-w-xl mx-auto">
        {invoices.length === 0 && (
          <Card className="p-4 space-y-3">
            <h2 className="font-medium">1. Load shipment Excel</h2>
            <p className="text-xs text-muted-foreground">Columns: Invoice, Weight, Place (any order, any case).</p>
            <Input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => e.target.files?.[0] && handleExcel(e.target.files[0])} />
          </Card>
        )}

        {invoices.length > 0 && (
          <>
            <div className="grid grid-cols-4 gap-2">
              <Card className="p-2 text-center">
                <div className="text-xl font-bold">{receivedCount}</div>
                <div className="text-[10px] uppercase text-muted-foreground">Received</div>
              </Card>
              <Card className="p-2 text-center">
                <div className="text-xl font-bold">{pendingCount}</div>
                <div className="text-[10px] uppercase text-muted-foreground">Pending</div>
              </Card>
              <Card className="p-2 text-center">
                <div className="text-xl font-bold text-destructive">{duplicates}</div>
                <div className="text-[10px] uppercase text-muted-foreground">Dupes</div>
              </Card>
              <Card className="p-2 text-center">
                <div className="text-xl font-bold text-amber-600">{unknowns}</div>
                <div className="text-[10px] uppercase text-muted-foreground">Unknown</div>
              </Card>
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
                      processCode(manualInput, "manual");
                      setManualInput("");
                    }
                  }}
                />
                <Button
                  variant="secondary"
                  onClick={() => { if (manualInput.trim()) { processCode(manualInput, "manual"); setManualInput(""); } }}
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
                    {lastResult.weight && (
                      <div>
                        <div className="text-xs text-muted-foreground">Weight</div>
                        <div className="font-medium">{lastResult.weight}</div>
                      </div>
                    )}
                    {lastResult.place && (
                      <div>
                        <div className="text-xs text-muted-foreground">Place</div>
                        <div className="font-medium">{lastResult.place}</div>
                      </div>
                    )}
                  </div>
                )}
              </Card>
            )}

            <Card className="p-3 space-y-2">
              <div className="flex items-center justify-between">
                <h3 className="font-medium text-sm">Scan history ({scans.length})</h3>
                <div className="flex gap-2">
                  <Button size="sm" variant="outline" onClick={exportResults}>Export</Button>
                  <Button size="sm" variant="ghost" onClick={() => { if (confirm("Reset everything?")) { setInvoices([]); setScans([]); setLastResult(null); } }}>Reset</Button>
                </div>
              </div>
              <div className="max-h-80 overflow-y-auto divide-y">
                {scans.length === 0 && <p className="text-xs text-muted-foreground py-2">No scans yet.</p>}
                {scans.map((s, i) => (
                  <div key={i} className="py-2 flex items-center justify-between gap-2 text-sm">
                    <div className="min-w-0">
                      <div className="font-medium truncate">{s.invoice}</div>
                      <div className="text-xs text-muted-foreground truncate">
                        {s.place || "—"}{s.weight ? ` · ${s.weight}` : ""}
                      </div>
                    </div>
                    <Badge variant={
                      s.status === "duplicate" ? "destructive" :
                      s.status === "unknown" ? "outline" : "secondary"
                    }>{s.status}</Badge>
                  </div>
                ))}
              </div>
            </Card>
          </>
        )}
      </div>
    </div>
  );
}
