import { useRef, useState } from "react";

// Dev: Vite proxies /api -> http://127.0.0.1:5000.
// Static build served elsewhere: set VITE_API_BASE=http://localhost:5000
const API_BASE = import.meta.env.VITE_API_BASE || "";

const fmt = (v) => (v ? Number(v).toLocaleString("id-ID") : "—");
const pct = (v) => `${(Number(v) * 100).toFixed(2)}%`;

function tone(v) {
  if (v >= 0.99) return "#2E6B4A";
  if (v >= 0.9) return "#8A6516";
  return "#A33A2A";
}

function saveFile(r) {
  const bin = atob(r.data_b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = r.out_name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

export default function App() {
  const [files, setFiles] = useState([]);
  const [tol, setTol] = useState(7);
  const [busy, setBusy] = useState(false);
  const [results, setResults] = useState([]);
  const [error, setError] = useState("");
  const inputRef = useRef(null);

  const addFiles = (list) => {
    const next = Array.from(list).filter((f) => /\.xlsm?x?$/i.test(f.name));
    if (next.length) setFiles((prev) => [...prev, ...next]);
  };

  const submit = async () => {
    if (!files.length) return;
    setBusy(true);
    setError("");
    setResults([]);
    try {
      const fd = new FormData();
      files.forEach((f) => fd.append("files", f));
      fd.append("tol", String(tol));
      const res = await fetch(`${API_BASE}/api/process`, { method: "POST", body: fd });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const json = await res.json();
      setResults(json.results || []);
      const errs = (json.results || []).filter((r) => r.error);
      if (errs.length) setError(errs.map((e) => `${e.name}: ${e.error}`).join("  ·  "));
    } catch (e) {
      setError(`Tidak bisa menghubungi backend (${e.message}). Pastikan Flask jalan di :5000.`);
    } finally {
      setBusy(false);
    }
  };

  const clearAll = () => {
    setFiles([]);
    setResults([]);
    setError("");
  };

  const ok = results.filter((r) => !r.error);

  return (
    <div className="min-h-full">
      <Masthead />

      <main className="mx-auto max-w-6xl px-6 pb-24">
        <section aria-labelledby="upload-h" className="pt-10">
          <SectionLabel id="upload-h">01 Berkas</SectionLabel>

          <div
            role="button"
            tabIndex={0}
            onClick={() => inputRef.current?.click()}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                inputRef.current?.click();
              }
            }}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => {
              e.preventDefault();
              addFiles(e.dataTransfer.files);
            }}
            className="mt-4 w-full border border-dashed border-rule-2 bg-paper-2/40 px-6 py-10 text-left transition-colors duration-200 ease-out hover:border-accent"
          >
            <p className="font-display text-3xl leading-snug">
              Seret berkas <span className="italic">.xlsx</span>, atau klik untuk memilih
            </p>
            <p className="mt-2 text-sm text-muted">
              Bisa banyak sekaligus: heonz, muncul, ultra sakti, us.
            </p>
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xlsm"
              multiple
              hidden
              onChange={(e) => {
                addFiles(e.target.files);
                e.target.value = "";
              }}
            />
          </div>

          {files.length > 0 && (
            <ul className="mt-4 flex flex-wrap gap-x-5 gap-y-2 font-mono text-xs text-ink-2">
              {files.map((f, i) => (
                <li key={i} className="border-b border-rule pb-0.5">
                  {f.name}
                </li>
              ))}
            </ul>
          )}

          <div className="mt-6 flex flex-wrap items-center gap-x-6 gap-y-4">
            <label className="flex items-baseline gap-2">
              <span className="font-mono text-[11px] tracking-[0.16em] text-muted">Toleransi</span>
              <input
                type="number"
                min={0}
                max={60}
                value={tol}
                onChange={(e) =>
                  setTol(Math.max(0, Math.min(60, parseInt(e.target.value || "7", 10) || 7)))
                }
                className="tnum w-14 border-b border-rule-2 bg-transparent px-1 py-0.5 text-center font-mono text-ink outline-none focus:border-accent"
              />
              <span className="text-sm text-muted">hari setelah planned receipt</span>
            </label>

            <button
              type="button"
              onClick={submit}
              disabled={busy || !files.length}
              className="min-h-[44px] bg-ink px-6 text-sm font-medium text-paper transition-colors duration-200 ease-out hover:bg-accent disabled:cursor-not-allowed disabled:opacity-40"
            >
              {busy ? "Menghitung…" : "Proses"}
            </button>
            <button
              type="button"
              onClick={clearAll}
              disabled={busy || (!files.length && !results.length)}
              className="min-h-[44px] border border-rule px-6 text-sm text-ink-2 transition-colors duration-200 ease-out hover:border-ink disabled:cursor-not-allowed disabled:opacity-40"
            >
              Bersihkan
            </button>
            {ok.length > 0 && (
              <button
                type="button"
                onClick={() => ok.forEach((r, i) => setTimeout(() => saveFile(r), i * 400))}
                className="min-h-[44px] border-b border-ink pb-0.5 text-sm text-ink transition-colors duration-200 ease-out hover:border-accent hover:text-accent"
              >
                Unduh semua ({ok.length})
              </button>
            )}
          </div>

          {error && (
            <p className="mt-4 border-l-2 border-accent bg-accent/5 px-4 py-3 text-sm text-accent">
              {error}
            </p>
          )}
        </section>

        {results.length > 0 && (
          <section aria-labelledby="result-h" className="pt-14">
            <SectionLabel id="result-h">02 Rekap</SectionLabel>
            <div className="mt-4 space-y-8">
              {results.map((r, i) => (
                <ResultSheet key={i} r={r} index={i} />
              ))}
            </div>
          </section>
        )}

        <Method />
      </main>
    </div>
  );
}

function Masthead() {
  return (
    <header className="border-b border-rule">
      <div className="mx-auto max-w-6xl px-6 pt-14 pb-9">
        <h1 className="font-display text-[3.25rem] font-medium leading-none tracking-tight sm:text-[4.5rem]">
          SL &amp; OTIF
        </h1>
        <p className="mt-5 max-w-2xl text-[15px] leading-relaxed text-ink-2">
          Rekap penerimaan barang: <em>Late</em> / <em>On Time</em>, FULFIL, dan OTIF, dari ekspor{" "}
          <span className="font-mono text-[13px] text-ink">PURCHASE ORDER LINES VS ARRIVAL</span>.
          Berkas asli tidak disentuh, header dan formatnya tetap, dan rumusnya ikut tertulis di
          dalam berkas.
        </p>
      </div>
    </header>
  );
}

function SectionLabel({ id, children }) {
  return (
    <h2
      id={id}
      className="flex items-center gap-4 font-mono text-[11px] tracking-[0.16em] text-muted"
    >
      <span>{children}</span>
      <span aria-hidden="true" className="h-px flex-1 bg-rule" />
    </h2>
  );
}

function ResultSheet({ r, index }) {
  if (r.error) {
    return (
      <p className="border-l-2 border-accent bg-accent/5 px-4 py-3 text-sm text-accent">
        <span className="font-mono">{r.name}</span>: {r.error}
      </p>
    );
  }
  const s = r.stats;
  const months = s.order || [];
  const agg = r.agg || {};
  const g = s.grand || {};
  const F = g.lateL + g.ontL;
  const G = g.lateS + g.ontS;

  return (
    <article
      className="animate-rise border border-rule bg-sheet"
      style={{ animationDelay: `${index * 60}ms` }}
    >
      <header className="flex flex-wrap items-baseline justify-between gap-x-6 gap-y-3 border-b border-rule px-6 py-4">
        <div>
          <h3 className="font-display text-3xl font-medium leading-tight">{s.title}</h3>
          <p className="mt-1 font-mono text-xs text-muted">
            {s.year} · {Number(s.nlines).toLocaleString("id-ID")} baris · {months.join(" / ") || "—"}
          </p>
          <p className="mt-1 font-mono text-[11px] text-muted">{r.out_name}</p>
        </div>
        <button
          type="button"
          onClick={() => saveFile(r)}
          className="min-h-[44px] bg-ink px-6 text-sm font-medium text-paper transition-colors duration-200 ease-out hover:bg-accent"
        >
          Unduh .xlsx
        </button>
      </header>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[760px] border-collapse text-[13px]">
          <thead>
            <tr className="border-b border-rule-2">
              <th scope="col" className="px-6 py-3 text-left font-mono text-[10px] uppercase tracking-[0.14em] text-muted">Bulan</th>
              {["Late · PO", "Late · Datang", "On Time · PO", "On Time · Datang"].map((h) => (
                <th key={h} scope="col" className="px-3 py-3 text-right font-mono text-[10px] uppercase tracking-[0.14em] text-muted">{h}</th>
              ))}
              <th scope="col" className="px-3 py-3 text-right font-mono text-[10px] uppercase tracking-[0.14em] text-muted">Total PO</th>
              <th scope="col" className="px-3 py-3 text-right font-mono text-[10px] uppercase tracking-[0.14em] text-muted">Total Datang</th>
              <th scope="col" className="px-3 py-3 text-right font-mono text-[10px] uppercase tracking-[0.14em] text-muted">FULFIL</th>
              <th scope="col" className="px-6 py-3 text-right font-mono text-[10px] uppercase tracking-[0.14em] text-muted">OTIF</th>
            </tr>
          </thead>
          <tbody>
            {months.map((m) => {
              const d = agg[m];
              const f = d.lateL + d.ontL;
              const gg = d.lateS + d.ontS;
              return (
                <tr key={m} className="border-b border-rule">
                  <th scope="row" className="px-6 py-2.5 text-left font-display text-xl font-normal">{m}</th>
                  <td className="tnum px-3 py-2.5 text-right font-mono text-ink-2">{fmt(d.lateL)}</td>
                  <td className="tnum px-3 py-2.5 text-right font-mono text-ink-2">{fmt(d.lateS)}</td>
                  <td className="tnum px-3 py-2.5 text-right font-mono text-ink-2">{fmt(d.ontL)}</td>
                  <td className="tnum px-3 py-2.5 text-right font-mono text-ink-2">{fmt(d.ontS)}</td>
                  <td className="tnum px-3 py-2.5 text-right font-mono text-ink">{f.toLocaleString("id-ID")}</td>
                  <td className="tnum px-3 py-2.5 text-right font-mono text-ink">{gg.toLocaleString("id-ID")}</td>
                  <td className="px-3 py-2.5 text-right"><Stat value={f ? gg / f : 0} /></td>
                  <td className="px-6 py-2.5 text-right"><Stat value={f ? d.ontS / f : 0} /></td>
                </tr>
              );
            })}
            <tr className="border-t-2 border-ink bg-paper-2/40">
              <th scope="row" className="px-6 py-3 text-left font-display text-xl">Grand Total</th>
              <td className="tnum px-3 py-3 text-right font-mono font-medium text-ink">{fmt(g.lateL)}</td>
              <td className="tnum px-3 py-3 text-right font-mono font-medium text-ink">{fmt(g.lateS)}</td>
              <td className="tnum px-3 py-3 text-right font-mono font-medium text-ink">{fmt(g.ontL)}</td>
              <td className="tnum px-3 py-3 text-right font-mono font-medium text-ink">{fmt(g.ontS)}</td>
              <td className="tnum px-3 py-3 text-right font-mono font-medium text-ink">{F.toLocaleString("id-ID")}</td>
              <td className="tnum px-3 py-3 text-right font-mono font-medium text-ink">{G.toLocaleString("id-ID")}</td>
              <td className="px-3 py-3 text-right"><Stat value={F ? G / F : 0} strong /></td>
              <td className="px-6 py-3 text-right"><Stat value={F ? g.ontS / F : 0} strong /></td>
            </tr>
          </tbody>
        </table>
      </div>
    </article>
  );
}

function Stat({ value, strong = false }) {
  const c = tone(value);
  return (
    <span className={`tnum inline-flex items-center justify-end gap-2 font-mono ${strong ? "font-medium" : ""}`}>
      <span aria-hidden="true" className="h-1.5 w-1.5 shrink-0 rounded-full" style={{ background: c }} />
      {pct(value)}
    </span>
  );
}

function Method() {
  const rows = [
    ["PERIODE", '=IF(O="","",TEXT(DATEVALUE(O),"MMM"))  →  bulan dari PLANNED RECEIPT DATE, jadi baris di SUMMARY.'],
    ["Planned + N", '=IF(O="","",DATEVALUE(O)+N)  →  batas toleransi, bawaan 7 hari.'],
    ["GAP", '=IF(OR(O="",R=0),-1,<Planned+N>−DATEVALUE(R))  →  selisih vs ARRIVAL DATE; belum datang berarti -1.'],
    ["Remarks", '=IF(GAP>=0,"On Time","Late")  →  jadi kolom Late / On Time di SUMMARY.'],
    ["SUMMARY B/C", '=SUMIFS(PO QTY / ARRIVAL QTY, PERIODE, bulan, Remarks, "Late")'],
    ["SUMMARY D/E", '=SUMIFS(PO QTY / ARRIVAL QTY, PERIODE, bulan, Remarks, "On Time")'],
    ["Total PO / Datang", "F = B + D    G = C + E"],
    ["FULFIL / OTIF", "H = G / F    I = E / F"],
  ];
  return (
    <section aria-labelledby="method-h" className="pt-16">
      <SectionLabel id="method-h">03 Metode</SectionLabel>
      <dl className="mt-4 grid grid-cols-1 gap-x-10 sm:grid-cols-[200px_1fr]">
        {rows.map(([k, v]) => (
          <div key={k} className="col-span-full grid grid-cols-1 gap-x-10 border-b border-rule py-3 sm:grid-cols-subgrid">
            <dt className="font-mono text-[11px] tracking-[0.14em] text-muted">{k}</dt>
            <dd className="mt-1 font-mono text-[12.5px] leading-relaxed text-ink-2 sm:mt-0">{v}</dd>
          </div>
        ))}
      </dl>
      <p className="mt-4 max-w-2xl text-xs leading-relaxed text-muted">
        SUMMARY memakai SUMIFS ke kolom bantuan, jadi tiap angka bisa ditelusuri: ubah tanggal atau
        qty di sheet Worksheet, rekap ikut berubah. Setiap rumus membawa nilai cache agar langsung
        tampil, dan berkas diset menghitung ulang saat dibuka.
      </p>
    </section>
  );
}
