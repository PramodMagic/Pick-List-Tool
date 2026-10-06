"use client";
import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { buildPickList, detectReq, detectStock, pickSheets, summarize } from "../lib/pickList";

const card = { background: "#fff", borderRadius: 12, padding: 24, boxShadow: "0 1px 4px rgba(0,0,0,.08)", marginTop: 16 };
const label = { display: "block", fontSize: 13, color: "#5b6b7a", marginBottom: 4 };
const select = { width: "100%", padding: "9px 10px", borderRadius: 8, border: "1px solid #b8c4d0", fontSize: 14, background: "#fff" };
const grid = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(210px, 1fr))", gap: 14 };

const readRows = (wb, name) =>
  XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "", raw: false });

function Field({ title, children }) {
  return (<div><label style={label}>{title}</label>{children}</div>);
}

function ColSelect({ headers, value, onChange }) {
  return (
    <select style={select} value={value} onChange={(e) => onChange(Number(e.target.value))}>
      <option value={-1}>— Column chuno —</option>
      {headers.map((h, i) => (
        <option key={i} value={i}>{XLSX.utils.encode_col(i)} — {String(h).trim() || "(blank)"}</option>
      ))}
    </select>
  );
}

export default function Home() {
  const inputRef = useRef(null);
  const [drag, setDrag] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState(null);
  const [wb, setWb] = useState(null);
  const [fileName, setFileName] = useState("");
  const [stockName, setStockName] = useState("");
  const [reqName, setReqName] = useState("");
  const [sm, setSm] = useState({ sku: -1, loc: -1, locType: "bin", qty: -1 });
  const [rm, setRm] = useState({ sku: -1, qty: -1 });

  const stockRows = useMemo(() => (wb && stockName ? readRows(wb, stockName) : []), [wb, stockName]);
  const reqRows = useMemo(() => (wb && reqName ? readRows(wb, reqName) : []), [wb, reqName]);
  const stockHeaders = stockRows[0] || [];
  const reqHeaders = reqRows[0] || [];

  function changeStockSheet(name) {
    setStockName(name);
    setSm(detectStock(readRows(wb, name)[0] || []));
  }
  function changeReqSheet(name) {
    setReqName(name);
    setRm(detectReq(readRows(wb, name)[0] || []));
  }

  async function handleFile(file) {
    setError(""); setInfo(null);
    if (!file) return;
    if (!/\.(xlsx|xls|xlsm)$/i.test(file.name)) { setError("Sirf Excel file (.xlsx / .xls) upload karein."); return; }
    try {
      const book = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const { stockName: s, reqName: r } = pickSheets(book.SheetNames);
      setWb(book); setFileName(file.name);
      setStockName(s); setReqName(r);
      setSm(detectStock(readRows(book, s)[0] || []));
      setRm(detectReq(readRows(book, r)[0] || []));
    } catch (e) {
      setWb(null);
      setError(e.message || "File padhne me dikkat aayi.");
    } finally {
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  function generate() {
    setError(""); setInfo(null);
    try {
      const list = buildPickList(stockRows, reqRows, sm, rm);
      if (!list.length) throw new Error("Requirement sheet me koi valid row (qty > 0) nahi mili.");
      const aoa = [["sku", "bin", "final bin", "qty", "remark"],
        ...list.map((r) => [r.sku, r.bin, r.finalBin, r.qty, r.remark])];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 16 }, { wch: 12 }, { wch: 28 }, { wch: 8 }, { wch: 38 }];
      const out = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(out, ws, "Pick List");
      XLSX.writeFile(out, "pick_list.xlsx"); // auto download
      setInfo({ ...summarize(list), preview: list.slice(0, 15) });
    } catch (e) {
      setError(e.message || "Kuch galat ho gaya.");
    }
  }

  function downloadTemplate() {
    const w = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([["SKU", "Final Barcode", "Closing Stock"], ["64059777", "64059777 DF-1", 5], ["64059777", "64059777 DF-2", 3]]), "Stock");
    XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet([["sku", "qty"], ["64059777", 6]]), "Requirement");
    XLSX.writeFile(w, "input_template.xlsx");
  }

  const sheetNames = wb ? wb.SheetNames : [];
  const SheetSelect = ({ value, onChange }) => (
    <select style={select} value={value} onChange={(e) => onChange(e.target.value)}>
      {sheetNames.map((n) => <option key={n} value={n}>{n}</option>)}
    </select>
  );
  const ready = sm.sku >= 0 && sm.loc >= 0 && sm.qty >= 0 && rm.sku >= 0 && rm.qty >= 0 && stockName !== reqName;

  return (
    <main style={{ maxWidth: 860, margin: "0 auto", padding: "40px 16px" }}>
      <h1 style={{ margin: "0 0 4px" }}>Pick List Generator</h1>
      <p style={{ margin: "0 0 8px", color: "#5b6b7a" }}>
        Excel upload karein, phir apne columns chunkar Pick List generate karein.
      </p>

      <div style={{ ...card, marginTop: 16 }}>
        <div
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files?.[0]); }}
          style={{ border: `2px dashed ${drag ? "#2563eb" : "#b8c4d0"}`, background: drag ? "#eff6ff" : "#fafbfc",
                   borderRadius: 10, padding: wb ? "20px 16px" : "48px 16px", textAlign: "center", cursor: "pointer" }}
        >
          <div style={{ fontSize: wb ? 24 : 40 }}>📄</div>
          <div style={{ fontWeight: 600, marginTop: 6 }}>
            {wb ? `${fileName} (badalne ke liye click karein)` : "Excel file yahan drop karein ya click karein"}
          </div>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm" hidden onChange={(e) => handleFile(e.target.files?.[0])} />
        </div>
        <button onClick={downloadTemplate} style={{ marginTop: 14, background: "none", border: "none", color: "#2563eb", cursor: "pointer", padding: 0 }}>
          ⬇ Sample input template download karein
        </button>
      </div>

      {wb && (
        <>
          <div style={card}>
            <div style={{ fontWeight: 600, marginBottom: 12 }}>1. Stock sheet</div>
            <div style={grid}>
              <Field title="Sheet"><SheetSelect value={stockName} onChange={changeStockSheet} /></Field>
              <Field title="SKU column"><ColSelect headers={stockHeaders} value={sm.sku} onChange={(v) => setSm({ ...sm, sku: v })} /></Field>
              <Field title="Bin / Final Barcode column"><ColSelect headers={stockHeaders} value={sm.loc} onChange={(v) => setSm({ ...sm, loc: v })} /></Field>
              <Field title="Is column me kya hai?">
                <select style={select} value={sm.locType} onChange={(e) => setSm({ ...sm, locType: e.target.value })}>
                  <option value="bin">Sirf Bin (jaise DF-1)</option>
                  <option value="final">Final Barcode (jaise 64059777 DF-1)</option>
                </select>
              </Field>
              <Field title="Stock Qty column"><ColSelect headers={stockHeaders} value={sm.qty} onChange={(v) => setSm({ ...sm, qty: v })} /></Field>
            </div>
          </div>

          <div style={card}>
            <div style={{ fontWeight: 600, marginBottom: 12 }}>2. Requirement sheet</div>
            <div style={grid}>
              <Field title="Sheet"><SheetSelect value={reqName} onChange={changeReqSheet} /></Field>
              <Field title="SKU column"><ColSelect headers={reqHeaders} value={rm.sku} onChange={(v) => setRm({ ...rm, sku: v })} /></Field>
              <Field title="Requirement Qty column"><ColSelect headers={reqHeaders} value={rm.qty} onChange={(v) => setRm({ ...rm, qty: v })} /></Field>
            </div>
            {stockName === reqName && <div style={{ color: "#b45309", marginTop: 12, fontSize: 14 }}>Stock aur Requirement ke liye alag-alag sheet chuno.</div>}
          </div>

          <button onClick={generate} disabled={!ready}
            style={{ marginTop: 16, width: "100%", padding: "14px", fontSize: 16, fontWeight: 600, border: "none", borderRadius: 10,
                     color: "#fff", background: ready ? "#2563eb" : "#9db4d9", cursor: ready ? "pointer" : "not-allowed" }}>
            Pick List Generate & Download
          </button>
        </>
      )}

      {error && <div style={{ ...card, background: "#fef2f2", color: "#b91c1c" }}>⚠ {error}</div>}

      {info && (
        <div style={card}>
          <div style={{ fontWeight: 600, marginBottom: 8 }}>✅ Pick list download ho gayi</div>
          <div style={{ color: "#5b6b7a", fontSize: 14, marginBottom: 12 }}>
            Total lines: <b>{info.lines}</b> · Short SKU: <b>{info.short}</b> · Out of stock SKU: <b>{info.outOfStock}</b>
          </div>
          <div style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 14 }}>
              <thead><tr>{["sku", "bin", "final bin", "qty", "remark"].map((h) => (
                <th key={h} style={{ textAlign: "left", padding: 8, borderBottom: "2px solid #e5e9ee" }}>{h}</th>))}</tr></thead>
              <tbody>{info.preview.map((r, i) => (
                <tr key={i} style={{ background: r.remark ? "#fff7ed" : "transparent" }}>
                  <td style={{ padding: 8, borderBottom: "1px solid #eef1f4" }}>{r.sku}</td>
                  <td style={{ padding: 8, borderBottom: "1px solid #eef1f4" }}>{r.bin}</td>
                  <td style={{ padding: 8, borderBottom: "1px solid #eef1f4" }}>{r.finalBin}</td>
                  <td style={{ padding: 8, borderBottom: "1px solid #eef1f4" }}>{r.qty}</td>
                  <td style={{ padding: 8, borderBottom: "1px solid #eef1f4", color: "#c2410c" }}>{r.remark}</td>
                </tr>))}</tbody>
            </table>
          </div>
          {info.lines > 15 && <div style={{ color: "#7b8794", fontSize: 13, marginTop: 8 }}>Preview me pehli 15 lines. Poori list file me hai.</div>}
        </div>
      )}
    </main>
  );
}
