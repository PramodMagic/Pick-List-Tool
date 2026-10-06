"use client";
import { useRef, useState } from "react";
import * as XLSX from "xlsx";
import { buildPickList, pickSheets, summarize } from "../lib/pickList";

const card = { background: "#fff", borderRadius: 12, padding: 24, boxShadow: "0 1px 4px rgba(0,0,0,.08)" };

export default function Home() {
  const inputRef = useRef(null);
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState(null);

  async function handleFile(file) {
    setError(""); setInfo(null);
    if (!file) return;
    if (!/\.(xlsx|xls|xlsm)$/i.test(file.name)) { setError("Sirf Excel file (.xlsx / .xls) upload karein."); return; }
    setBusy(true);
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const { stockName, reqName } = pickSheets(wb.SheetNames);
      const toRows = (n) => XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: "", raw: false });
      const list = buildPickList(toRows(stockName), toRows(reqName));
      if (!list.length) throw new Error("Requirement sheet me koi valid row nahi mili.");

      const aoa = [["sku", "bin", "final bin", "qty", "remark"],
        ...list.map((r) => [r.sku, r.bin, r.finalBin, r.qty, r.remark])];
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 16 }, { wch: 12 }, { wch: 28 }, { wch: 8 }, { wch: 38 }];
      const out = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(out, ws, "Pick List");
      XLSX.writeFile(out, "pick_list.xlsx"); // auto download

      setInfo({ ...summarize(list), stockName, reqName, preview: list.slice(0, 15) });
    } catch (e) {
      setError(e.message || "Kuch galat ho gaya.");
    } finally {
      setBusy(false);
      if (inputRef.current) inputRef.current.value = "";
    }
  }

  function downloadTemplate() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["sku", "bin", "qty"], ["64059777", "DF-1", 5], ["64059777", "DF-2", 3]]), "Stock");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["sku", "qty"], ["64059777", 6]]), "Requirement");
    XLSX.writeFile(wb, "input_template.xlsx");
  }

  return (
    <main style={{ maxWidth: 820, margin: "0 auto", padding: "40px 16px" }}>
      <h1 style={{ margin: "0 0 4px" }}>Pick List Generator</h1>
      <p style={{ margin: "0 0 24px", color: "#5b6b7a" }}>
        Ek Excel upload karein: Sheet 1 = Stock (sku, bin, qty), Sheet 2 = Requirement (sku, qty).
      </p>

      <div style={card}>
        <div
          onClick={() => inputRef.current?.click()}
          onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
          onDragLeave={() => setDrag(false)}
          onDrop={(e) => { e.preventDefault(); setDrag(false); handleFile(e.dataTransfer.files?.[0]); }}
          style={{ border: `2px dashed ${drag ? "#2563eb" : "#b8c4d0"}`, background: drag ? "#eff6ff" : "#fafbfc",
                   borderRadius: 10, padding: "48px 16px", textAlign: "center", cursor: "pointer" }}
        >
          <div style={{ fontSize: 40 }}>📄</div>
          <div style={{ fontWeight: 600, marginTop: 8 }}>{busy ? "Pick list ban rahi hai..." : "Excel file yahan drop karein ya click karein"}</div>
          <div style={{ color: "#7b8794", fontSize: 13, marginTop: 4 }}>Generate hote hi pick_list.xlsx apne aap download hogi</div>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.xlsm" hidden onChange={(e) => handleFile(e.target.files?.[0])} />
        </div>
        <button onClick={downloadTemplate} style={{ marginTop: 14, background: "none", border: "none", color: "#2563eb", cursor: "pointer", padding: 0 }}>
          ⬇ Sample input template download karein
        </button>
      </div>

      {error && <div style={{ ...card, marginTop: 16, background: "#fef2f2", color: "#b91c1c" }}>⚠ {error}</div>}

      {info && (
        <div style={{ ...card, marginTop: 16 }}>
          <div style={{ fontWeight: 600, marginBottom: 8 }}>✅ Pick list download ho gayi</div>
          <div style={{ color: "#5b6b7a", fontSize: 14, marginBottom: 12 }}>
            Stock: <b>{info.stockName}</b> · Requirement: <b>{info.reqName}</b> · Total lines: <b>{info.lines}</b> ·
            Short SKU: <b>{info.short}</b> · Out of stock SKU: <b>{info.outOfStock}</b>
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
