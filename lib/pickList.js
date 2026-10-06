// Pick list allocation logic (browser + node dono me chalta hai)

const ALIASES = {
  sku: ["sku", "skuid", "skucode", "itemcode", "item", "article", "product", "productcode", "fsn", "barcode"],
  bin: ["bin", "binno", "binlocation", "location", "loc", "binid"],
  qty: ["qty", "quantity", "stock", "stockqty", "availableqty", "available", "requirement", "required", "requiredqty", "reqqty", "req"],
};

const norm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
const clean = (v) => String(v ?? "").trim();

function findCol(headers, key) {
  const h = headers.map(norm);
  for (const a of ALIASES[key]) {
    const i = h.indexOf(a);
    if (i !== -1) return i;
  }
  return -1;
}

function parseSheet(rows, sheetLabel, needBin) {
  if (!rows.length) throw new Error(`"${sheetLabel}" sheet khali hai.`);
  const headers = rows[0];
  const iSku = findCol(headers, "sku");
  const iQty = findCol(headers, "qty");
  const iBin = needBin ? findCol(headers, "bin") : -1;
  const missing = [];
  if (iSku < 0) missing.push("SKU");
  if (iQty < 0) missing.push("Qty");
  if (needBin && iBin < 0) missing.push("Bin");
  if (missing.length) {
    throw new Error(
      `"${sheetLabel}" sheet me ye column nahi mile: ${missing.join(", ")}. Mile huye headers: ${headers.map(clean).join(", ")}`
    );
  }
  const out = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const sku = clean(row[iSku]);
    if (!sku) continue;
    const qty = Number(String(row[iQty] ?? "").replace(/,/g, ""));
    out.push({ sku, bin: needBin ? clean(row[iBin]) : "", qty: Number.isFinite(qty) ? qty : 0 });
  }
  return out;
}

// Stock sheet = "stock" naam wali, Requirement sheet = "req" naam wali.
// Naam na mile to pehli sheet = stock, doosri = requirement.
export function pickSheets(sheetNames) {
  const low = sheetNames.map((n) => n.toLowerCase());
  let s = low.findIndex((n) => n.includes("stock") || n.includes("inventory"));
  let q = low.findIndex((n) => n.includes("req") || n.includes("order") || n.includes("demand"));
  if (s < 0 && q < 0) { s = 0; q = 1; }
  else if (s < 0) s = [0, 1].find((i) => i !== q) ?? 0;
  else if (q < 0) q = [0, 1].find((i) => i !== s) ?? 1;
  if (s === q || s >= sheetNames.length || q >= sheetNames.length) {
    throw new Error("Excel me 2 sheet chahiye: 1 Stock ki aur 1 Requirement ki.");
  }
  return { stockName: sheetNames[s], reqName: sheetNames[q] };
}

const binCmp = (a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });

export function buildPickList(stockRows, reqRows) {
  const stockRaw = parseSheet(stockRows, "Stock", true);
  const reqRaw = parseSheet(reqRows, "Requirement", false);

  // Stock: same SKU + same bin ki qty jod do
  const stock = new Map(); // sku -> Map(bin -> qty)
  for (const r of stockRaw) {
    if (r.qty <= 0 || !r.bin) continue;
    if (!stock.has(r.sku)) stock.set(r.sku, new Map());
    const m = stock.get(r.sku);
    m.set(r.bin, (m.get(r.bin) || 0) + r.qty);
  }

  // Requirement: same SKU ki qty jod do (order same rakho)
  const need = new Map();
  for (const r of reqRaw) {
    if (r.qty <= 0) continue;
    need.set(r.sku, (need.get(r.sku) || 0) + r.qty);
  }

  const result = [];
  for (const [sku, required] of need) {
    const bins = [...(stock.get(sku) || new Map()).entries()].map(([bin, qty]) => ({ bin, qty }));
    const total = bins.reduce((a, b) => a + b.qty, 0);

    if (total === 0) {
      result.push({ sku, bin: "", finalBin: "", qty: 0, remark: `Out of stock (Required ${required})` });
      continue;
    }

    // Rule 4: ek hi bin me poori qty ho to usi se (sabse chhota qualifying bin)
    const single = bins.filter((b) => b.qty >= required).sort((a, b) => a.qty - b.qty || binCmp(a.bin, b.bin))[0];
    if (single) {
      result.push({ sku, bin: single.bin, finalBin: `${sku} ${single.bin}`, qty: required, remark: "" });
      continue;
    }

    // Rule 5 & 6: alag-alag bins se (bada bin pehle), jitna mile utna
    bins.sort((a, b) => b.qty - a.qty || binCmp(a.bin, b.bin));
    let left = required;
    for (const b of bins) {
      if (left <= 0) break;
      const take = Math.min(b.qty, left);
      result.push({ sku, bin: b.bin, finalBin: `${sku} ${b.bin}`, qty: take, remark: "" });
      left -= take;
    }
    if (left > 0) {
      // Rule 7: remark us SKU ki last row par
      const last = result[result.length - 1];
      last.remark = `Short by ${left} (Required ${required}, Available ${total})`;
    }
  }

  // Rule 8: bin ke hisaab se sort (DF-2 < DF-10), phir sku. Bin-less (out of stock) rows end me.
  result.sort((a, b) => {
    if (!a.bin && b.bin) return 1;
    if (a.bin && !b.bin) return -1;
    return binCmp(a.bin, b.bin) || a.sku.localeCompare(b.sku, undefined, { numeric: true });
  });

  return result;
}

export function summarize(list) {
  const short = new Set(), oos = new Set();
  for (const r of list) {
    if (r.remark.startsWith("Out of stock")) oos.add(r.sku);
    else if (r.remark.startsWith("Short")) short.add(r.sku);
  }
  return { lines: list.length, short: short.size, outOfStock: oos.size };
}
