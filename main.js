const $ = (id) => document.getElementById(id);

const els = {
  sidebar: $("sidebar"),
  scrim: $("scrim"),
  navToggle: $("navToggle"),
  navClose: $("navClose"),

  file: $("file"),
  fileHint: $("file-hint"),
  status: $("status"),

  subject: $("subject"),
  docDate: $("docDate"),
  purpose: $("purpose"),
  notes: $("notes"),

  btnExtract: $("btn-extract"),
  btnAutoOutline: $("btn-auto-outline"),
  btnCopyPurpose: $("btn-copy-purpose"),

  extractSummary: $("extractSummary"),
  itemsTbody: $("itemsTbody"),
};

const state = {
  extracted: {
    source: null,
    items: [],
    total: null,
    rawText: "",
    rawRows: null, // for assisted extraction on messy formats
    pageImages: null, // data URLs for scanned/image-only PDFs (AI vision)
  },
};

function blankExtracted() {
  return { source: null, items: [], total: null, rawText: "", rawRows: null, pageImages: null };
}

function todayISO() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function setStatus(msg) {
  const m = msg ? String(msg) : "";
  if (els.status) els.status.textContent = m;
  if (m) console.log(m);
}

function fmtMoney(v) {
  if (v === null || v === undefined || v === "") return "";
  const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
  if (!Number.isFinite(n)) return String(v);
  return new Intl.NumberFormat("ko-KR").format(n);
}

function hasBatchim(s) {
  const t = String(s || "").trim();
  if (!t) return false;
  const ch = t.charCodeAt(t.length - 1);
  if (ch < 0xac00 || ch > 0xd7a3) return false;
  return (ch - 0xac00) % 28 !== 0;
}

function josa(word, a, b) {
  // Choose particle based on 받침 여부.
  return hasBatchim(word) ? a : b;
}

function numToKorean(n) {
  // Sino-Korean number words (good enough for currency amounts).
  const units = ["", "만", "억", "조", "경"];
  const digits = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
  const small = ["", "십", "백", "천"];

  const num = Math.floor(Number(n));
  if (!Number.isFinite(num) || num <= 0) return "";

  const chunk = (x) => {
    let out = "";
    const ds = String(x).padStart(4, "0").split("").map((c) => Number(c));
    for (let i = 0; i < 4; i++) {
      const d = ds[i];
      if (!d) continue;
      const pos = 3 - i;
      out += (d === 1 && pos > 0 ? "" : digits[d]) + small[pos];
    }
    return out;
  };

  let x = num;
  let i = 0;
  let out = "";
  while (x > 0 && i < units.length) {
    const part = x % 10000;
    if (part) {
      const c = chunk(part);
      out = c + units[i] + out;
    }
    x = Math.floor(x / 10000);
    i++;
  }
  return out;
}

function numToKoreanWon(n) {
  const intN = Math.floor(Number(n));
  if (!Number.isFinite(intN) || intN < 0) return "";
  if (intN === 0) return "금영원";
  const words = numToKorean(intN);
  if (!words) return "";
  return `금${words}원`;
}

function budgetLine(n) {
  const intN = Math.floor(Number(n));
  const safe = Number.isFinite(intN) && intN >= 0 ? intN : 0;
  return `금${fmtMoney(safe)}원(${numToKoreanWon(safe)})`;
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function copyText(text) {
  const t = String(text || "").trim();
  if (!t) return false;
  try {
    await navigator.clipboard.writeText(t);
    return true;
  } catch {
    // Fallback: temporary textarea
    const ta = document.createElement("textarea");
    ta.value = t;
    ta.setAttribute("readonly", "true");
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  }
}

function normalizeItems(items) {
  const toNum = (v) => {
    if (v === undefined || v === null || v === "") return "";
    if (typeof v === "number") return Number.isFinite(v) ? v : "";
    const t = String(v).replace(/[^\d.\-]/g, "");
    if (!t) return "";
    const n = Number(t);
    return Number.isFinite(n) ? n : "";
  };

  return (items || [])
    .map((it) => ({
      name: (it?.name ?? "").toString().trim(),
      spec: (it?.spec ?? it?.note ?? "").toString().trim(),
      qty: toNum(it?.qty),
      unitPrice:
        toNum(it?.unitPrice),
      amount: toNum(it?.amount),
    }))
    .filter((it) => it.name || it.amount || it.unitPrice || it.spec);
}

function computeTotal(items) {
  const nums = (items || [])
    .map((it) => {
      if (Number.isFinite(it.amount)) return it.amount;
      if (Number.isFinite(it.qty) && Number.isFinite(it.unitPrice)) return it.qty * it.unitPrice;
      return 0;
    })
    .filter((n) => Number.isFinite(n));
  const t = nums.reduce((a, b) => a + b, 0);
  return t > 0 ? t : null;
}

function renderExtractSummary() {
  const { source, items, total } = state.extracted;
  const parts = [];
  if (source) parts.push(`source: ${source}`);
  parts.push(`items: ${items.length}`);
  if (total) parts.push(`total: ${fmtMoney(total)} KRW`);
  els.extractSummary.textContent = parts.join(" | ");
}

function renderItems(items) {
  els.itemsTbody.innerHTML = "";

  const normalized = normalizeItems(items);
  normalized.forEach((it, idx) => {
    const tr = document.createElement("tr");
    const amount = it.amount === "" ? "" : fmtMoney(it.amount);
    const qty = it.qty === "" ? "" : fmtMoney(it.qty);
    const unit = it.unitPrice === "" ? "" : fmtMoney(it.unitPrice);

    tr.innerHTML = `
      <td class="num">${idx + 1}</td>
      <td>
        <div class="cell-copy">
          <div class="text">${escapeHtml(it.name || "")}</div>
          <button class="pill js-copy" type="button" data-copy="${escapeHtml(it.name || "")}">내용복사</button>
        </div>
      </td>
      <td>${escapeHtml(it.spec || "")}</td>
      <td class="num">${escapeHtml(qty)}</td>
      <td class="num">${escapeHtml(unit)}</td>
      <td class="num">${escapeHtml(amount)}</td>
    `;
    els.itemsTbody.appendChild(tr);
  });
}

async function loadScript(src) {
  // Cache loads so multiple extracts don't inject duplicate tags.
  loadScript._p ||= new Map();
  if (loadScript._p.has(src)) return loadScript._p.get(src);

  const p = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(s);
  });
  loadScript._p.set(src, p);
  return p;
}

async function loadPdfJs() {
  loadPdfJs._p ||= (async () => {
    // Use ESM builds (static-hosting friendly). Try multiple CDNs for reliability.
    try {
      const url = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.mjs";
      const mod = await import(url);
      const workerUrl = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.mjs";
      mod.GlobalWorkerOptions.workerSrc = workerUrl;
      return mod;
    } catch (e) {
      try {
        const url = "https://unpkg.com/pdfjs-dist@4.10.38/build/pdf.mjs";
        const mod = await import(url);
        mod.GlobalWorkerOptions.workerSrc = "https://unpkg.com/pdfjs-dist@4.10.38/build/pdf.worker.mjs";
        return mod;
      } catch (e2) {
        const url = "https://unpkg.com/pdfjs-dist@4.10.38/legacy/build/pdf.mjs";
        const mod = await import(url);
        mod.GlobalWorkerOptions.workerSrc = "https://unpkg.com/pdfjs-dist@4.10.38/legacy/build/pdf.worker.mjs";
        return mod;
      }
    }
  })();
  return loadPdfJs._p;
}

function isSmallScreen() {
  try {
    return window.matchMedia && window.matchMedia("(max-width: 820px)").matches;
  } catch {
    return false;
  }
}

async function renderPdfPageToCanvas(pdfDoc, pageNumber, scale) {
  const page = await pdfDoc.getPage(pageNumber);
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d", { willReadFrequently: true });
  canvas.width = Math.floor(viewport.width);
  canvas.height = Math.floor(viewport.height);

  await page.render({ canvasContext: ctx, viewport }).promise;
  return canvas;
}

async function renderPdfPageToDataUrl(pdfDoc, pageNumber, { scale, quality }) {
  const tries = [
    { scale: scale ?? 1.15, quality: quality ?? 0.72 },
    { scale: Math.max(0.9, (scale ?? 1.15) * 0.9), quality: 0.65 },
    { scale: 0.9, quality: 0.6 },
  ];

  for (const t of tries) {
    const canvas = await renderPdfPageToCanvas(pdfDoc, pageNumber, t.scale);
    const dataUrl = canvas.toDataURL("image/jpeg", t.quality);
    // Keep under server guardrail (~2.5MB per image).
    if (dataUrl.length <= 2400000) return dataUrl;
  }

  // Last resort: return the last attempt even if large; server may drop it.
  const canvas = await renderPdfPageToCanvas(pdfDoc, pageNumber, 0.9);
  return canvas.toDataURL("image/jpeg", 0.6);
}

async function extractPdfTextQuick(file, { maxPages, onProgress } = {}) {
  const ab = await file.arrayBuffer();
  const pdfjs = await loadPdfJs();

  // If worker init fails on some iOS environments, run without worker.
  if (pdfjs?.GlobalWorkerOptions && !pdfjs.GlobalWorkerOptions.workerSrc) {
    try {
      pdfjs.disableWorker = true;
    } catch {
      // ignore
    }
  }

  const loadingTask = pdfjs.getDocument({ data: ab });
  const doc = await loadingTask.promise;

  let all = "";
  const pages = maxPages ? Math.min(doc.numPages, Math.max(1, maxPages)) : doc.numPages;
  for (let p = 1; p <= pages; p++) {
    onProgress?.(`PDF 텍스트 추출 중... (${p}/${pages})`);
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const text = content.items.map((it) => it.str).join(" ");
    all += text + "\n";
  }

  return all.trim();
}

async function extractScannedPdfViaAi(file, { scale, quality, chunkSize }) {
  const pdfjs = await loadPdfJs();
  const ab = await file.arrayBuffer();
  const pdfDoc = await pdfjs.getDocument({ data: ab }).promise;

  const totalPages = pdfDoc.numPages || 1;
  const chunk = Math.max(1, Number(chunkSize) || 4);

  const allItems = [];
  let shipping = 0;
  let discount = 0;
  let statedTotal = 0;

  for (let start = 1; start <= totalPages; start += chunk) {
    const end = Math.min(totalPages, start + chunk - 1);
    const imgs = [];

    for (let p = start; p <= end; p++) {
      setStatus(`스캔 PDF 페이지 렌더링 중... (${p}/${totalPages})`);
      imgs.push(await renderPdfPageToDataUrl(pdfDoc, p, { scale, quality }));
    }

    setStatus(`AI 추출 중... (${start}-${end}/${totalPages})`);
    const data = await callExtractApi({
      source: "pdf",
      filename: file.name,
      rows: null,
      rawText: "",
      pageImages: imgs,
    });

    const items = normalizeItems(data?.items || []);
    allItems.push(...items);

    const ship = Number.isFinite(data?.shipping) ? data.shipping : Number.isFinite(data?.totals?.shipping) ? data.totals.shipping : 0;
    const disc = Number.isFinite(data?.discount) ? data.discount : Number.isFinite(data?.totals?.discount) ? data.totals.discount : 0;
    const stated = Number.isFinite(data?.statedTotal) ? data.statedTotal : 0;

    // Prefer the maximum shipping/discount candidate across chunks (avoids double count when repeated).
    if (ship > shipping) shipping = ship;
    if (disc > discount) discount = disc;
    if (stated > statedTotal) statedTotal = stated;
  }

  const subtotal = computeTotal(allItems) ?? 0;
  const total = subtotal + (shipping || 0) - (discount || 0);

  return {
    items: allItems,
    total: total > 0 ? total : null,
    shipping: shipping || null,
    discount: discount || null,
    statedTotal: statedTotal || null,
  };
}

function looksLikeHtmlBytes(bytes) {
  // Some marketplaces export "xls" as HTML tables.
  // Detect by first non-whitespace byte being '<' (0x3c) or BOM + '<'.
  let i = 0;
  while (i < bytes.length && (bytes[i] === 0x20 || bytes[i] === 0x0a || bytes[i] === 0x0d || bytes[i] === 0x09)) i++;
  if (i + 3 < bytes.length && bytes[i] === 0xef && bytes[i + 1] === 0xbb && bytes[i + 2] === 0xbf) i += 3;
  return i < bytes.length && bytes[i] === 0x3c;
}

function tableToRows(table) {
  const rows = [];
  for (const tr of table.querySelectorAll("tr")) {
    const cells = Array.from(tr.querySelectorAll("th,td")).map((td) =>
      String(td.textContent || "")
        .replace(/\u00a0/g, " ")
        .replace(/\s+/g, " ")
        .trim()
    );
    if (cells.some(Boolean)) rows.push(cells);
  }
  return rows;
}

function pickBestTable(doc) {
  const tables = Array.from(doc.querySelectorAll("table"));
  if (!tables.length) return null;

  let best = { score: -1, rows: null };

  for (const t of tables) {
    const rows = tableToRows(t);
    if (rows.length < 2) continue;

    let score = 0;
    const headerIdx = findHeaderRow(rows, { maxScan: Math.min(rows.length, 400) });
    if (headerIdx >= 0) {
      // Strongly prefer tables that look like an item list (header + many data rows).
      score += 1000;
      score += Math.max(0, 200 - headerIdx); // earlier header is better
      const header = rows[headerIdx].map((c) => String(c || "").trim());
      const col = mapCols(header);
      let itemCount = 0;
      for (let i = headerIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        const name = pickCell(r, col.name);
        const spec = pickCell(r, col.spec);
        const unitPrice = parseNum(pickCell(r, col.unitPrice));
        const amount = parseNum(pickCell(r, col.amount));
        if (!name && !Number.isFinite(amount) && !Number.isFinite(unitPrice) && !spec) continue;
        itemCount++;
      }
      score += Math.min(300, itemCount) * 10;
    } else {
      // Fallback heuristic: keyword presence within the first N rows.
      const joined = rows
        .slice(0, 200)
        .map((r) => r.join(" "))
        .join(" ");
      const keys = ["상품명", "품목", "항목", "내용", "규격", "수량", "단가", "판매가", "공급가액", "공급합계", "금액", "합계", "총액"];
      for (const k of keys) if (joined.includes(k)) score += 2;
    }

    const maxCols = Math.max(...rows.map((r) => r.length));
    score += Math.min(20, rows.length) + Math.min(10, maxCols);
    if (score > best.score) best = { score, rows };
  }
  return best.rows;
}

function extractTotalsFromHtmlDoc(doc) {
  const text = String(doc?.body?.textContent || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const findAll = (labels) => {
    const out = [];
    for (const label of labels) {
      // Capture a KRW amount that appears after the label.
      const re = new RegExp(`${label}\\s*([0-9][0-9,]*)\\s*원`);
      const m = text.match(re);
      if (!m) continue;
      const n = parseNum(m[1]);
      if (Number.isFinite(n)) out.push(n);
    }
    return out;
  };

  // Many HTML-exported "xls" quote formats include these labels (e.g., Auction/Gmarket).
  const grandCandidates = findAll(["총 구매금액", "총구매금액", "결제금액", "총액", "총 금액"]);
  const shippingCandidates = findAll(["배송비"]);
  const discountCandidates = findAll(["할인금액", "할인 금액"]);

  const pickMax = (arr) => (arr.length ? Math.max(...arr) : null);
  return {
    grandTotal: pickMax(grandCandidates),
    shipping: pickMax(shippingCandidates),
    discount: pickMax(discountCandidates),
  };
}

async function extractFromHtmlXls(file) {
  const html = await file.text();
  const doc = new DOMParser().parseFromString(html, "text/html");
  const rows = pickBestTable(doc) || [];
  return rows;
}

async function extractFromXlsx(file) {
  // Some .xls files are actually HTML. Detect and parse accordingly.
  const head = new Uint8Array(await file.slice(0, 256).arrayBuffer());
  let rows;
  let htmlTotals = null;
  if (looksLikeHtmlBytes(head)) {
    const html = await file.text();
    const doc = new DOMParser().parseFromString(html, "text/html");
    rows = pickBestTable(doc) || [];
    htmlTotals = extractTotalsFromHtmlDoc(doc);
  } else {
    if (!window.XLSX) {
      await loadScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
      if (!window.XLSX) throw new Error("엑셀 파서 로드에 실패했습니다. 네트워크 상태를 확인하세요.");
    }
    const ab = await file.arrayBuffer();
    try {
      const wb = window.XLSX.read(ab, { type: "array" });
      const firstName = wb.SheetNames[0];
      const ws = wb.Sheets[firstName];
      rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
    } catch (e) {
      // Fallback: try HTML parse even if the extension says xls/xlsx.
      console.warn("XLSX.read failed; falling back to HTML parse:", e);
      const html = await file.text();
      const doc = new DOMParser().parseFromString(html, "text/html");
      rows = pickBestTable(doc) || [];
      htmlTotals = extractTotalsFromHtmlDoc(doc);
    }
  }

  const headerIdx = findHeaderRow(rows, { maxScan: Math.min(rows.length, 400) });
  const items = [];
  if (headerIdx >= 0) {
    const header = rows[headerIdx].map((c) => String(c).trim());
    const col = mapCols(header);
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = pickCell(r, col.name);
      const spec = pickCell(r, col.spec);
      const qty = parseNum(pickCell(r, col.qty));
      const unitPrice = parseNum(pickCell(r, col.unitPrice));
      const amount = parseNum(pickCell(r, col.amount));
      if (!name && !Number.isFinite(amount) && !Number.isFinite(unitPrice) && !spec) continue;
      items.push({ name, spec, qty, unitPrice, amount });
    }
  }

  const normalized = normalizeItems(items);
  const computed = computeTotal(normalized);
  const total =
    Number.isFinite(htmlTotals?.grandTotal) ? htmlTotals.grandTotal : computed !== null ? computed : null;
  // Keep more rows for server-assisted extraction when local header detection fails.
  state.extracted = { source: "xlsx", items: normalized, total, rawText: "", rawRows: rows.slice(0, 2000) };
}

function findHeaderRow(rows, { maxScan = 250 } = {}) {
  const nameKeys = ["상품명", "품목", "항목", "내용", "내역", "상품"];
  const qtyOrMoneyKeys = [
    "수량",
    "주문수량",
    "구매수량",
    "단가",
    "판매가",
    "금액",
    "공급가액",
    "공급합계",
    "합계",
    "총액",
    "총금액",
    "결제금액",
  ];

  const scanN = Math.min(rows.length, Math.max(1, Number(maxScan) || 250));
  for (let i = 0; i < scanN; i++) {
    const row = (rows[i] || []).map((c) => String(c || "").trim());
    if (row.length < 4) continue;
    const joined = row.join(" ");
    const hasName = nameKeys.some((k) => joined.includes(k));
    const hasQtyOrMoney = qtyOrMoneyKeys.some((k) => joined.includes(k));
    if (hasName && hasQtyOrMoney) return i;
  }
  return -1;
}

function mapCols(header) {
  const idxOf = (pred) => {
    for (let i = 0; i < header.length; i++) if (pred(header[i])) return i;
    return -1;
  };
  const by = (arr) => (h) => arr.some((k) => h.includes(k));

  const name = idxOf(by(["품목", "항목", "내용", "제품", "서비스", "내역", "상품명", "상품", "상품정보"]));
  const spec = idxOf(by(["규격", "사양", "옵션", "모델", "모델명", "spec", "SPEC"]));
  const qty = idxOf(by(["수량", "수 량", "qty", "QTY", "주문수량", "구매수량"]));
  const unitPrice = idxOf(by(["단가", "단 가", "판매가", "판매 단가", "unit", "price", "단위금액", "단위 금액"]));
  // Prefer "line total" columns (after discount) when present, e.g. '공급합계'.
  const amount = idxOf(
    by([
      "공급합계",
      "공급 합계",
      "결제금액",
      "총액",
      "총금액",
      "합계금액",
      "합계 금액",
      "금액",
      "공급가액",
      "공급가",
      "주문금액",
      "amount",
    ])
  );

  return { name, spec, qty, unitPrice, amount };
}

function pickCell(row, idx) {
  if (idx < 0) return "";
  const v = row?.[idx];
  return v === undefined || v === null ? "" : String(v).trim();
}

function parseNum(s) {
  const t = String(s ?? "").replace(/[^\d.\-]/g, "");
  if (!t) return "";
  const n = Number(t);
  return Number.isFinite(n) ? n : "";
}

async function extractFromPdf(file) {
  const small = isSmallScreen();
  const maxPages = null; // all pages

  setStatus("PDF 읽는 중... (전체 페이지)");
  let text = "";
  try {
    text = await extractPdfTextQuick(file, { maxPages, onProgress: (m) => setStatus(m) });
  } catch (e) {
    console.warn("pdf text extract failed:", e);
  }

  // If text is too short, treat as scanned/image-only PDF and send page images for AI vision.
  const compactLen = text.replace(/\s+/g, "").length;
  if (compactLen < 80) {
    setStatus("텍스트가 거의 없어 스캔본으로 판단했습니다. 전체 페이지 AI 추출을 진행합니다... (GPT-5.2)");
    const res = await extractScannedPdfViaAi(file, {
      scale: small ? 1.0 : 1.15,
      quality: 0.72,
      chunkSize: small ? 2 : 4,
    });

    state.extracted = {
      source: "pdf",
      items: res.items,
      total: res.total,
      rawText: "",
      rawRows: null,
      pageImages: null,
    };
    return;
  }

  state.extracted = {
    source: "pdf",
    items: [],
    total: null,
    rawText: String(text || "").trim(),
    rawRows: null,
    pageImages: null,
  };
}

function toDocPayload() {
  const items = normalizeItems(state.extracted.items);
  // Prefer server-determined grand total (includes shipping/discount) over recomputing from line items.
  const total = state.extracted.total ?? computeTotal(items);

  return {
    meta: {
      subject: els.subject.value.trim(),
      docDate: (els.docDate.value || "").trim(),
      purpose: els.purpose.value.trim(),
      notes: els.notes.value.trim(),
    },
    quote: {
      source: state.extracted.source,
      currency: "KRW",
      items,
      total,
      rawText: state.extracted.rawText,
    },
  };
}

function validateForGenerate(payload) {
  const errs = [];
  if (!payload.meta.subject) errs.push("제목을 입력하세요.");
  // Allow generating even without a quote; we won't invent numbers and will mark unknowns as '미기재'.
  return errs;
}

async function callExtractApi({ source, filename, rows, rawText, pageImages }) {
  const res = await fetch("/api/extract", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ source, filename, rows, rawText, pageImages, forceAI: true }),
  });

  const txt = await res.text();
  let data;
  try {
    data = JSON.parse(txt);
  } catch {
    throw new Error(`API 응답 파싱 실패: ${txt.slice(0, 200)}`);
  }
  if (!res.ok) throw new Error(data?.error || `API 오류 (HTTP ${res.status})`);
  return data;
}

async function callOutlineApi(payload) {
  const res = await fetch("/api/outline", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ payload }),
  });

  const txt = await res.text();
  let data;
  try {
    data = JSON.parse(txt);
  } catch {
    throw new Error(`API 응답 파싱 실패: ${txt.slice(0, 200)}`);
  }
  if (!res.ok) throw new Error(data?.error || `API 오류 (HTTP ${res.status})`);
  return data;
}

function buildOfficialOutline(payload) {
  const items = payload?.quote?.items || [];
  const total = payload?.quote?.total ?? computeTotal(items) ?? null;

  const itemNames = items
    .map((it) => String(it?.name || "").trim())
    .filter(Boolean)
    .slice(0, 6);
  const itemLine = itemNames.length ? itemNames.join(", ") + (items.length > itemNames.length ? " 등" : "") : "미기재";

  const lines = items
    .filter((it) => it && (it.amount || (Number.isFinite(it.qty) && Number.isFinite(it.unitPrice))))
    .slice(0, 6)
    .map((it) => {
      const qty = Number.isFinite(it.qty) ? it.qty : null;
      const unit = Number.isFinite(it.unitPrice) ? it.unitPrice : null;
      const amt = Number.isFinite(it.amount) ? it.amount : qty !== null && unit !== null ? qty * unit : null;
      if (qty !== null && unit !== null && amt !== null) return `${fmtMoney(unit)}원 X ${fmtMoney(qty)}개 = ${fmtMoney(amt)}원`;
      if (amt !== null) return `${fmtMoney(amt)}원`;
      return "";
    })
    .filter(Boolean);

  const purpose = payload?.meta?.purpose?.trim();
  const purposeLine = purpose ? purpose.replace(/\s+/g, " ").slice(0, 160) : "미기재";
  const totalN = Number.isFinite(total) && total >= 0 ? total : 0;

  const subject = String(payload?.meta?.subject || "").trim() || "미기재";
  const buySentence = `${subject}${josa(subject, "을", "를")} 다음과 같이 구입하고자 합니다.`;

  const basisLine = lines.length ? lines.join(" / ") : total ? `${fmtMoney(total)}원` : "미기재";

  return (
    `${buySentence}\n` +
    `1. 목적: ${purposeLine}\n` +
    `2. 품명: ${itemLine}\n` +
    `3. 소요 예산: ${budgetLine(totalN)}\n` +
    `4. 산출 근거: ${basisLine}\n\n` +
    "붙임 지출품의서 1부. 끝."
  );
}

function setBusy(busy) {
  const b = Boolean(busy);
  els.btnExtract.disabled = b;
  els.btnCopyPurpose.disabled = b;
  els.btnAutoOutline.disabled = b;
}

async function onAutoOutline() {
  const payload = toDocPayload();
  if (!payload.meta.subject) {
    alert("제목을 먼저 입력하세요.");
    return;
  }

  setBusy(true);
  setStatus("개요 자동 생성 중...");
  try {
    const data = await callOutlineApi(payload);
    const outline = String(data?.outline || "").trim();
    if (!outline) throw new Error("개요 생성 결과가 비어 있습니다.");
    // Enforce the budget line format: 금0,000원(금영원)
    els.purpose.value = forceBudgetLine(outline, payload);
    setStatus(data.mode === "ai" ? "자동 생성이 완료되었습니다." : "개요가 생성되었습니다.");
  } catch (e) {
    console.error(e);
    // Fallback to deterministic outline
    els.purpose.value = buildOfficialOutline(payload);
    setStatus(`자동 생성 연결 문제로 기본 개요로 채웠습니다: ${e?.message || String(e)}`);
  } finally {
    setBusy(false);
  }
}

function forceBudgetLine(outline, payload) {
  const items = payload?.quote?.items || [];
  const total = payload?.quote?.total ?? computeTotal(items) ?? 0;
  const line = `3. 소요 예산: ${budgetLine(total)}`;
  const lines = String(outline || "").split(/\r?\n/);
  const idx = lines.findIndex((l) => l.trim().startsWith("3. 소요 예산:"));
  if (idx >= 0) {
    lines[idx] = line;
    return lines.join("\n").trim();
  }
  const idx2 = lines.findIndex((l) => l.trim().startsWith("2. 품명:"));
  if (idx2 >= 0) {
    lines.splice(idx2 + 1, 0, line);
    return lines.join("\n").trim();
  }
  return (String(outline || "").trim() + "\n" + line).trim();
}

async function extractOneFile(f) {
  // Reset state per file so XLSX/PDF extractors can keep their current contract.
  state.extracted = blankExtracted();

  const lower = (f.name || "").toLowerCase();
  const isXls = lower.endsWith(".xlsx") || lower.endsWith(".xls");
  const isPdf = lower.endsWith(".pdf") || f.type === "application/pdf";
  if (isXls) {
    await extractFromXlsx(f);
  } else if (isPdf) {
    await extractFromPdf(f);
  } else {
    throw new Error(`지원하지 않는 파일 형식입니다: ${f.name || "unknown"} (.xls, .xlsx, .pdf)`);
  }

  // If heuristic extraction failed, use the server-assisted structuring if configured.
  if (!state.extracted.items?.length) {
    try {
      const rows = state.extracted.rawRows;
      const rawText = state.extracted.rawText;
      const data = await callExtractApi({
        source: state.extracted.source,
        filename: f.name,
        rows,
        rawText,
        pageImages: state.extracted.pageImages,
      });

      const items = normalizeItems(data?.items || []);
      const apiTotal =
        Number.isFinite(data?.totals?.grandTotal) ? data.totals.grandTotal : Number.isFinite(data?.total) ? data.total : null;
      const total = apiTotal ?? computeTotal(items) ?? null;
      if (items.length) {
        state.extracted = {
          ...state.extracted,
          source: String(data.mode || "").startsWith("ai") ? "ai" : state.extracted.source,
          items,
          total,
        };
      }
    } catch (e) {
      console.warn("assisted extract failed:", e);
    }
  }

  if (!state.extracted.items?.length) {
    if (state.extracted.source === "xlsx") {
      throw new Error(
        "엑셀에서 품목 헤더(내용/수량/단가/금액)를 찾지 못했습니다. 제공된 업로드 양식을 사용하거나 표 헤더를 포함해 주세요."
      );
    }
    if (state.extracted.source === "pdf") {
      throw new Error(
        "PDF에서 품목을 찾지 못했습니다. 스캔본이거나 품목표가 이미지/표로만 존재할 수 있습니다. (이미지 기반 추출이 필요)"
      );
    }
  }

  return {
    filename: f.name,
    source: state.extracted.source,
    items: Array.isArray(state.extracted.items) ? [...state.extracted.items] : [],
    total: state.extracted.total,
    rawText: state.extracted.rawText,
    rawRows: state.extracted.rawRows,
  };
}

async function onExtract() {
  const files = Array.from(els.file.files || []);
  if (!files.length) {
    setStatus("파일을 선택하세요.");
    return;
  }

  setBusy(true);

  try {
    const ok = [];
    const failed = [];

    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      setStatus(`처리 중... (${i + 1}/${files.length}) ${f.name}`);
      try {
        ok.push(await extractOneFile(f));
      } catch (e) {
        console.warn("extract failed:", f?.name, e);
        failed.push({ name: f?.name || "unknown", err: e });
      }
    }

    if (!ok.length) {
      const msg = failed.length
        ? `실패: ${failed[0]?.err?.message || String(failed[0]?.err || "")}`
        : "실패: 추출할 파일이 없습니다.";
      setStatus(msg);
      return;
    }

    if (ok.length === 1) {
      // Preserve previous behavior: render a single file's extracted state.
      state.extracted = {
        source: ok[0].source,
        items: ok[0].items,
        total: ok[0].total,
        rawText: ok[0].rawText,
        rawRows: ok[0].rawRows,
      };
    } else {
      const items = ok.flatMap((r) => r.items || []);
      const computed = computeTotal(items);
      const summed = ok
        .map((r) => r.total)
        .filter((n) => Number.isFinite(n))
        .reduce((a, b) => a + b, 0);
      // Prefer summing per-file totals (which may include shipping/discount) over recomputing from merged items.
      const total = (summed > 0 ? summed : null) ?? computed;
      const rawText = ok
        .map((r) => (r.rawText ? `----- ${r.filename} -----\n${r.rawText}` : ""))
        .filter(Boolean)
        .join("\n\n");
      state.extracted = { source: "multi", items, total, rawText, rawRows: null };
    }

    renderItems(state.extracted.items);
    renderExtractSummary();
    if (!failed.length) {
      setStatus(ok.length > 1 ? `추출이 완료되었습니다. (${ok.length}개 파일)` : "업로드/추출이 완료되었습니다.");
    } else {
      setStatus(`추출 완료 (${ok.length}개 성공, ${failed.length}개 실패)`);
    }
  } catch (e) {
    console.error(e);
    setStatus(`실패: ${e?.message || String(e)}`);
  } finally {
    setBusy(false);
  }
}

async function onCopyPurpose() {
  const ok = await copyText(els.purpose.value);
  setStatus(ok ? "개요가 복사되었습니다." : "복사에 실패했습니다.");
}

function setNavOpen(open) {
  const on = Boolean(open);
  document.body.classList.toggle("nav-open", on);
  els.navToggle.setAttribute("aria-expanded", on ? "true" : "false");
}

function init() {
  els.docDate.value = els.docDate.value || todayISO();
  setStatus("ready");

  els.btnExtract.addEventListener("click", onExtract);
  els.btnAutoOutline.addEventListener("click", onAutoOutline);
  els.btnCopyPurpose.addEventListener("click", onCopyPurpose);

  // 제목만 입력하면 개요 자동 생성 (개요가 비어 있을 때만)
  els.subject.addEventListener("blur", async () => {
    if (!els.subject.value.trim()) return;
    if (els.purpose.value.trim()) return;
    await onAutoOutline();
  });

  els.file.addEventListener("change", () => {
    const files = Array.from(els.file.files || []);
    if (!files.length) {
      els.fileHint.textContent = "지원: .xls, .xlsx, .pdf";
      return;
    }
    if (files.length === 1) {
      const f = files[0];
      const kb = Math.max(1, Math.round(f.size / 1024));
      els.fileHint.textContent = `${f.name} (${kb} KB)`;
    } else {
      const names = files
        .slice(0, 3)
        .map((f) => f.name)
        .join(", ");
      const more = files.length > 3 ? ` 외 ${files.length - 3}개` : "";
      els.fileHint.textContent = `${files.length}개 선택됨: ${names}${more}`;
    }
    // Auto extract on file selection.
    onExtract();
  });

  els.itemsTbody.addEventListener("click", async (e) => {
    const btn = e.target?.closest?.(".js-copy");
    if (!btn) return;
    const txt = btn.getAttribute("data-copy") || "";
    const ok = await copyText(txt);
    setStatus(ok ? "내용이 복사되었습니다." : "복사에 실패했습니다.");
  });

  els.navToggle.addEventListener("click", () => setNavOpen(true));
  els.navClose.addEventListener("click", () => setNavOpen(false));
  els.scrim.addEventListener("click", () => setNavOpen(false));
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") setNavOpen(false);
  });
}

init();
