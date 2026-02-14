const $ = (id) => document.getElementById(id);

const els = {
  file: $("file"),
  fileHint: $("file-hint"),

  subject: $("subject"),
  department: $("department"),
  requester: $("requester"),
  docDate: $("docDate"),
  purpose: $("purpose"),
  notes: $("notes"),
  quoteText: $("quoteText"),

  btnExtract: $("btn-extract"),
  btnGenerate: $("btn-generate"),
  btnPrint: $("btn-print"),
  btnCopy: $("btn-copy"),

  status: $("status"),
  extractSummary: $("extractSummary"),
  itemsTbody: $("itemsTbody"),

  doc: $("doc"),
  docTemplate: $("docTemplate"),
};

// Minimal app state; never persist API keys.
const state = {
  extracted: {
    source: null,
    vendor: null,
    currency: "KRW",
    items: [],
    total: null,
    rawText: "",
  },
};

function setStatus(msg) {
  els.status.textContent = msg || "";
}

function fmtMoney(v) {
  if (v === null || v === undefined || v === "") return "";
  const n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
  if (!Number.isFinite(n)) return String(v);
  return new Intl.NumberFormat("ko-KR").format(n);
}

function todayISO() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function normalizeItems(items) {
  return (items || [])
    .map((it) => ({
      name: (it?.name ?? "").toString().trim(),
      qty: it?.qty === undefined || it?.qty === null || it?.qty === "" ? "" : Number(it.qty),
      unitPrice:
        it?.unitPrice === undefined || it?.unitPrice === null || it?.unitPrice === "" ? "" : Number(it.unitPrice),
      amount: it?.amount === undefined || it?.amount === null || it?.amount === "" ? "" : Number(it.amount),
      note: (it?.note ?? "").toString().trim(),
    }))
    .filter((it) => it.name || it.amount || it.unitPrice);
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

function renderItems(items) {
  els.itemsTbody.innerHTML = "";
  for (const it of items) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(it.name || "")}</td>
      <td class="num">${escapeHtml(it.qty === "" ? "" : fmtMoney(it.qty))}</td>
      <td class="num">${escapeHtml(it.unitPrice === "" ? "" : fmtMoney(it.unitPrice))}</td>
      <td class="num">${escapeHtml(it.amount === "" ? "" : fmtMoney(it.amount))}</td>
      <td>${escapeHtml(it.note || "")}</td>
    `;
    els.itemsTbody.appendChild(tr);
  }
}

function renderExtractSummary() {
  const { source, vendor, items, total } = state.extracted;
  const parts = [];
  if (source) parts.push(`source: ${source}`);
  if (vendor) parts.push(`vendor: ${vendor}`);
  parts.push(`items: ${items.length}`);
  if (total) parts.push(`total: ${fmtMoney(total)} KRW`);
  els.extractSummary.textContent = parts.join(" | ");
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

async function extractFromXlsx(file) {
  // Load SheetJS via CDN only when needed.
  if (!window.XLSX) {
    await loadScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
  }

  const ab = await file.arrayBuffer();
  const wb = window.XLSX.read(ab, { type: "array" });
  const firstName = wb.SheetNames[0];
  const ws = wb.Sheets[firstName];
  const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });

  // Heuristic extraction: find header row containing "품목" or "항목" etc.
  const headerIdx = findHeaderRow(rows);
  const items = [];

  if (headerIdx >= 0) {
    const header = rows[headerIdx].map((c) => String(c).trim());
    const col = mapCols(header);

    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = pickCell(r, col.name);
      const qty = parseNum(pickCell(r, col.qty));
      const unitPrice = parseNum(pickCell(r, col.unitPrice));
      const amount = parseNum(pickCell(r, col.amount));
      const note = pickCell(r, col.note);

      const it = { name, qty, unitPrice, amount, note };
      if (!it.name && !Number.isFinite(it.amount) && !Number.isFinite(it.unitPrice)) continue;
      items.push(it);
    }
  }

  const normalized = normalizeItems(items);
  const total = computeTotal(normalized);

  state.extracted = {
    source: "xlsx",
    vendor: null,
    currency: "KRW",
    items: normalized,
    total,
    rawText: "",
  };
}

function findHeaderRow(rows) {
  const keys = ["품목", "항목", "내용", "제품", "서비스", "수량", "단가", "금액", "합계"];
  const maxScan = Math.min(rows.length, 50);
  for (let i = 0; i < maxScan; i++) {
    const row = rows[i].map((c) => String(c || "").trim());
    const joined = row.join(" ");
    let score = 0;
    for (const k of keys) if (joined.includes(k)) score++;
    if (score >= 3) return i;
  }
  return -1;
}

function mapCols(header) {
  const idxOf = (pred) => {
    for (let i = 0; i < header.length; i++) if (pred(header[i])) return i;
    return -1;
  };
  const by = (arr) => (h) => arr.some((k) => h.includes(k));

  const name = idxOf(by(["품목", "항목", "내용", "제품", "서비스", "내역"]));
  const qty = idxOf(by(["수량", "수 량", "qty", "QTY"]));
  const unitPrice = idxOf(by(["단가", "단 가", "unit", "price", "단위금액", "단위 금액"]));
  const amount = idxOf(by(["금액", "공급가", "공급가액", "합계", "총액", "amount"]));
  const note = idxOf(by(["비고", "설명", "note"]));

  return { name, qty, unitPrice, amount, note };
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
  // Uses PDF.js via CDN; best-effort text extraction.
  const ab = await file.arrayBuffer();

  const pdfjs = await loadPdfJs();
  const loadingTask = pdfjs.getDocument({ data: ab });
  const doc = await loadingTask.promise;

  let all = "";
  const pages = Math.min(doc.numPages, 10); // MVP: cap pages
  for (let p = 1; p <= pages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const text = content.items.map((it) => it.str).join(" ");
    all += text + "\n";
  }

  const rawText = all.trim();
  state.extracted = {
    source: "pdf",
    vendor: null,
    currency: "KRW",
    items: [],
    total: null,
    rawText,
  };
}

async function loadPdfJs() {
  // pdfjs-dist ESM build from cdnjs; loaded via dynamic import.
  const url = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.mjs";
  const mod = await import(url);
  // Worker
  const workerUrl = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.mjs";
  mod.GlobalWorkerOptions.workerSrc = workerUrl;
  return mod;
}

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(s);
  });
}

function toDocPayload() {
  const items = normalizeItems(state.extracted.items);
  const total = computeTotal(items) ?? state.extracted.total;

  return {
    meta: {
      subject: els.subject.value.trim(),
      department: els.department.value.trim(),
      requester: els.requester.value.trim(),
      docDate: (els.docDate.value || "").trim(),
      purpose: els.purpose.value.trim(),
      notes: els.notes.value.trim(),
    },
    quote: {
      source: state.extracted.source,
      vendor: state.extracted.vendor,
      currency: "KRW",
      items,
      total,
      rawText: [state.extracted.rawText, els.quoteText.value.trim()].filter(Boolean).join("\n\n"),
    },
  };
}

function validateForGenerate(payload) {
  const errs = [];
  if (!payload.meta.subject) errs.push("제목을 입력하세요.");
  if (!payload.meta.department) errs.push("부서를 입력하세요.");
  if (!payload.meta.requester) errs.push("기안자를 입력하세요.");
  if (!payload.meta.docDate) errs.push("작성일을 입력하세요.");
  if (!payload.meta.purpose) errs.push("목적/배경을 입력하세요.");

  if ((!payload.quote.items || payload.quote.items.length === 0) && !payload.quote.rawText) {
    errs.push("견적 내역이 비어 있습니다. 추출을 하거나 견적서 텍스트를 붙여넣으세요.");
  }
  return errs;
}

function renderDocFromTemplate(docData) {
  const frag = els.docTemplate.content.cloneNode(true);
  const root = frag.querySelector(".doc-page");

  const bindText = (key, val) => {
    const el = root.querySelector(`[data-bind="${key}"]`);
    if (el) el.textContent = val ?? "";
  };

  bindText("docDate", docData.docDate || "");
  bindText("department", docData.department || "");
  bindText("requester", docData.requester || "");
  bindText("subject", docData.subject || "");
  bindText("purpose", docData.purpose || "");
  bindText("approval", docData.approval || "");
  bindText("notes", docData.notes || "");
  bindText("total", docData.total || "");

  const tbody = root.querySelector(`[data-bind="items"]`);
  tbody.innerHTML = "";
  for (const it of docData.items || []) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(it.name || "")}</td>
      <td class="num">${escapeHtml(it.qty || "")}</td>
      <td class="num">${escapeHtml(it.unitPrice || "")}</td>
      <td class="num">${escapeHtml(it.amount || "")}</td>
      <td>${escapeHtml(it.note || "")}</td>
    `;
    tbody.appendChild(tr);
  }

  els.doc.innerHTML = "";
  els.doc.appendChild(frag);
}

function buildTemplateDoc(payload) {
  const items = (payload.quote.items || []).map((it) => ({
    name: it.name || "",
    qty: it.qty === "" ? "" : fmtMoney(it.qty),
    unitPrice: it.unitPrice === "" ? "" : fmtMoney(it.unitPrice),
    amount:
      it.amount === "" ? "" : fmtMoney(it.amount ?? (Number.isFinite(it.qty) && Number.isFinite(it.unitPrice) ? it.qty * it.unitPrice : "")),
    note: it.note || "",
  }));

  const total = payload.quote.total ? fmtMoney(payload.quote.total) : fmtMoney(computeTotal(payload.quote.items || []) || "");

  return {
    subject: payload.meta.subject,
    department: payload.meta.department,
    requester: payload.meta.requester,
    docDate: payload.meta.docDate,
    purpose: payload.meta.purpose,
    notes: payload.meta.notes || "-",
    approval:
      `상기 목적 달성을 위해 견적 내역과 같이 구매/결제를 진행하고자 하오니 검토 후 결재를 요청드립니다.\n` +
      `집행 기준 및 예산 범위 내에서 진행 예정입니다.`,
    items,
    total: total ? `${total} 원` : "",
  };
}

async function callGenerateApi(payload, apiKey) {
  const res = await fetch("/api/generate", {
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

async function onExtract() {
  setStatus("추출 중...");

  const f = els.file.files?.[0];
  if (!f) {
    setStatus("파일을 선택하세요.");
    return;
  }

  try {
    if (f.name.toLowerCase().endsWith(".xlsx")) {
      await extractFromXlsx(f);
      setStatus("엑셀에서 항목을 추출했습니다.");
    } else if (f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf")) {
      await extractFromPdf(f);
      setStatus("PDF 텍스트를 추출했습니다. (형식에 따라 품질이 달라질 수 있어요)");
      if (state.extracted.rawText) {
        els.quoteText.value = state.extracted.rawText;
      }
    } else {
      setStatus("지원하지 않는 파일 형식입니다. (.xlsx, .pdf)");
      return;
    }

    renderItems(state.extracted.items);
    renderExtractSummary();
  } catch (e) {
    console.error(e);
    setStatus(`추출 실패: ${e?.message || String(e)}`);
  }
}

async function onGenerate() {
  const payload = toDocPayload();
  const errs = validateForGenerate(payload);
  if (errs.length) {
    setStatus(errs.join(" "));
    return;
  }

  setStatus("품의서 생성 중...");

  try {
    const data = await callGenerateApi(payload);
    renderDocFromTemplate(data.document);
    setStatus(data.mode === "ai" ? "AI로 품의서를 생성했습니다." : "템플릿으로 품의서를 생성했습니다.");
  } catch (e) {
    console.error(e);
    // Fallback to deterministic template on any failure.
    const doc = buildTemplateDoc(payload);
    renderDocFromTemplate(doc);
    setStatus(`API 실패로 템플릿으로 생성했습니다: ${e?.message || String(e)}`);
  }
}

async function onCopy() {
  const text = els.doc.innerText.trim();
  if (!text) {
    setStatus("복사할 문서가 없습니다.");
    return;
  }
  await navigator.clipboard.writeText(text);
  setStatus("미리보기 텍스트를 클립보드에 복사했습니다.");
}

function onPrint() {
  window.print();
}

function initDefaults() {
  if (!els.docDate.value) els.docDate.value = todayISO();
}

els.btnExtract.addEventListener("click", onExtract);
els.btnGenerate.addEventListener("click", onGenerate);
els.btnCopy.addEventListener("click", onCopy);
els.btnPrint.addEventListener("click", onPrint);

initDefaults();
setStatus("파일을 올리고 '파일에서 항목 추출' 또는 바로 '품의서 생성'을 누르세요.");
