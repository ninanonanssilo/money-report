const $ = (id) => document.getElementById(id);

const els = {
  sidebar: $("sidebar"),
  scrim: $("scrim"),
  navToggle: $("navToggle"),
  navClose: $("navClose"),

  file: $("file"),
  fileHint: $("file-hint"),

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
  },
};

function todayISO() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function setStatus(msg) {
  if (msg) console.log(msg);
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
  return (items || [])
    .map((it) => ({
      name: (it?.name ?? "").toString().trim(),
      spec: (it?.spec ?? it?.note ?? "").toString().trim(),
      qty: it?.qty === undefined || it?.qty === null || it?.qty === "" ? "" : Number(it.qty),
      unitPrice:
        it?.unitPrice === undefined || it?.unitPrice === null || it?.unitPrice === "" ? "" : Number(it.unitPrice),
      amount: it?.amount === undefined || it?.amount === null || it?.amount === "" ? "" : Number(it.amount),
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
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.async = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(s);
  });
}

async function loadPdfJs() {
  const url = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.mjs";
  const mod = await import(url);
  const workerUrl = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.mjs";
  mod.GlobalWorkerOptions.workerSrc = workerUrl;
  return mod;
}

async function extractFromXlsx(file) {
  if (!window.XLSX) {
    await loadScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
  }

  const ab = await file.arrayBuffer();
  const wb = window.XLSX.read(ab, { type: "array" });
  const firstName = wb.SheetNames[0];
  const ws = wb.Sheets[firstName];
  const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });

  const headerIdx = findHeaderRow(rows);
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
  const total = computeTotal(normalized);
  state.extracted = { source: "xlsx", items: normalized, total, rawText: "" };
}

function findHeaderRow(rows) {
  const keys = ["품목", "항목", "내용", "규격", "수량", "단가", "금액", "합계"];
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
  const spec = idxOf(by(["규격", "사양", "spec", "SPEC"]));
  const qty = idxOf(by(["수량", "수 량", "qty", "QTY"]));
  const unitPrice = idxOf(by(["단가", "단 가", "unit", "price", "단위금액", "단위 금액"]));
  const amount = idxOf(by(["금액", "공급가", "공급가액", "합계", "총액", "amount"]));

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
  const ab = await file.arrayBuffer();
  const pdfjs = await loadPdfJs();
  const loadingTask = pdfjs.getDocument({ data: ab });
  const doc = await loadingTask.promise;

  let all = "";
  const pages = Math.min(doc.numPages, 10);
  for (let p = 1; p <= pages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const text = content.items.map((it) => it.str).join(" ");
    all += text + "\n";
  }

  state.extracted = { source: "pdf", items: [], total: null, rawText: all.trim() };
}

function toDocPayload() {
  const items = normalizeItems(state.extracted.items);
  const total = computeTotal(items) ?? state.extracted.total;

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
    setStatus(data.mode === "ai" ? "AI로 개요가 생성되었습니다." : "개요가 생성되었습니다.");
  } catch (e) {
    console.error(e);
    // Fallback to deterministic outline
    els.purpose.value = buildOfficialOutline(payload);
    setStatus(`AI 연결 문제로 기본 개요로 채웠습니다: ${e?.message || String(e)}`);
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

async function onExtract() {
  const f = els.file.files?.[0];
  if (!f) {
    setStatus("파일을 선택하세요.");
    return;
  }

  setBusy(true);
  setStatus("처리 중...");

  try {
    const lower = f.name.toLowerCase();
    if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
      await extractFromXlsx(f);
      setStatus("엑셀 업로드/추출이 완료되었습니다.");
    } else if (f.type === "application/pdf" || f.name.toLowerCase().endsWith(".pdf")) {
      await extractFromPdf(f);
      setStatus("PDF 텍스트 추출이 완료되었습니다.");
    } else {
      setStatus("지원하지 않는 파일 형식입니다. (.xls, .xlsx, .pdf)");
      return;
    }

    renderItems(state.extracted.items);
    renderExtractSummary();
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
    const f = els.file.files?.[0];
    if (!f) {
      els.fileHint.textContent = "지원: .xls, .xlsx, .pdf";
      return;
    }
    const kb = Math.max(1, Math.round(f.size / 1024));
    els.fileHint.textContent = `${f.name} (${kb} KB)`;
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
