function jsonResponse(obj, { status = 200, headers = {} } = {}) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      ...headers,
    },
  });
}

function okCors() {
  return {
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "POST,OPTIONS",
    "access-control-allow-headers": "content-type,authorization",
    "access-control-max-age": "86400",
  };
}

function asString(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

function clamp(s, n) {
  const t = asString(s);
  return t.length <= n ? t : t.slice(0, n);
}

function toNum(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  const t = asString(v).replace(/[^\d.\-]/g, "");
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

function computeTotal(items) {
  const nums = (items || [])
    .map((it) => {
      const a = toNum(it?.amount);
      if (Number.isFinite(a)) return a;
      const q = toNum(it?.qty);
      const u = toNum(it?.unitPrice);
      if (Number.isFinite(q) && Number.isFinite(u)) return q * u;
      return 0;
    })
    .filter((n) => Number.isFinite(n));
  const t = nums.reduce((a, b) => a + b, 0);
  return t > 0 ? t : null;
}

function normalizeItems(items) {
  return (items || [])
    .map((it) => ({
      name: clamp(it?.name || "", 200),
      spec: clamp(it?.spec || it?.note || "", 200),
      qty: toNum(it?.qty),
      unitPrice: toNum(it?.unitPrice),
      amount: toNum(it?.amount),
      note: clamp(it?.note || "", 200),
    }))
    .filter((it) => it.name || Number.isFinite(it.amount) || Number.isFinite(it.unitPrice) || it.spec);
}

function findHeaderRow(rows) {
  const keys = [
    "품목",
    "항목",
    "내용",
    "규격",
    "옵션",
    "수량",
    "주문수량",
    "구매수량",
    "단가",
    "판매가",
    "금액",
    "합계",
    "합계금액",
    "주문금액",
    "결제금액",
  ];
  const maxScan = Math.min(rows.length, 60);
  for (let i = 0; i < maxScan; i++) {
    const row = rows[i].map((c) => asString(c || "").trim());
    const joined = row.join(" ");
    let score = 0;
    for (const k of keys) if (joined.includes(k)) score++;
    // Allow looser threshold for marketplace exports.
    if (score >= 2 && row.length >= 4) return i;
  }
  return -1;
}

function mapCols(header) {
  const h = header.map((c) => asString(c).trim());
  const idxOf = (pred) => {
    for (let i = 0; i < h.length; i++) if (pred(h[i])) return i;
    return -1;
  };
  const by = (arr) => (x) => arr.some((k) => x.includes(k));

  const name = idxOf(by(["품목", "항목", "내용", "제품", "서비스", "내역", "상품명", "상품", "상품정보"]));
  const spec = idxOf(by(["규격", "사양", "옵션", "모델", "모델명", "spec", "SPEC"]));
  const qty = idxOf(by(["수량", "수 량", "qty", "QTY", "주문수량", "구매수량"]));
  const unitPrice = idxOf(by(["단가", "단 가", "판매가", "판매 단가", "unit", "price", "단위금액", "단위 금액"]));
  const amount = idxOf(by(["금액", "합계", "합계금액", "주문금액", "결제금액", "총액", "총금액", "amount"]));
  const note = idxOf(by(["비고", "설명", "note"]));

  return { name, spec, qty, unitPrice, amount, note };
}

function pickCell(row, idx) {
  if (idx < 0) return "";
  const v = row?.[idx];
  return v === undefined || v === null ? "" : asString(v).trim();
}

function extractFromRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return { items: [], total: null };
  const headerIdx = findHeaderRow(rows);
  if (headerIdx < 0) return { items: [], total: null };

  const header = rows[headerIdx].map((c) => asString(c).trim());
  const col = mapCols(header);

  const items = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const name = pickCell(r, col.name);
    const spec = pickCell(r, col.spec);
    const qty = pickCell(r, col.qty);
    const unitPrice = pickCell(r, col.unitPrice);
    const amount = pickCell(r, col.amount);
    const note = pickCell(r, col.note);

    if (!name && !spec && !qty && !unitPrice && !amount) continue;
    items.push({ name, spec, qty, unitPrice, amount, note });
  }

  const normalized = normalizeItems(items);
  const total = computeTotal(normalized);
  return { items: normalized, total };
}

function buildPrompt({ source, filename, rows, rawText, pageImages }) {
  const safe = {
    source: clamp(source, 40),
    filename: clamp(filename, 120),
    // Truncate rows aggressively; model doesn't need everything.
    rows: Array.isArray(rows) ? rows.slice(0, 220).map((r) => (Array.isArray(r) ? r.slice(0, 18) : [])) : null,
    rawText: clamp(rawText, 9000),
    // Images are sent separately as multimodal inputs; keep only metadata here.
    pageImages: Array.isArray(pageImages) && pageImages.length ? { count: pageImages.length } : null,
  };

  const instructions = [
    "You extract line items from Korean marketplace quotation/order exports (XLS/XLSX/PDF text).",
    "If page images are provided (scanned PDF / image-only tables), use them as the primary source of truth.",
    "Return ONLY valid JSON with this schema:",
    "{",
    '  "items": [{"name": string, "spec": string, "qty": number|null, "unitPrice": number|null, "amount": number|null, "note": string}],',
    '  "total": number|null',
    "}",
    "Rules:",
    "- Do not invent numbers. If missing, use null.",
    "- Prefer amount; else compute amount=qty*unitPrice when both exist.",
    "- Remove shipping rows / summary rows from items when possible.",
    "- Keep items <= 80.",
    "- Currency is KRW unless explicitly stated otherwise.",
  ].join("\n");

  return { instructions, safe };
}

function normalizePageImages(v) {
  if (!Array.isArray(v)) return [];
  const out = [];
  for (const x of v) {
    if (typeof x !== "string") continue;
    const s = x.trim();
    if (!s.startsWith("data:image/")) continue;
    // Guardrail: avoid huge payloads / abuse.
    if (s.length > 2500000) continue;
    out.push(s);
    if (out.length >= 3) break;
  }
  return out;
}

async function callOpenAI({ apiKey, model, source, filename, rows, rawText, pageImages }) {
  const images = normalizePageImages(pageImages);
  const { instructions, safe } = buildPrompt({ source, filename, rows, rawText, pageImages: images });

  const useModel = (asString(model).trim() || "gpt-5.2").trim();

  const schema = {
    type: "object",
    additionalProperties: false,
    properties: {
      items: {
        type: "array",
        items: {
          type: "object",
          additionalProperties: false,
          properties: {
            name: { type: "string" },
            spec: { type: "string" },
            qty: { anyOf: [{ type: "number" }, { type: "null" }] },
            unitPrice: { anyOf: [{ type: "number" }, { type: "null" }] },
            amount: { anyOf: [{ type: "number" }, { type: "null" }] },
            note: { type: "string" },
          },
          required: ["name", "spec", "qty", "unitPrice", "amount", "note"],
        },
      },
      total: { anyOf: [{ type: "number" }, { type: "null" }] },
    },
    required: ["items", "total"],
  };

  const res = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: useModel,
      instructions,
      input: [
        {
          role: "user",
          content: [
            { type: "input_text", text: JSON.stringify(safe) },
            ...images.map((dataUrl) => ({ type: "input_image", image_url: dataUrl })),
          ],
        },
      ],
      text: {
        format: {
          type: "json_schema",
          name: "extract_items",
          description: "Extract estimate line items and totals as strict JSON.",
          schema,
          strict: true,
        },
      },
    }),
  });

  const txt = await res.text();
  let data;
  try {
    data = JSON.parse(txt);
  } catch {
    throw new Error(`OpenAI 응답 파싱 실패: ${txt.slice(0, 180)}`);
  }
  if (!res.ok) {
    const msg = data?.error?.message || `OpenAI HTTP ${res.status}`;
    throw new Error(msg);
  }

  const content =
    data?.output_text ||
    (Array.isArray(data?.output)
      ? data.output
          .flatMap((o) => (o?.content || []).filter((c) => c?.type === "output_text").map((c) => c.text))
          .filter(Boolean)
          .join("\n")
      : "");
  if (!content) throw new Error("OpenAI 응답 content가 비어 있습니다.");

  let out;
  try {
    out = JSON.parse(content);
  } catch {
    throw new Error("OpenAI가 JSON이 아닌 내용을 반환했습니다.");
  }

  const items = normalizeItems(Array.isArray(out?.items) ? out.items.slice(0, 80) : []);
  const total = Number.isFinite(out?.total) ? out.total : computeTotal(items);
  return { items, total };
}

export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: okCors() });
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.json();
    const source = body?.source || "";
    const filename = body?.filename || "";
    const rows = body?.rows || null;
    const rawText = body?.rawText || "";
    const pageImages = body?.pageImages || null;
    const rev = context.env?.CF_PAGES_COMMIT_SHA || null;

    // 1) Deterministic extraction first (works without AI key).
    const deterministic = extractFromRows(rows || []);
    if (deterministic.items.length >= 1) {
      return jsonResponse({ mode: "heuristic", items: deterministic.items, total: deterministic.total, rev }, { headers: okCors() });
    }

    // 2) AI-assisted extraction if configured.
    const apiKey = context.env?.OPENAI_API_KEY;
    if (!apiKey) {
      return jsonResponse({ mode: "none", items: [], total: null, rev }, { headers: okCors() });
    }

    const model = context.env?.OPENAI_MODEL;
    const ai = await callOpenAI({ apiKey, model, source, filename, rows, rawText, pageImages });
    return jsonResponse({ mode: "ai", items: ai.items, total: ai.total, rev }, { headers: okCors() });
  } catch (e) {
    const msg = e?.message ? String(e.message) : "unknown error";
    const rev = context.env?.CF_PAGES_COMMIT_SHA || null;
    return jsonResponse({ error: msg, rev }, { status: 500, headers: okCors() });
  }
}
