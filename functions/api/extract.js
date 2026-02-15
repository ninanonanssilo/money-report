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

function computeItemsSubtotal(items) {
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

function splitAdjustments(items) {
  const out = [];
  let shipping = 0;
  let discount = 0;

  for (const it of items || []) {
    const name = asString(it?.name).replace(/\s+/g, " ").trim();
    const amt = toNum(it?.amount);
    const isSummary = /^(합계|총계|총\s*합계|총\s*금액|결제\s*금액)$/i.test(name);

    if (isSummary) continue;

    if (Number.isFinite(amt) && amt >= 0) {
      if (/배송/.test(name) || /선결제/.test(name)) {
        shipping += amt;
        continue;
      }
      if (/할인|쿠폰|프로모션/.test(name)) {
        discount += amt;
        continue;
      }
    }
    out.push(it);
  }

  return {
    items: out,
    shipping: shipping > 0 ? shipping : null,
    discount: discount > 0 ? discount : null,
  };
}

function computeTotals({ items, shipping, discount }) {
  const itemsSubtotal = computeItemsSubtotal(items) ?? 0;
  const ship = Number.isFinite(toNum(shipping)) ? toNum(shipping) : 0;
  const disc = Number.isFinite(toNum(discount)) ? toNum(discount) : 0;
  const grandTotal = itemsSubtotal + ship - disc;
  return {
    itemsSubtotal: itemsSubtotal > 0 ? itemsSubtotal : null,
    shipping: ship > 0 ? ship : null,
    discount: disc > 0 ? disc : null,
    grandTotal: grandTotal > 0 ? grandTotal : null,
  };
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
  if (!Array.isArray(rows) || rows.length === 0) return { items: [], totals: computeTotals({ items: [], shipping: null, discount: null }) };
  const headerIdx = findHeaderRow(rows);
  if (headerIdx < 0) return { items: [], totals: computeTotals({ items: [], shipping: null, discount: null }) };

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
  const split = splitAdjustments(normalized);
  const totals = computeTotals({ items: split.items, shipping: split.shipping, discount: split.discount });
  return { items: split.items, totals: totals };
}

function buildPrompt({ source, filename, rows, rawText, pageImages, initialItems }) {
  const safe = {
    source: clamp(source, 40),
    filename: clamp(filename, 120),
    // Truncate rows aggressively; model doesn't need everything.
    rows: Array.isArray(rows) ? rows.slice(0, 220).map((r) => (Array.isArray(r) ? r.slice(0, 18) : [])) : null,
    rawText: clamp(rawText, 9000),
    // Images are sent separately as multimodal inputs; keep only metadata here.
    pageImages: Array.isArray(pageImages) && pageImages.length ? { count: pageImages.length } : null,
    // If heuristic extraction ran, provide it as a hint for correction/normalization.
    initialItems: Array.isArray(initialItems) && initialItems.length ? initialItems.slice(0, 60) : null,
  };

  const instructions = [
    "You extract line items from Korean marketplace quotation/order exports (XLS/XLSX/PDF text).",
    "If page images are provided (scanned PDF / image-only tables), use them as the primary source of truth.",
    "If initialItems are provided, treat them as a hint but correct any missing/wrong qty/unitPrice/amount.",
    "Return ONLY valid JSON with this schema:",
    "{",
    '  "items": [{"name": string, "spec": string, "qty": number|null, "unitPrice": number|null, "amount": number|null, "note": string}],',
    '  "shipping": number|null,',
    '  "discount": number|null,',
    '  "statedTotal": number|null',
    "}",
    "Rules:",
    "- Do not invent numbers. If missing, use null.",
    "- Prefer amount; else compute amount=qty*unitPrice when both exist.",
    "- Do NOT include shipping/discount/summary lines in items; put them into shipping/discount/statedTotal.",
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

function shouldRefineWithAI({ items, total }) {
  const arr = Array.isArray(items) ? items : [];
  if (!arr.length) return true;
  const missAmount = arr.filter((it) => !(Number.isFinite(it?.amount) && it.amount >= 0)).length;
  const missAll = arr.filter((it) => {
    const hasAmt = Number.isFinite(it?.amount) && it.amount >= 0;
    const hasCalc = Number.isFinite(it?.qty) && Number.isFinite(it?.unitPrice);
    return !hasAmt && !hasCalc;
  }).length;

  // If total missing, or lots of missing amounts, ask the model to refine.
  if (!(Number.isFinite(total) && total > 0)) return true;
  if (missAll / arr.length > 0.25) return true;
  if (missAmount / arr.length > 0.55) return true;
  return false;
}

async function callOpenAI({ apiKey, model, source, filename, rows, rawText, pageImages, initialItems }) {
  const images = normalizePageImages(pageImages);
  const { instructions, safe } = buildPrompt({ source, filename, rows, rawText, pageImages: images, initialItems });

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
      shipping: { anyOf: [{ type: "number" }, { type: "null" }] },
      discount: { anyOf: [{ type: "number" }, { type: "null" }] },
      statedTotal: { anyOf: [{ type: "number" }, { type: "null" }] },
    },
    required: ["items", "shipping", "discount", "statedTotal"],
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
  const split = splitAdjustments(items);
  const shipping = Number.isFinite(out?.shipping) ? out.shipping : split.shipping;
  const discount = Number.isFinite(out?.discount) ? out.discount : split.discount;
  const statedTotal = Number.isFinite(out?.statedTotal) ? out.statedTotal : null;
  const totals = computeTotals({ items: split.items, shipping, discount });
  return { items: split.items, shipping: totals.shipping, discount: totals.discount, statedTotal, totals };
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
    const detTotals = deterministic?.totals || computeTotals({ items: deterministic.items || [], shipping: null, discount: null });

    // 2) AI-assisted extraction if configured.
    const apiKey = context.env?.OPENAI_API_KEY;
    if (!apiKey) {
      if (deterministic.items.length >= 1) {
        return jsonResponse(
          {
            mode: "heuristic",
            items: deterministic.items,
            shipping: detTotals.shipping,
            discount: detTotals.discount,
            statedTotal: null,
            totals: detTotals,
            total: detTotals.grandTotal,
            rev,
          },
          { headers: okCors() }
        );
      }
      return jsonResponse({ mode: "none", items: [], total: null, rev }, { headers: okCors() });
    }

    const model = context.env?.OPENAI_MODEL;
    if (deterministic.items.length >= 1) {
      if (!shouldRefineWithAI({ items: deterministic.items, total: detTotals.grandTotal })) {
        return jsonResponse(
          {
            mode: "heuristic",
            items: deterministic.items,
            shipping: detTotals.shipping,
            discount: detTotals.discount,
            statedTotal: null,
            totals: detTotals,
            total: detTotals.grandTotal,
            rev,
          },
          { headers: okCors() }
        );
      }
      const ai = await callOpenAI({
        apiKey,
        model,
        source,
        filename,
        rows,
        rawText,
        pageImages,
        initialItems: deterministic.items,
      });
      return jsonResponse(
        { mode: "ai_refine", items: ai.items, shipping: ai.shipping, discount: ai.discount, statedTotal: ai.statedTotal, totals: ai.totals, total: ai.totals?.grandTotal ?? null, rev },
        { headers: okCors() }
      );
    }

    const ai = await callOpenAI({ apiKey, model, source, filename, rows, rawText, pageImages, initialItems: null });
    return jsonResponse(
      { mode: "ai", items: ai.items, shipping: ai.shipping, discount: ai.discount, statedTotal: ai.statedTotal, totals: ai.totals, total: ai.totals?.grandTotal ?? null, rev },
      { headers: okCors() }
    );
  } catch (e) {
    const msg = e?.message ? String(e.message) : "unknown error";
    const rev = context.env?.CF_PAGES_COMMIT_SHA || null;
    return jsonResponse({ error: msg, rev }, { status: 500, headers: okCors() });
  }
}
