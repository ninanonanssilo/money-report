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

function fmtMoney(v) {
  const n = typeof v === "number" ? v : Number(asString(v).replace(/,/g, ""));
  if (!Number.isFinite(n)) return asString(v);
  return new Intl.NumberFormat("ko-KR").format(n);
}

function computeTotal(items) {
  const nums = (items || [])
    .map((it) => {
      if (Number.isFinite(it?.amount)) return it.amount;
      if (Number.isFinite(it?.qty) && Number.isFinite(it?.unitPrice)) return it.qty * it.unitPrice;
      return 0;
    })
    .filter((n) => Number.isFinite(n));
  const t = nums.reduce((a, b) => a + b, 0);
  return t > 0 ? t : null;
}

function buildTemplateDoc(payload) {
  const meta = payload?.meta || {};
  const quote = payload?.quote || {};
  const itemsIn = Array.isArray(quote.items) ? quote.items : [];
  const totalN = Number.isFinite(quote.total) ? quote.total : computeTotal(itemsIn);

  const items = itemsIn.slice(0, 80).map((it) => ({
    name: clamp(it?.name || "", 160),
    qty: it?.qty === "" || it?.qty === null || it?.qty === undefined ? "" : fmtMoney(it.qty),
    unitPrice: it?.unitPrice === "" || it?.unitPrice === null || it?.unitPrice === undefined ? "" : fmtMoney(it.unitPrice),
    amount:
      it?.amount === "" || it?.amount === null || it?.amount === undefined
        ? ""
        : fmtMoney(it.amount),
    note: clamp(it?.note || "", 200),
  }));

  return {
    subject: clamp(meta.subject || "", 120),
    requester: clamp(meta.requester || "", 40),
    docDate: clamp(meta.docDate || "", 20),
    purpose: clamp(meta.purpose || "", 2000),
    notes: clamp(meta.notes || "-", 1200) || "-",
    approval:
      "상기 목적 달성을 위해 견적 내역과 같이 구매/결제를 진행하고자 하오니 검토 후 결재를 요청드립니다.\n" +
      "집행 기준 및 예산 범위 내에서 진행 예정입니다.",
    items,
    total: totalN ? `${fmtMoney(totalN)} 원` : "",
  };
}

function buildPrompt(payload) {
  const meta = payload?.meta || {};
  const quote = payload?.quote || {};
  const items = Array.isArray(quote.items) ? quote.items : [];

  const safe = {
    meta: {
      subject: clamp(meta.subject, 120),
      requester: clamp(meta.requester, 40),
      docDate: clamp(meta.docDate, 20),
      purpose: clamp(meta.purpose, 2000),
      notes: clamp(meta.notes, 1200),
    },
    quote: {
      vendor: clamp(quote.vendor, 120),
      currency: "KRW",
      total: Number.isFinite(quote.total) ? quote.total : null,
      items: items.slice(0, 80).map((it) => ({
        name: clamp(it?.name, 160),
        qty: Number.isFinite(it?.qty) ? it.qty : null,
        unitPrice: Number.isFinite(it?.unitPrice) ? it.unitPrice : null,
        amount: Number.isFinite(it?.amount) ? it.amount : null,
        note: clamp(it?.note, 200),
      })),
      rawText: clamp(quote.rawText, 8000),
    },
  };

  const instructions = [
    "You are drafting a Korean official-style approval request document (품의서) from quotation info.",
    "Return ONLY valid JSON matching this schema:",
    "{",
    '  "subject": string,',
    '  "requester": string,',
    '  "docDate": string,',
    '  "purpose": string,',
    '  "approval": string,',
    '  "notes": string,',
    '  "items": [{"name":string,"qty":string,"unitPrice":string,"amount":string,"note":string}],',
    '  "total": string',
    "}",
    "Rules:",
    "- Use polite, formal Korean (공문서/품의서 톤).",
    "- Do not invent vendor or numbers not supported; if unknown, leave blank or write '미기재'.",
    "- Ensure totals align with items when possible.",
    "- Keep it concise but complete for approval.",
  ].join("\n");

  return { instructions, safe };
}

async function callOpenAI({ apiKey, payload }) {
  const { instructions, safe } = buildPrompt(payload);

  const res = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      temperature: 0.2,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: instructions },
        { role: "user", content: JSON.stringify(safe) },
      ],
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

  const content = data?.choices?.[0]?.message?.content;
  if (!content) throw new Error("OpenAI 응답 content가 비어 있습니다.");

  let doc;
  try {
    doc = JSON.parse(content);
  } catch {
    throw new Error("OpenAI가 JSON이 아닌 내용을 반환했습니다.");
  }

  // Very light validation/sanitization.
  doc.subject = clamp(doc.subject, 120);
  doc.requester = clamp(doc.requester, 40);
  doc.docDate = clamp(doc.docDate, 20);
  doc.purpose = clamp(doc.purpose, 3000);
  doc.approval = clamp(doc.approval, 3000);
  doc.notes = clamp(doc.notes, 2000);
  doc.total = clamp(doc.total, 40);
  doc.items = Array.isArray(doc.items)
    ? doc.items.slice(0, 80).map((it) => ({
        name: clamp(it?.name, 160),
        qty: clamp(it?.qty, 40),
        unitPrice: clamp(it?.unitPrice, 40),
        amount: clamp(it?.amount, 40),
        note: clamp(it?.note, 200),
      }))
    : [];

  return doc;
}

export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: okCors() });
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.json();
    const payload = body?.payload;
    if (!payload) return jsonResponse({ error: "payload is required" }, { status: 400, headers: okCors() });

    const envKey = context.env?.OPENAI_API_KEY;
    const apiKey = envKey;

    if (!apiKey) {
      const document = buildTemplateDoc(payload);
      return jsonResponse({ mode: "template", document }, { headers: okCors() });
    }

    const document = await callOpenAI({ apiKey, payload });
    return jsonResponse({ mode: "ai", document }, { headers: okCors() });
  } catch (e) {
    // Avoid leaking secrets; return generic error.
    const msg = e?.message ? String(e.message) : "unknown error";
    return jsonResponse({ error: msg }, { status: 500, headers: okCors() });
  }
}
