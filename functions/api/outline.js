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

function buildOutlineTemplate(payload) {
  const meta = payload?.meta || {};
  const quote = payload?.quote || {};
  const items = Array.isArray(quote.items) ? quote.items : [];
  const total = Number.isFinite(quote.total) ? quote.total : computeTotal(items);

  const itemNames = items
    .map((it) => asString(it?.name).trim())
    .filter(Boolean)
    .slice(0, 6);
  const itemLine = itemNames.length ? itemNames.join(", ") + (items.length > itemNames.length ? " 등" : "") : "미기재";

  const subject = asString(meta.subject).trim() || "미기재";
  const purposeLine = asString(meta.purpose).trim()
    ? asString(meta.purpose).trim().replace(/\s+/g, " ").slice(0, 160)
    : "미기재";
  const totalN = Number.isFinite(total) && total >= 0 ? total : 0;
  const totalLine = `금${fmtMoney(totalN)}원(${numToKoreanWon(totalN)})`;

  const basis = items
    .filter((it) => it && (it.amount || (Number.isFinite(it.qty) && Number.isFinite(it.unitPrice))))
    .slice(0, 4)
    .map((it) => {
      const qty = Number.isFinite(it.qty) ? it.qty : null;
      const unit = Number.isFinite(it.unitPrice) ? it.unitPrice : null;
      const amt = Number.isFinite(it.amount) ? it.amount : qty !== null && unit !== null ? qty * unit : null;
      if (qty !== null && unit !== null && amt !== null) return `${fmtMoney(unit)}원 X ${fmtMoney(qty)}개 = ${fmtMoney(amt)}원`;
      if (amt !== null) return `${fmtMoney(amt)}원`;
      return "";
    })
    .filter(Boolean)
    .join(" / ");

  const basisLine = basis || (total ? `${fmtMoney(total)}원` : "미기재");

  return (
    `${subject}을(를) 다음과 같이 구입하고자 합니다.\n` +
    `1. 목적: ${purposeLine}\n` +
    `2. 품명: ${itemLine}\n` +
    `3. 소요 예산: ${totalLine}\n` +
    `4. 산출 근거: ${basisLine}\n\n` +
    "붙임 지출품의서 1부. 끝."
  );
}

function buildPrompt(payload) {
  const meta = payload?.meta || {};
  const quote = payload?.quote || {};
  const items = Array.isArray(quote.items) ? quote.items : [];

  const safe = {
    meta: {
      subject: clamp(meta.subject, 120),
      purposeHint: clamp(meta.purpose, 500), // optional hint; can be empty
    },
    quote: {
      currency: "KRW",
      total: Number.isFinite(quote.total) ? quote.total : null,
      items: items.slice(0, 60).map((it) => ({
        name: clamp(it?.name, 160),
        qty: Number.isFinite(it?.qty) ? it.qty : null,
        unitPrice: Number.isFinite(it?.unitPrice) ? it.unitPrice : null,
        amount: Number.isFinite(it?.amount) ? it.amount : null,
        note: clamp(it?.note, 120),
      })),
      rawText: clamp(quote.rawText, 4000),
    },
  };

  const instructions = [
    "You draft ONLY the outline section (품의개요) for a Korean official expense approval document (지출 품의서).",
    "Return ONLY valid JSON: {\"outline\": string}.",
    "The outline MUST follow this exact shape (use newlines):",
    "1) '<제목>을(를) 다음과 같이 구입하고자 합니다.'",
    "2) '1. 목적: ...'",
    "3) '2. 품명: ...'",
    "4) '3. 소요 예산: 금0,000원(금영원)' (반드시 이 형태. 괄호 안은 한글 금액)",
    "5) '4. 산출 근거: ...'",
    "6) blank line",
    "7) '붙임 지출품의서 1부. 끝.'",
    "Rules:",
    "- Use Korean formal tone.",
    "- Do not invent numbers; use provided items/total only.",
    "- If unknown, write '미기재'.",
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

  let out;
  try {
    out = JSON.parse(content);
  } catch {
    throw new Error("OpenAI가 JSON이 아닌 내용을 반환했습니다.");
  }

  const outline = clamp(out?.outline || "", 6000).trim();
  if (!outline) throw new Error("outline이 비어 있습니다.");
  return outline;
}

function hasBatchim(s) {
  const t = String(s || "").trim();
  if (!t) return false;
  const ch = t.charCodeAt(t.length - 1);
  if (ch < 0xac00 || ch > 0xd7a3) return false;
  return (ch - 0xac00) % 28 !== 0;
}

function josa(word, a, b) {
  return hasBatchim(word) ? a : b;
}

function numToKorean(n) {
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
    if (part) out = chunk(part) + units[i] + out;
    x = Math.floor(x / 10000);
    i++;
  }
  return out;
}

function numToKoreanWon(n) {
  const intN = Math.floor(Number(n));
  if (!Number.isFinite(intN) || intN < 0) return "";
  if (intN === 0) return "금영원";
  const w = numToKorean(intN);
  return w ? `금${w}원` : "";
}

function budgetLine(total) {
  const intN = Math.floor(Number(total));
  const safe = Number.isFinite(intN) && intN >= 0 ? intN : 0;
  return `금${fmtMoney(safe)}원(${numToKoreanWon(safe)})`;
}

function forceBudgetLine(outline, payload) {
  const quote = payload?.quote || {};
  const items = Array.isArray(quote.items) ? quote.items : [];
  const total = Number.isFinite(quote.total) ? quote.total : computeTotal(items);
  const line = `3. 소요 예산: ${budgetLine(total)}`;

  const lines = String(outline || "").split(/\r?\n/);
  const idx = lines.findIndex((l) => l.trim().startsWith("3. 소요 예산:"));
  if (idx >= 0) {
    lines[idx] = line;
    return lines.join("\n").trim();
  }

  // Insert after "2. 품명" if possible, else append.
  const idx2 = lines.findIndex((l) => l.trim().startsWith("2. 품명:"));
  if (idx2 >= 0) {
    lines.splice(idx2 + 1, 0, line);
    return lines.join("\n").trim();
  }

  return (String(outline || "").trim() + "\n" + line).trim();
}

export async function onRequestOptions() {
  return new Response(null, { status: 204, headers: okCors() });
}

export async function onRequestPost(context) {
  try {
    const body = await context.request.json();
    const payload = body?.payload;
    if (!payload) return jsonResponse({ error: "payload is required" }, { status: 400, headers: okCors() });

    const apiKey = context.env?.OPENAI_API_KEY;
    if (!apiKey) {
      const outline = buildOutlineTemplate(payload);
      return jsonResponse({ mode: "template", outline }, { headers: okCors() });
    }

    const outlineRaw = await callOpenAI({ apiKey, payload });
    const outline = forceBudgetLine(outlineRaw, payload);
    return jsonResponse({ mode: "ai", outline }, { headers: okCors() });
  } catch (e) {
    const msg = e?.message ? String(e.message) : "unknown error";
    return jsonResponse({ error: msg }, { status: 500, headers: okCors() });
  }
}
