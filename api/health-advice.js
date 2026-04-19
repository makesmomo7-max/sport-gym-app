/**
 * Vercel Serverless: POST /api/health-advice
 * Env: GEMINI_API_KEY (required), GEMINI_MODEL (optional; same default as /api/gemini)
 *
 * Body (JSON):
 * - messages: { role: "user"|"assistant", text: string }[]  … 会話履歴（先頭は user 推奨）
 * - question: string … messages が無いときの単発質問
 * - context: { fatigue?, sleep?, stress?, pain?, note?, date? } … 任意。未指定は穏当な既定値
 * - hint: string … 任意。サーバー側で長さ制限（クライアントからの補足指示）
 * - facility_rules_ja: string … 任意。テナント設定の追加ルール（system に追記）
 *
 * user_id はプロンプトに含めません（識別子の扱いを避けるため）。
 */

var PAIN_JA = {
  shoulder: "肩",
  neck: "首",
  lower_back: "腰",
  knee: "膝",
  ankle: "足首",
  hip: "股関節",
  elbow: "肘",
  wrist: "手首"
};

function redactPII(text) {
  if (typeof text !== "string") return "";
  var t = text;
  t = t.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, "[メール]");
  t = t.replace(/(?:\+?81[-\s]?)?0\d{1,4}[-\s]?\d{1,4}[-\s]?\d{3,4}/g, "[電話]");
  t = t.replace(/\b\d{3}-\d{4}\b/g, "[郵便番号]");
  t = t.replace(
    /(東京都|北海道|(?:京都|大阪)府|.{2,3}県).{0,30}?(市|区|町|村).{0,30}?(丁目|番地|号)/g,
    "[住所]"
  );
  t = t.replace(/(株式会社|有限会社|合同会社|Inc\.|LLC|Ltd\.)\s*[^\s]{1,40}/g, "[会社名]");
  return t;
}

function clampInt(n, lo, hi) {
  var x = Math.floor(Number(n));
  if (isNaN(x)) return lo;
  return Math.max(lo, Math.min(hi, x));
}

function normalizePain(raw) {
  if (!Array.isArray(raw)) return [];
  var out = [];
  for (var i = 0; i < raw.length && out.length < 12; i++) {
    var s = String(raw[i] == null ? "" : raw[i])
      .trim()
      .slice(0, 48);
    if (s) out.push(s);
  }
  return out;
}

function painLine(arr) {
  if (!arr.length) return "なし";
  var parts = [];
  for (var i = 0; i < arr.length; i++) {
    var k = arr[i];
    parts.push(PAIN_JA[k] || k);
  }
  return parts.join("、");
}

function normalizeContext(body) {
  var c = body && body.context && typeof body.context === "object" ? body.context : {};
  var fatigue =
    c.fatigue == null || c.fatigue === "" ? 3 : clampInt(c.fatigue, 1, 5);

  var sleep = String(c.sleep != null ? c.sleep : "不明").trim().slice(0, 40) || "不明";
  var stress = String(c.stress != null ? c.stress : "不明").trim().slice(0, 40) || "不明";
  var note = String(c.note != null ? c.note : "")
    .trim()
    .slice(0, 280);
  var date = String(c.date != null ? c.date : "")
    .trim()
    .slice(0, 32);

  return {
    fatigue: fatigue,
    sleep: sleep,
    stress: stress,
    pain: normalizePain(c.pain),
    note: note,
    date: date
  };
}

function formatUserStateBlock(ctx) {
  var lines = [];
  if (ctx.date) lines.push("記録日: " + ctx.date);
  lines.push("疲労: " + ctx.fatigue + "/5（5が最も強い）");
  lines.push("睡眠: " + ctx.sleep);
  lines.push("ストレス: " + ctx.stress);
  lines.push("痛み: " + painLine(ctx.pain));
  if (ctx.note) lines.push("メモ: " + ctx.note);
  return lines.join("\n");
}

function buildSystemInstruction(ctx, hint, facilityRulesJa) {
  var base =
    "あなたはウェルビーイング（心身のコンディション）のアドバイザーです。momo fit 利用者の生活リズムとセルフケアを支えます。\n\n" +
    "【ユーザー状態（参考・最新）】\n" +
    formatUserStateBlock(ctx) +
    "\n\n" +
    "【ルール】\n" +
    "- 医療診断・病名の断定・治療方針の指示は禁止。\n" +
    "- 一般的生活習慣（休息、睡眠衛生、軽いストレッチ、仕事の区切り、受診を検討するタイミングの示唆など）に限定する。\n" +
    "- 日本語・です・ます調。優しく、短く、実用的に（最大8行程度。箇条書きは3つまで）。\n" +
    "- 個人の氏名・住所・電話・会員番号などの特定情報は求めず、入力されても復唱しない。\n" +
    "- 強い痛み、発熱、胸の痛み、呼吸困難、意識の異常などは医療受診を優先するよう促してよい（断定はしない）。\n";

  var fr =
    typeof facilityRulesJa === "string"
      ? redactPII(facilityRulesJa).trim().slice(0, 3500)
      : "";
  if (fr) {
    base += "\n【施設からの追加ルール】\n" + fr + "\n";
  }

  var h = typeof hint === "string" ? hint.trim().slice(0, 600) : "";
  if (h) {
    base += "\n【このターンの追加注意】\n" + h + "\n";
  }
  return base;
}

function geminiErrorDetail(d, httpStatus) {
  if (!d) return "Geminiからの応答をJSONとして読めませんでした (HTTP " + httpStatus + ")。";
  if (d.error && d.error.message) {
    return String(d.error.message) + (d.error.status ? " [" + d.error.status + "]" : "");
  }
  return "Gemini API error (HTTP " + httpStatus + ")";
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Cache-Control", "no-store, max-age=0");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ ok: false, message: "Method not allowed" });
  }

  var apiKey = String(process.env.GEMINI_API_KEY || "").trim();
  var model = (process.env.GEMINI_MODEL || "gemini-3-flash-preview").trim();

  if (!apiKey) {
    return res.status(500).json({ ok: false, message: "Missing GEMINI_API_KEY" });
  }

  try {
    var body = req.body;
    if (body == null || body === "") {
      return res.status(400).json({ ok: false, message: "Empty body" });
    }
    if (Buffer.isBuffer(body)) {
      try {
        body = JSON.parse(body.toString("utf8"));
      } catch (e) {
        return res.status(400).json({ ok: false, message: "Invalid JSON (buffer)" });
      }
    } else if (typeof body === "string") {
      try {
        body = JSON.parse(body);
      } catch (e) {
        return res.status(400).json({ ok: false, message: "Invalid JSON (string)" });
      }
    }
    body = body || {};

    var ctx = normalizeContext(body);
    var hint = body.hint;
    var facilityRulesJa = body.facility_rules_ja;

    var messages = Array.isArray(body.messages) ? body.messages : [];
    var question = typeof body.question === "string" ? body.question.trim() : "";

    if (messages.length === 0 && question) {
      messages = [{ role: "user", text: question }];
    }

    var raw = messages.filter(function (m) {
      return m && typeof m.role === "string" && typeof m.text === "string";
    });
    var start = 0;
    while (start < raw.length && raw[start].role !== "user") {
      start += 1;
    }
    var trimmed = raw.slice(start);
    if (trimmed.length > 24) trimmed = trimmed.slice(-24);
    if (trimmed.length === 0) {
      return res.status(400).json({ ok: false, message: "No user message (set messages or question)" });
    }

    var contents = trimmed.map(function (m) {
      var role = m.role === "user" ? "user" : "model";
      return {
        role: role,
        parts: [{ text: redactPII(m.text) }]
      };
    });

    var systemInstruction = buildSystemInstruction(ctx, hint, facilityRulesJa);

    var payload = {
      contents: contents,
      generationConfig: { temperature: 0.55 },
      systemInstruction: { parts: [{ text: systemInstruction }] }
    };

    var url =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      encodeURIComponent(model) +
      ":generateContent?key=" +
      encodeURIComponent(apiKey);

    var r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    var data = await r.json().catch(function () {
      return null;
    });

    if (!r.ok || !data) {
      return res.status(502).json({
        ok: false,
        message: geminiErrorDetail(data, r.status),
        model: model,
        httpStatus: r.status
      });
    }

    if (data.error && data.error.message) {
      return res.status(502).json({
        ok: false,
        message: geminiErrorDetail(data, r.status),
        model: model,
        httpStatus: r.status
      });
    }

    var text = "";
    try {
      text = String(
        data.candidates &&
          data.candidates[0] &&
          data.candidates[0].content &&
          data.candidates[0].content.parts &&
          data.candidates[0].content.parts[0] &&
          data.candidates[0].content.parts[0].text
          ? data.candidates[0].content.parts[0].text
          : ""
      );
    } catch (e) {
      text = "";
    }

    if (!text) {
      var cand = data.candidates && data.candidates[0];
      return res.status(502).json({
        ok: false,
        message: "Gemini returned empty text",
        finishReason: cand && cand.finishReason,
        blockReason: data.promptFeedback || (cand && cand.safetyRatings) || undefined
      });
    }

    return res.status(200).json({ ok: true, text: text });
  } catch (e) {
    return res.status(500).json({
      ok: false,
      message: e && e.message ? e.message : "Unexpected error"
    });
  }
};
