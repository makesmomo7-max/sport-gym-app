/**
 * Vercel Serverless: POST /api/gemini
 * Env: GEMINI_API_KEY (required), GEMINI_MODEL (optional, default gemini-3-flash-preview)
 */
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

  const apiKey = process.env.GEMINI_API_KEY;
  const model = process.env.GEMINI_MODEL || "gemini-3-flash-preview";

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
    const messages = Array.isArray(body.messages) ? body.messages : [];
    const system = typeof body.system === "string" ? body.system : "";

    function redactPII(text) {
      if (typeof text !== "string") return "";
      let t = text;
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

    // Gemini の会話は原則「user」から始める。先頭の model（初期挨拶など）があると空レスポンスになりやすい。
    var raw = messages.filter(function (m) {
      return m && typeof m.role === "string" && typeof m.text === "string";
    });
    var start = 0;
    while (start < raw.length && raw[start].role !== "user") {
      start += 1;
    }
    var trimmed = raw.slice(start);
    if (trimmed.length === 0) {
      return res.status(400).json({ ok: false, message: "No user message in payload" });
    }

    const contents = trimmed.map(function (m) {
      return {
        role: m.role === "user" ? "user" : "model",
        parts: [{ text: redactPII(m.text) }]
      };
    });

    const payload = {
      contents: contents,
      generationConfig: { temperature: 0.7 }
    };
    if (system) {
      payload.systemInstruction = { parts: [{ text: system }] };
    }

    const url =
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      encodeURIComponent(model) +
      ":generateContent?key=" +
      encodeURIComponent(apiKey);

    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await r.json().catch(function () {
      return null;
    });

    function geminiErrorDetail(d, httpStatus) {
      if (!d) return "Geminiからの応答をJSONとして読めませんでした (HTTP " + httpStatus + ")。";
      if (d.error && d.error.message) {
        return String(d.error.message) + (d.error.status ? " [" + d.error.status + "]" : "");
      }
      return "Gemini API error (HTTP " + httpStatus + ")";
    }

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
