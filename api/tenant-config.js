/**
 * Vercel Serverless: GET /api/tenant-config?tenant=...&office=...
 * 返却は「既定JSON + config/tenants.v1.json の部分上書き」。
 * 機密（APIキー等）は含めない。
 *
 * 参照ファイル（リポジトリ直下）:
 * - config/tenant-defaults.v1.json
 * - config/tenants.v1.json  … { "tenants": { "<tenantId>": { ...部分 } } }
 */

var path = require("path");
var fs = require("fs");

function readJsonFile(relFromRepoRoot) {
  try {
    var p = path.join(__dirname, "..", relFromRepoRoot);
    var s = fs.readFileSync(p, "utf8");
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

function slice64(s) {
  return String(s == null ? "" : s)
    .trim()
    .slice(0, 64);
}

function shallowMerge(base, patch) {
  var out = {};
  var k;
  for (k in base) {
    if (Object.prototype.hasOwnProperty.call(base, k)) out[k] = base[k];
  }
  if (!patch || typeof patch !== "object") return out;
  for (k in patch) {
    if (!Object.prototype.hasOwnProperty.call(patch, k)) continue;
    if (!Object.prototype.hasOwnProperty.call(base, k)) continue;
    out[k] = patch[k];
  }
  return out;
}

var FALLBACK_DEFAULTS = {
  schema_version: "1.0",
  brand_name: "momo fit",
  support_url: "",
  booking_url: "",
  crisis_help_url: "",
  consult_line_url: "",
  consult_form_url: "",
  ai_enabled: true,
  health_advisor_enabled: true,
  ai_disclaimer_ja:
    "本アプリのAIは、運動・睡眠・ストレスなどウェルビーイング（心身のコンディション）の一般情報・セルフケアの参考としてお答えします。医療診断・治療指示は行いません。強い症状がある場合は医療機関へご相談ください。",
  ai_extra_rules_ja: "",
  gemini_model_override: "",
  feature_share_trainer: true,
  feature_sleep_log: true,
  feature_pulse: true,
  feature_gas_submit: true,
  privacy_policy_url: "",
  data_retention_note_ja:
    "会話本文は原則としてサーバーに長期保存しない方針です。コンディション等の記録は端末内を主とし、ジム設定により外部へ集約される場合があります（運用ポリシーに合わせて変更してください）。"
};

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Cache-Control", "public, max-age=60, stale-while-revalidate=120");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({ ok: false, message: "Method not allowed" });
  }

  var tenant = slice64(req.query && req.query.tenant);
  var office = slice64(req.query && req.query.office);

  var defaults = readJsonFile("config/tenant-defaults.v1.json");
  if (!defaults || typeof defaults !== "object") defaults = FALLBACK_DEFAULTS;

  var tenantsFile = readJsonFile("config/tenants.v1.json");
  var byTenant =
    tenantsFile &&
    tenantsFile.tenants &&
    typeof tenantsFile.tenants === "object"
      ? tenantsFile.tenants
      : {};

  var patch = tenant && byTenant[tenant] ? byTenant[tenant] : {};
  var config = shallowMerge(defaults, patch);

  return res.status(200).json({
    ok: true,
    tenant_id: tenant,
    office_id: office,
    config: config
  });
};
