/**
 * テナント設定（/api/tenant-config）から相談導線URLを読み、画面の a 要素を更新する。
 */
(function (g) {
  "use strict";

  var TENANT_KEY = "eap_tenant";
  var OFFICE_KEY = "eap_office";

  function slice64(s) {
    return String(s == null ? "" : s)
      .trim()
      .slice(0, 64);
  }

  function tenantOfficeFromStorage() {
    var t = "";
    var o = "";
    try {
      t = slice64(localStorage.getItem(TENANT_KEY));
    } catch (e) {}
    try {
      o = slice64(localStorage.getItem(OFFICE_KEY));
    } catch (e) {}
    return { tenant: t, office: o };
  }

  function isHttp(u) {
    return /^https?:\/\//i.test(String(u || "").trim());
  }

  function hrefAppPage(htmlFile, searchParams) {
    var ids = tenantOfficeFromStorage();
    var q =
      searchParams instanceof URLSearchParams
        ? searchParams
        : new URLSearchParams();
    if (ids.tenant && !q.has("tenant")) q.set("tenant", ids.tenant);
    if (ids.office && !q.has("office")) q.set("office", ids.office);
    var qs = q.toString();
    return "./" + htmlFile + (qs ? "?" + qs : "");
  }

  function showIfHttp(el, url) {
    if (!el) return;
    var u = String(url || "").trim();
    if (isHttp(u)) {
      el.href = u;
      el.style.display = "";
    } else {
      el.style.display = "none";
    }
  }

  function applyConsultUi(cfg) {
    cfg = cfg && typeof cfg === "object" ? cfg : {};

    var chatQ = new URLSearchParams();
    chatQ.set("eap", "1");
    var chatHref = hrefAppPage("chat.html", chatQ);
    var chatA = document.getElementById("consultChatLink");
    if (chatA) chatA.href = chatHref;
    var eapChat = document.getElementById("eapChatBtn");
    if (eapChat) eapChat.href = chatHref;

    var book = String(cfg.booking_url || "").trim();
    var bookA = document.getElementById("consultBookLink");
    if (bookA) {
      if (isHttp(book)) bookA.href = book;
      else bookA.href = hrefAppPage("yoyaku_employee.html", new URLSearchParams());
    }

    showIfHttp(document.getElementById("consultLineLink"), cfg.consult_line_url);
    showIfHttp(document.getElementById("consultFormLink"), cfg.consult_form_url);
    showIfHttp(document.getElementById("eapLineBtn"), cfg.consult_line_url);
    showIfHttp(document.getElementById("eapFormBtn"), cfg.consult_form_url);
  }

  function load(cb) {
    var ids = tenantOfficeFromStorage();
    var q = new URLSearchParams();
    if (ids.tenant) q.set("tenant", ids.tenant);
    if (ids.office) q.set("office", ids.office);
    fetch("/api/tenant-config?" + q.toString(), { cache: "no-store" })
      .then(function (r) {
        return r.json().catch(function () {
          return null;
        });
      })
      .then(function (data) {
        var cfg = data && data.ok === true && data.config ? data.config : null;
        applyConsultUi(cfg);
        if (cb) cb(cfg);
      })
      .catch(function () {
        applyConsultUi(null);
        if (cb) cb(null);
      });
  }

  g.MomoTenantConsult = {
    load: load,
    applyConsultUi: applyConsultUi
  };
})(typeof window !== "undefined" ? window : this);
