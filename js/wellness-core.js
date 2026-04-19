/**
 * momo fit — 日次ウェルネス・リスク判定（クライアントのみ）
 * 医療判断ではなく、セルフモニタリングと相談導線用の簡易ルールです。
 */
(function (global) {
  "use strict";

  var LOG_KEY = "momo_wellness_daily_v1";
  var EAP_DISMISS_KEY = "momo_eap_red_dismissed_ymd";

  function pad2(n) {
    return String(n).length < 2 ? "0" + n : String(n);
  }

  function ymdLocal(d) {
    var x = d instanceof Date ? d : new Date();
    return x.getFullYear() + "-" + pad2(x.getMonth() + 1) + "-" + pad2(x.getDate());
  }

  function parseYmd(s) {
    var p = String(s || "").split("-");
    if (p.length !== 3) return null;
    var y = +p[0];
    var m = +p[1];
    var day = +p[2];
    if (!y || !m || !day) return null;
    return new Date(y, m - 1, day);
  }

  function addDaysYmd(ymd, delta) {
    var d = parseYmd(ymd);
    if (!d) return ymd;
    d.setDate(d.getDate() + delta);
    return ymdLocal(d);
  }

  function safeJsonParse(raw, fallback) {
    try {
      var o = JSON.parse(raw);
      return o == null ? fallback : o;
    } catch (e) {
      return fallback;
    }
  }

  function readWellnessLogs() {
    var raw = "";
    try {
      raw = global.localStorage.getItem(LOG_KEY) || "";
    } catch (e) {
      raw = "";
    }
    var arr = safeJsonParse(raw, []);
    return Array.isArray(arr) ? arr : [];
  }

  function logsByDate() {
    var map = {};
    readWellnessLogs().forEach(function (row) {
      if (!row || typeof row.date !== "string") return;
      var k = row.date.slice(0, 10);
      map[k] = row;
    });
    return map;
  }

  function readCheckHistory() {
    try {
      var a = safeJsonParse(global.localStorage.getItem("history_check") || "[]", []);
      return Array.isArray(a) ? a : [];
    } catch (e) {
      return [];
    }
  }

  function stressFromMoodLevel(mood) {
    var m = Math.floor(Number(mood));
    if (isNaN(m) || m < 1 || m > 5) return null;
    return 6 - m;
  }

  function ymdFromIso(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return null;
    return ymdLocal(d);
  }

  /** その日の check 履歴からストレス1〜5（高いほどストレス大）を推定 */
  function stressFromCheckForYmd(ymd) {
    var hist = readCheckHistory();
    var best = null;
    for (var i = hist.length - 1; i >= 0; i--) {
      var r = hist[i];
      if (!r || !r.date) continue;
      var d = ymdFromIso(r.date);
      if (d !== ymd) continue;
      var s = stressFromMoodLevel(r.mood);
      if (s != null) best = s;
    }
    return best;
  }

  function readSleepHistory() {
    try {
      var a = safeJsonParse(global.localStorage.getItem("history_sleep") || "[]", []);
      return Array.isArray(a) ? a : [];
    } catch (e) {
      return [];
    }
  }

  function sleepQualityForYmd(ymd) {
    var hist = readSleepHistory();
    var q = null;
    for (var i = hist.length - 1; i >= 0; i--) {
      var r = hist[i];
      if (!r || String(r.recordDate || "").slice(0, 10) !== ymd) continue;
      var n = Math.floor(Number(r.quality));
      if (!isNaN(n) && n >= 1 && n <= 5) q = n;
    }
    return q;
  }

  function effectiveStress(ymd, wmap) {
    var row = wmap[ymd];
    if (row && row.stress1to5 != null) {
      var s = Math.floor(Number(row.stress1to5));
      if (!isNaN(s) && s >= 1 && s <= 5) return s;
    }
    return stressFromCheckForYmd(ymd);
  }

  function effectiveSleepSat(ymd, wmap) {
    var row = wmap[ymd];
    if (row && row.sleepSat1to5 != null) {
      var q = Math.floor(Number(row.sleepSat1to5));
      if (!isNaN(q) && q >= 1 && q <= 5) return q;
    }
    return sleepQualityForYmd(ymd);
  }

  function hasWellnessRow(ymd, wmap) {
    return !!wmap[ymd];
  }

  /** true / false / null（null＝トレ記録なし→未実施連続には含めない） */
  function trainingDoneForYmd(ymd, wmap) {
    var row = wmap[ymd];
    if (!row) return null;
    if (row.trainingDone === true) return true;
    if (row.trainingDone === false) return false;
    return null;
  }

  function consecutiveStressHigh(wmap, todayYmd, minDays) {
    var streak = 0;
    for (var i = 0; i < 14; i++) {
      var y = addDaysYmd(todayYmd, -i);
      var s = effectiveStress(y, wmap);
      if (s == null) {
        streak = 0;
        continue;
      }
      if (s >= 4) {
        streak++;
        if (streak >= minDays) return true;
      } else {
        streak = 0;
      }
    }
    return false;
  }

  function consecutiveSleepLow(wmap, todayYmd, minDays) {
    var streak = 0;
    for (var i = 0; i < 14; i++) {
      var y = addDaysYmd(todayYmd, -i);
      var sat = effectiveSleepSat(y, wmap);
      if (sat == null) {
        streak = 0;
        continue;
      }
      if (sat <= 2) {
        streak++;
        if (streak >= minDays) return true;
      } else {
        streak = 0;
      }
    }
    return false;
  }

  /** ウェルネス記録がある連続日で、トレーニング未完了が続く */
  function consecutiveTrainingMissed(wmap, todayYmd, need) {
    var streak = 0;
    for (var i = 0; i < 14; i++) {
      var y = addDaysYmd(todayYmd, -i);
      if (!hasWellnessRow(y, wmap)) {
        streak = 0;
        continue;
      }
      var done = trainingDoneForYmd(y, wmap);
      if (done === true) {
        streak = 0;
      } else if (done === false) {
        streak++;
        if (streak >= need) return true;
      } else {
        streak = 0;
      }
    }
    return false;
  }

  var TEMPLATE = {
    green: "いい流れです、この調子でいきましょう",
    yellow: "少し疲れが見えます。無理しすぎていませんか？",
    red: "少し気になる状態です。無理せず休みも大事にしてください"
  };

  function computeRiskState() {
    var today = ymdLocal(new Date());
    var wmap = logsByDate();
    var reasons = [];

    var redStress = consecutiveStressHigh(wmap, today, 2);
    var redSleep = consecutiveSleepLow(wmap, today, 2);
    var redTrain = consecutiveTrainingMissed(wmap, today, 3);
    if (redStress) reasons.push("ストレスが高めの日が続いています");
    if (redSleep) reasons.push("睡眠の満足度が低めの日が続いています");
    if (redTrain) reasons.push("トレーニング記録の完了が続いていません");

    var level;
    if (redStress || redSleep || redTrain) {
      level = "red";
    } else {
      var yellowHit = false;
      for (var yi = 0; yi < 3; yi++) {
        var ymd = addDaysYmd(today, -yi);
        var s = effectiveStress(ymd, wmap);
        var z = effectiveSleepSat(ymd, wmap);
        if (s === 3 || z === 3) {
          yellowHit = true;
          break;
        }
      }
      level = yellowHit ? "yellow" : "green";
    }

    return {
      level: level,
      reasons: reasons,
      templateLine: TEMPLATE[level],
      todayYmd: today
    };
  }

  function isEapDismissedToday() {
    try {
      return global.localStorage.getItem(EAP_DISMISS_KEY) === ymdLocal(new Date());
    } catch (e) {
      return false;
    }
  }

  function dismissEapForToday() {
    try {
      global.localStorage.setItem(EAP_DISMISS_KEY, ymdLocal(new Date()));
    } catch (e) {}
  }

  function stress15FromMoodLevel(mood) {
    var m = Math.floor(Number(mood));
    if (isNaN(m) || m < 1 || m > 5) return null;
    return 6 - m;
  }

  /**
   * check.html 保存成功時に同日のウェルネス行へマージ（ストレス・睡眠満足度・一言）。
   * トレ完了は未設定のままにし、未実施連続判定に影響しないようにする。
   */
  function mergeCheckIntoWellness(bodyLevel, moodLevel, noteTrim) {
    var ymd = ymdLocal(new Date());
    var stress15 = stress15FromMoodLevel(moodLevel);
    if (stress15 == null) return;
    var list = readWellnessLogs();
    if (!Array.isArray(list)) list = [];
    var idx = -1;
    for (var i = 0; i < list.length; i++) {
      if (list[i] && list[i].date === ymd) idx = i;
    }
    var sleepQ = sleepQualityForYmd(ymd);
    var note = String(noteTrim || "")
      .trim()
      .slice(0, 300);
    if (idx >= 0) {
      var row = list[idx];
      row.stress1to5 = stress15;
      if (row.sleepSat1to5 == null && sleepQ != null) row.sleepSat1to5 = sleepQ;
      if (note) {
        var prev = String(row.oneLine || "").trim();
        row.oneLine = (prev ? prev + "／" : "") + note;
        if (row.oneLine.length > 300) row.oneLine = row.oneLine.slice(0, 300);
      }
      row.checkMergedAt = new Date().toISOString();
      list[idx] = row;
    } else {
      list.push({
        date: ymd,
        stress1to5: stress15,
        sleepSat1to5: sleepQ != null ? sleepQ : null,
        steps: null,
        sleepHours: null,
        trainingMenu: "",
        trainingComment: "",
        mealText: "",
        oneLine: note,
        didWin: false,
        fromCheckMerge: true,
        savedAt: new Date().toISOString()
      });
    }
    list.sort(function (a, b) {
      return String(a.date || "").localeCompare(String(b.date || ""));
    });
    if (list.length > 400) list = list.slice(-400);
    try {
      global.localStorage.setItem(LOG_KEY, JSON.stringify(list));
    } catch (e) {}
  }

  global.MomoWellness = {
    LOG_KEY: LOG_KEY,
    ymdLocal: ymdLocal,
    readWellnessLogs: readWellnessLogs,
    logsByDate: logsByDate,
    computeRiskState: computeRiskState,
    isEapDismissedToday: isEapDismissedToday,
    dismissEapForToday: dismissEapForToday,
    mergeCheckIntoWellness: mergeCheckIntoWellness,
    effectiveStress: function (ymd) {
      return effectiveStress(ymd, logsByDate());
    },
    effectiveSleepSat: function (ymd) {
      return effectiveSleepSat(ymd, logsByDate());
    }
  };
})(typeof window !== "undefined" ? window : this);
