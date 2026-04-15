// すべてこの1冊に集約（予約・コンディション・脈拍・睡眠・月次集計）
// 「makes-momo EAP 予約管理」など既存ブックのIDをそのまま使用
const EAP_SPREADSHEET_ID = "1UcojM3nRGxuyAI_q8QJZbLLQ8B7E67i3nSlvmnkAYG8";
const RESERVATION_SPREADSHEET_ID = EAP_SPREADSHEET_ID;
const HEALTH_SPREADSHEET_ID = EAP_SPREADSHEET_ID;

const RESERVATION_SHEET_NAME = "予約一覧";
const GUIDE_SHEET_NAME = "00_使い方";
const ENTERPRISE_MASTER_SHEET_NAME = "01_企業マスタ";
const OFFICE_MASTER_SHEET_NAME = "02_事業所マスタ";
/**
 * 02_事業所マスタ の列定義（左からこの順・列位置固定）
 * A=officeId, B=tenantId, C=事業所表示名 は GAS の lookup 用。途中に列を挟まないこと。
 */
const OFFICE_MASTER_HEADER = [
  "officeId",
  "tenantId",
  "事業所表示名",
  "所在地",
  "従業員数_参考",
  "メモ"
];
/**
 * 01_企業マスタ の列定義（左からこの順・列位置固定）
 * A=tenantId, B=企業表示名 は GAS の lookup 用。途中に列を挟まないこと。
 */
const ENTERPRISE_MASTER_HEADER = [
  "tenantId",
  "企業表示名",
  "正式社名",
  "契約ステータス",
  "契約開始日",
  "契約終了日",
  "従業員数_参考",
  "窓口_部署",
  "窓口_担当者",
  "窓口_連絡先",
  "月次サマリ送付先",
  "配布用URLメモ",
  "メモ"
];

const RESERVATION_HEADER = [
  "受付日時",
  "お名前",
  "所属企業名",
  "メールアドレス",
  "相談方法",
  "第1希望",
  "第2希望",
  "第3希望",
  "相談内容",
  "tenantId",
  "officeId"
];

const HEALTH_SHEET_NAME = "記録一覧";
const HEALTH_HEADER = [
  "受付日時",
  "利用者種別",
  "体の調子",
  "気分・メンタル",
  "今日の一言",
  "tenantId",
  "officeId"
];

const PULSE_SHEET_NAME = "脈拍記録";
const PULSE_HEADER = [
  "受付日時",
  "自律神経スコア",
  "体調レベル",
  "心拍数(bpm)",
  "HRV(ms)",
  "tenantId",
  "officeId"
];

// 睡眠：数値・固定タグのみ（自由記述なし）。企業向けは「睡眠_月次集計」シートの集計行のみを利用する想定。
const SLEEP_SHEET_NAME = "睡眠記録";
const SLEEP_HEADER = [
  "受付日時",
  "tenantId",
  "recordDate",
  "sleepMinutes",
  "latencyMin",
  "wakeups",
  "quality",
  "tags",
  "officeId",
  "alarmTargetHm",
  "sessionStartedAt",
  "alarmDismissedAt"
];

const SLEEP_MONTHLY_SHEET_NAME = "睡眠_月次集計";
const SLEEP_MONTHLY_HEADER = [
  "年月",
  "tenantId",
  "officeId",
  "企業表示名",
  "事業所表示名",
  "回答数",
  "集計可否",
  "平均スコア",
  "平均睡眠時間_分",
  "平均夜間覚醒回数",
  "平均入眠潜時_分",
  "タグ_飲酒_pct",
  "タグ_カフェイン_pct",
  "タグ_運動_pct",
  "タグ_食事_pct",
  "タグ_喫煙_pct",
  "集計実行日時"
];

/** この人数未満の tenant は「集計可否」を母体不足にし、平均値は空欄にする */
const SLEEP_AGG_MIN_N = 5;

/** 会員がトレーナー共有画面から送信したサマリー（任意・ジム運用で利用） */
const TRAINER_SHARE_SHEET_NAME = "トレーナー共有";
const TRAINER_SHARE_HEADER = [
  "受付日時",
  "表示名・呼び名",
  "本人からのメッセージ",
  "自動まとめ本文",
  "tenantId",
  "officeId"
];

const COL_PREFER1 = 6; // 「第1希望」（予約一覧）

function doPost(e) {
  try {
    touchAuxiliarySheets_();

    const p = (e && e.parameter) ? e.parameter : {};
    const formType = String(p.formType || "reservation").trim();

    if (formType === "health") {
      const sheet = getSheet_(HEALTH_SPREADSHEET_ID, HEALTH_SHEET_NAME);
      ensureHeader_(sheet, HEALTH_HEADER);
      const tid = String(p.tenantId || "").trim().slice(0, 64);
      const oid = String(p.officeId || "").trim().slice(0, 64);
      var healthRow = [
        p.submittedAt || "",
        p.userType || "",
        p.bodyLevel || "",
        p.moodLevel || "",
        p.note || "",
        tid,
        oid
      ];
      sheet.appendRow(healthRow);
      appendOfficeMirrorRow_(oid, "記録", HEALTH_HEADER, healthRow);
      return jsonResponse_({ ok: true });
    }

    if (formType === "pulse") {
      const sheet = getSheet_(HEALTH_SPREADSHEET_ID, PULSE_SHEET_NAME);
      ensureHeader_(sheet, PULSE_HEADER);
      const tid = String(p.tenantId || "").trim().slice(0, 64);
      const oid = String(p.officeId || "").trim().slice(0, 64);
      var pulseRow = [
        p.submittedAt || "",
        p.score || "",
        p.condition || "",
        p.bpm || "",
        p.hrv || "",
        tid,
        oid
      ];
      sheet.appendRow(pulseRow);
      appendOfficeMirrorRow_(oid, "脈拍", PULSE_HEADER, pulseRow);
      return jsonResponse_({ ok: true });
    }

    if (formType === "sleep") {
      const sheet = getSheet_(HEALTH_SPREADSHEET_ID, SLEEP_SHEET_NAME);
      ensureHeader_(sheet, SLEEP_HEADER);
      const tenantId = String(p.tenantId || "").trim().slice(0, 64);
      const officeId = String(p.officeId || "").trim().slice(0, 64);
      const recordDate = normalizeRecordDate_(p.recordDate);
      if (!recordDate) {
        return jsonResponse_({ ok: false, message: "recordDate invalid" }, 400);
      }
      const sleepMinutes = clampInt_(parseNum_(p.sleepMinutes), 30, 960);
      const latencyMin = clampInt_(parseNum_(p.latencyMin), 0, 240);
      const wakeups = clampInt_(parseNum_(p.wakeups), 0, 30);
      const quality = clampInt_(parseNum_(p.quality), 1, 5);
      const tags = sanitizeSleepTags_(p.tags);
      var alarmTargetHm = String(p.alarmTargetHm || "").trim().slice(0, 8);
      var sessionStartedAt = String(p.sessionStartedAt || "").trim().slice(0, 64);
      var alarmDismissedAt = String(p.alarmDismissedAt || "").trim().slice(0, 64);
      var sleepRow = [
        p.submittedAt || "",
        tenantId,
        recordDate,
        sleepMinutes,
        latencyMin,
        wakeups,
        quality,
        tags,
        officeId,
        alarmTargetHm,
        sessionStartedAt,
        alarmDismissedAt
      ];
      sheet.appendRow(sleepRow);
      appendOfficeMirrorRow_(officeId, "睡眠", SLEEP_HEADER, sleepRow);
      return jsonResponse_({ ok: true });
    }

    if (formType === "trainerShare") {
      const sheet = getSheet_(HEALTH_SPREADSHEET_ID, TRAINER_SHARE_SHEET_NAME);
      ensureHeader_(sheet, TRAINER_SHARE_HEADER);
      const tid = String(p.tenantId || "").trim().slice(0, 64);
      const oid = String(p.officeId || "").trim().slice(0, 64);
      var shareRow = [
        p.submittedAt || "",
        String(p.memberLabel || "").trim().slice(0, 128),
        String(p.memberNote || "").trim().slice(0, 2000),
        String(p.summaryText || "").trim().slice(0, 8000),
        tid,
        oid
      ];
      sheet.appendRow(shareRow);
      appendOfficeMirrorRow_(oid, "トレーナー共有", TRAINER_SHARE_HEADER, shareRow);
      return jsonResponse_({ ok: true });
    }

    // default: reservation
    const sheet = getSheet_(RESERVATION_SPREADSHEET_ID, RESERVATION_SHEET_NAME);
    ensureHeader_(sheet, RESERVATION_HEADER);
    const resTid = String(p.tenantId || "").trim().slice(0, 64);
    const resOid = String(p.officeId || "").trim().slice(0, 64);
    var resRow = [
      p.submittedAt || "",
      p.fullName || "",
      p.company || "",
      p.email || "",
      p.method || "",
      p.prefer1 || "",
      p.prefer2 || "",
      p.prefer3 || "",
      p.detail || "",
      resTid,
      resOid
    ];
    sheet.appendRow(resRow);
    appendOfficeMirrorRow_(resOid, "予約", RESERVATION_HEADER, resRow);

    // メール通知
    try {
      var subject = "【momo fit】新しい予約が入りました";
      var body =
        "新しい予約が入りました。\n\n" +
        "━━━━━━━━━━━━━━━━\n" +
        "お名前：" + (p.fullName || "") + "\n" +
        "所属企業：" + (p.company || "") + "\n" +
        "メール：" + (p.email || "") + "\n" +
        "相談方法：" + (p.method || "") + "\n" +
        "第1希望：" + (p.prefer1 || "") + "\n" +
        "第2希望：" + (p.prefer2 || "") + "\n" +
        "第3希望：" + (p.prefer3 || "") + "\n" +
        "相談内容：" + (p.detail || "") + "\n" +
        "━━━━━━━━━━━━━━━━\n\n" +
        "スプレッドシートで確認：\n" +
        "https://docs.google.com/spreadsheets/d/" + RESERVATION_SPREADSHEET_ID + "/edit";
      MailApp.sendEmail("makesmomo7@gmail.com", subject, body);
    } catch (mailErr) {
      // メール送信失敗しても予約自体は成功とする
    }

    return jsonResponse_({ ok: true });
  } catch (error) {
    return jsonResponse_(
      { ok: false, message: error && error.message ? error.message : "Unexpected error" },
      500
    );
  }
}

function doGet() {
  try {
    const sheet = getSheet_(RESERVATION_SPREADSHEET_ID, RESERVATION_SHEET_NAME);
    ensureHeader_(sheet, RESERVATION_HEADER);
    const bookedPrefer1 = getBookedPrefer1_(sheet);

    return jsonResponse_({
      ok: true,
      bookedPrefer1: bookedPrefer1
    });
  } catch (error) {
    return jsonResponse_(
      {
        ok: false,
        message: error && error.message ? error.message : "Unexpected error"
      },
      500
    );
  }
}

function doOptions() {
  return jsonResponse_({ ok: true });
}

function getSheet_(spreadsheetId, sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 1行目をヘッダー行として確保。列が足りない場合は右側に列名を拡張（既存シートへの列追加に対応）
 */
function ensureHeader_(sheet, header) {
  var ncol = header.length;
  var cur = sheet.getLastColumn();
  if (cur === 0) {
    sheet.getRange(1, 1, 1, ncol).setValues([header]);
    return;
  }
  if (cur < ncol) {
    var row1 = sheet.getRange(1, 1, 1, cur).getValues()[0];
    var merged = [];
    for (var i = 0; i < ncol; i++) {
      merged.push(i < cur && String(row1[i]).trim() !== "" ? row1[i] : header[i]);
    }
    sheet.getRange(1, 1, 1, ncol).setValues([merged]);
    return;
  }
  var full = sheet.getRange(1, 1, 1, ncol).getValues()[0];
  var hasAny = full.some(function (v) { return String(v).trim() !== ""; });
  if (!hasAny) {
    sheet.getRange(1, 1, 1, ncol).setValues([header]);
  }
}

/** ガイド・企業マスタを用意（初回のみタブ位置を調整） */
function touchAuxiliarySheets_() {
  try {
    var ss = SpreadsheetApp.openById(EAP_SPREADSHEET_ID);
    ensureGuideSheet_(ss);
    ensureEnterpriseMasterSheet_(ss);
    ensureOfficeMasterSheet_(ss);
  } catch (err) {}
}

function ensureGuideSheet_(ss) {
  var sh = ss.getSheetByName(GUIDE_SHEET_NAME);
  var created = !sh;
  if (!sh) sh = ss.insertSheet(GUIDE_SHEET_NAME);
  var text =
    "【makes-momo EAP データの見方（1冊に集約）】\n\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "■ まず最初に（データが企業ごとに分かれるしくみ）\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "例：企業Aに勤める「あいさん」と、企業Bに勤める「いきさん」が、同じ画面（コンディション記録・睡眠など）から送信したとき、\n" +
    "システムが自動で「どちらの企業のデータか」を判断するには、**あらかじめ企業ごとに違うURL（リンク）を渡す**必要があります。\n\n" +
    "【ポイント】\n" +
    "・人の名前や、予約フォームの「所属企業名」の入力だけでは、機械的には企業を識別しません（自由記述のため）。\n" +
    "・**URL に付いている「企業コード（tenantId）」**が、スプレッドシートの各行の末尾にそのまま保存されます。\n" +
    "・従業員が**初めて正しいURLでページを開く**と、その端末（スマホのブラウザ）にコードが覚えられ、次回以降も同じ企業として送られます。\n\n" +
    "【具体例】\n" +
    "・あいさん（企業A）には … **gym.html?tenant=acme** または **check.html?tenant=acme** のようなリンクを渡す → tenantId 列に acme と入る\n" +
    "・いきさん（企業B）には … **gym.html?tenant=fuji** または **check.html?tenant=fuji** … → tenantId 列に fuji と入る\n" +
    "※ acme や fuji は例です。実際には 01_企業マスタ の A列に書いた英数字コードと**一字一句そろえる**必要があります。\n\n" +
    "【サイトのURLと gym.html について】\n" +
    "・**「あなたのサイト」＝公開しているドメイン全体**です。企業ごとに別サイトを用意するのではなく、**同じアドレスの下にページが並んでいる**だけです。\n" +
    "・現在の公開例（Vercel）：**https://makes-momo-eap.vercel.app** … このドメインに gym.html / check.html などがあります。\n" +
    "・**gym.html** … ジム会員向け**ホーム（入口・メニュー）**。「コンディション記録」「睡眠」などへのボタンがあるページです。\n" +
    "・**check.html** … コンディション記録、**sleep.html** … 睡眠、**pulse.html** … 脈拍、**yoyaku_employee.html** … 予約。用途ごとに名前が違います。\n" +
    "・配布は **gym.html?tenant=コード**（ホームから入る）と **check.html?tenant=コード**（機能に直行）の**どちらでも可**。どちらも同じドメインです。\n" +
    "・gym.html を **?tenant= 付きで開いたあと**、画面内のリンク（記録・睡眠・予約など）には **同じ tenant / office が自動で付き**、端末にも覚えられます（ホーム1本だけコード付きで配っても、中のページでコードが消えにくい仕様）。\n" +
    "・独自ドメインに変えた場合は、説明中の **https://makes-momo-eap.vercel.app** を、自分の **https://（自分のドメイン）** に読み替えてください。\n\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "■ 運用の手順（初めての方・細かい順番）\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "【手順1】契約した企業を「01_企業マスタ」に1社ずつ登録する\n" +
    "　1-1. タブ「01_企業マスタ」を開く。\n" +
    "　1-2. 2行目以降に、企業ごとに1行ずつ入力する（1行目は列名の行のままにする）。\n" +
    "　1-3. **A列 tenantId** … その企業専用の**短いコード**（英数字・他社と重複しないもの）。例：acme、fuji2026\n" +
    "　1-4. **B列 企業表示名** … 表や月次レポートに出したい名前（株式会社◯◯ など短くても可）\n" +
    "　1-5. その他の列（正式社名・契約日など）は分かる範囲で。空欄でも動作自体はします。\n\n" +
    "【手順2】従業員が使う「サイト」とページ名を確認する\n" +
    "　2-1. ブラウザで **https://makes-momo-eap.vercel.app/gym.html** を開けるか確認する（別ドメインにしている場合はそちら）。\n" +
    "　2-2. **gym.html** … ホーム。**check.html** … コンディション、**sleep.html** … 睡眠、**pulse.html** … 脈拍、**yoyaku_employee.html** … 予約。\n" +
    "　2-3. 企業コードを付けるのは **gym.html でも各 .html でも可**。中身は同じサイト上の別ページです。\n\n" +
    "【手順3】企業ごとに「専用リンク」を作る（コピー用の例）\n" +
    "　3-1. 企業Aの tenantId が **acme** のとき（ホームから入る）：\n" +
    "　　　https://makes-momo-eap.vercel.app/gym.html?tenant=acme\n" +
    "　3-2. 同じ企業でコンディションに直行する例：\n" +
    "　　　https://makes-momo-eap.vercel.app/check.html?tenant=acme\n" +
    "　3-3. 企業Bの tenantId が **fuji** のとき：\n" +
    "　　　https://makes-momo-eap.vercel.app/gym.html?tenant=fuji\n" +
    "　3-4. 事業所まで分けたい場合（**02_事業所マスタ**の officeId と一致させる）：\n" +
    "　　　https://makes-momo-eap.vercel.app/gym.html?tenant=acme&office=tokyo_hq\n" +
    "　　　または …/check.html?tenant=acme&office=tokyo_hq\n" +
    "　　　**&** で tenant と office を続けます。office= の値は「02_事業所マスタ」の A列と、B列 tenantId がその企業と一致している必要があります。\n\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "■ アドレスの設定の仕方（1社ぶんをゼロから決める）\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "【1】01_企業マスタに、その会社の1行を書く\n" +
    "　・**A列 tenantId** … 英数字だけの**社専用コード**（他社と重ならないもの）。短く・打ち間違いしにくいものがおすすめ。日本語は使えません。\n" +
    "　・**B列 企業表示名** … 「旭食品」など、人が見て分かる社名・通称。\n" +
    "　※ URL に載せるのは **A列のコード**です。**会社名そのものは URL には書きません**（B列は表・レポート用）。\n\n" +
    "【2】その会社の従業員に渡すアドレス（リンク）を組み立てる\n" +
    "　・土台は次のどちらか（よく使うのはホーム）。\n" +
    "　　　ホーム … https://makes-momo-eap.vercel.app/gym.html?tenant=（ここにA列と同じコード）\n" +
    "　　　記録だけ先 … https://makes-momo-eap.vercel.app/check.html?tenant=（ここにA列と同じコード）\n" +
    "　・**（ここにA列と同じコード）**の部分を、マスタの A列と**一字一句同じ**に置き換える。\n" +
    "　・独自ドメインにしている場合は、**https://makes-momo-eap.vercel.app** だけを自分のドメインに読み替える。\n\n" +
    "【具体例】企業表示名が「旭食品」の場合（tenantId は運用で決める。ここでは例として **asashoku** とする）\n" +
    "　・マスタの登録例：A列 **asashoku** ／ B列 **旭食品**\n" +
    "　・従業員にメール・QRで渡す文面の例（ホーム）：\n" +
    "　　　https://makes-momo-eap.vercel.app/gym.html?tenant=asashoku\n" +
    "　・記録ページに直行させる例：\n" +
    "　　　https://makes-momo-eap.vercel.app/check.html?tenant=asashoku\n" +
    "　※ 実際の **asashoku** は例です。必ず 01_企業マスタ の A列に書いたコードに合わせてください。\n\n" +
    "【事業所別に渡す場合】\n" +
    "　・02_事業所マスタ に、その拠点の **officeId**（A列）と **tenantId**（B列＝その企業のA列と一致）を登録したうえで、\n" +
    "　　　…/gym.html?tenant=asashoku&office=osaka\n" +
    "　のように **&office=（officeId）** を足す。osaka も例なので、マスタの A列と一致させること。\n\n" +
    "【控えの書き方】\n" +
    "　・決まったリンクは、01_企業マスタの **L列 配布用URLメモ** に貼っておくと、後から迷いにくいです。\n\n" +
    "【手順4】従業員にリンクを渡す\n" +
    "　4-1. メール本文・社内ポータル・QRコードなど、**企業（と事業所）ごとにリンクが混ざらないよう**配布する。\n" +
    "　4-2. **全員同じ汎用URL**（?tenant= なし）だけを配ると、tenantId が空のままになり、**企業別に自動では分かれません**。必ず企業ごとに分けること。\n\n" +
    "【手順5】従業員側で一度ブラウザを開く（ここで「紐づけ」が完了）\n" +
    "　5-1. 従業員は、**自分の会社用のリンクをタップしてページを開く**。\n" +
    "　5-2. 初回に ?tenant= が付いていれば、端末内に保存され、**同じ端末では次回からURLを付けなくても**同じ企業コードで送られることが多いです。\n" +
    "　5-3. 別の端末・別ブラウザ・プライベート閲覧・キャッシュ削除後は、**もう一度会社用URLから開き直す**と確実です。\n\n" +
    "【手順6】スプレッドシートで確認する\n" +
    "　6-1. 「予約一覧」「記録一覧」「脈拍記録」「睡眠記録」の**右の方の列**に tenantId（と officeId）があります。\n" +
    "　6-2. データメニューからフィルタをかけて、tenantId で絞り込むと**その企業の行だけ**が表示されます。\n" +
    "　6-3. 睡眠の月次は「睡眠_月次集計」タブ。企業表示名・事業所表示名はマスタから自動で引きます。\n" +
    "　6-4. officeId を付けて送っている場合、**支_（事業所コード）_予約** などのタブに同じ行がコピーされます（事業所単位の確認用）。\n\n" +
    "【よくある誤解】\n" +
    "・予約フォームの「所属企業名」は利用者の入力です。**集計で企業を機械的に分ける主なキーは tenantId** です（名前の表記ゆれがあり得るため）。\n\n" +
    "【この説明文を更新したいとき】\n" +
    "・00_使い方 の A1 を**空にしてから**、フォーム送信や GAS の runEapWorkbookSetup を実行すると、最新の定型文が入り直します。\n" +
    "・手書きで追記した内容は消えるので、必要なら別メモ欄にコピーしてから空にしてください。\n\n" +
    "■ タブの意味\n" +
    "・予約一覧 … 相談予約。所属企業名は利用者の入力。末尾の tenantId / officeId は配布URLの ?tenant= ?office= と同じ（未設定は空）\n" +
    "・記録一覧 … コンディション記録\n" +
    "・脈拍記録 … 自律神経チェック\n" +
    "・睡眠記録 … 睡眠ログ（企業キー tenantId・事業所キー officeId）\n" +
    "・睡眠_月次集計 … 月次統計（企業×事業所別。GAS runSleepMonthlyAggregation / トリガー）\n" +
    "・01_企業マスタ … 契約企業ごとの管理表（下記）。B列「企業表示名」は睡眠_月次集計に自動転記されます\n" +
    "・02_事業所マスタ … 事業所ごとの管理。C列「事業所表示名」は睡眠_月次集計に転記。officeId ごとに「支_{officeId}_*」タブへ行が複製されます\n\n" +
    "■ 01_企業マスタの列（1行目ヘッダーと揃える）\n" +
    "A tenantId … アプリ配布用コード（英数字・一意）。例: acme / fuji2026\n" +
    "B 企業表示名 … レポート・一覧用の短い名前（必須に近い）\n" +
    "C 正式社名 … 契約書・請求との照合用\n" +
    "D 契約ステータス … 例: 契約中 / トライアル / 終了\n" +
    "E 契約開始日 … yyyy-MM-dd 推奨\n" +
    "F 契約終了日 … 空欄は継続など運用はメモ欄に\n" +
    "G 従業員数_参考 … 規模把握（概数で可）\n" +
    "H 窓口_部署 … 貴社側の人事・総務など\n" +
    "I 窓口_担当者\n" +
    "J 窓口_連絡先 … メール or 電話\n" +
    "K 月次サマリ送付先 … 企業向け匿名集計の送付メール（複数はカンマ区切り可）\n" +
    "L 配布用URLメモ … 従業員に渡すリンクの控え（?tenant= 付き）\n" +
    "M メモ … 自由記述\n\n" +
    "■ 02_事業所マスタの列\n" +
    "A officeId … 配布URLの ?office= と一致させる一意コード\n" +
    "B tenantId … 所属企業（01のA列と一致）\n" +
    "C 事業所表示名 … レポート用\n" +
    "D〜F 所在地・従業員数_参考・メモ\n\n" +
    "■ 企業・事業所別にデータを見る\n" +
    "・各データシートの tenantId / officeId 列でフィルタ\n" +
    "・または「支_{officeId}_予約」「支_{officeId}_記録」など事業所専用タブ（自動複製）\n" +
    "・予約一覧は「所属企業名」と併用可\n\n" +
    "■ タブの並べ替え・企業マスタの見た目\n" +
    "GAS で runEapWorkbookSetup を実行すると、タブ順の整理と企業マスタ1行目の強調・固定を行います。";
  if (created || String(sh.getRange(1, 1).getValue()).trim() === "") {
    sh.getRange(1, 1).setValue(text);
    sh.getRange(1, 1).setWrap(true);
    sh.setColumnWidth(1, 720);
  }
  if (created) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(1);
  }
}

function ensureEnterpriseMasterSheet_(ss) {
  var sh = ss.getSheetByName(ENTERPRISE_MASTER_SHEET_NAME);
  var created = !sh;
  if (!sh) sh = ss.insertSheet(ENTERPRISE_MASTER_SHEET_NAME);
  ensureHeader_(sh, ENTERPRISE_MASTER_HEADER);
  if (created) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(2);
  }
}

function ensureOfficeMasterSheet_(ss) {
  var sh = ss.getSheetByName(OFFICE_MASTER_SHEET_NAME);
  var created = !sh;
  if (!sh) sh = ss.insertSheet(OFFICE_MASTER_SHEET_NAME);
  ensureHeader_(sh, OFFICE_MASTER_HEADER);
  if (created) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(3);
  }
}

/** 事業所専用ミラーシート名（Google のタブ名上限に合わせて短くする） */
function sanitizeOfficeSheetKey_(officeId) {
  var t = String(officeId || "").trim();
  if (!t) return "";
  t = t.replace(/[^a-zA-Z0-9_-]/g, "_");
  t = t.replace(/_+/g, "_");
  if (t.length > 50) t = t.slice(0, 50);
  return t;
}

/** 本線シートへ追記した行を、officeId 別タブ「支_{id}_種別」にも複製（集団分析で事業所単位に絞り込みやすくする） */
function appendOfficeMirrorRow_(officeId, typeSuffix, header, row) {
  var k = sanitizeOfficeSheetKey_(officeId);
  if (!k) return;
  var base = "支_" + k + "_" + typeSuffix;
  if (base.length > 99) base = base.slice(0, 99);
  var sh = getSheet_(EAP_SPREADSHEET_ID, base);
  ensureHeader_(sh, header);
  sh.appendRow(row);
}

/** 手動実行: タブを推奨順に並べ替え */
function runEapWorkbookSetup() {
  var ss = SpreadsheetApp.openById(EAP_SPREADSHEET_ID);
  ensureGuideSheet_(ss);
  ensureEnterpriseMasterSheet_(ss);
  ensureOfficeMasterSheet_(ss);
    var order = [
    GUIDE_SHEET_NAME,
    ENTERPRISE_MASTER_SHEET_NAME,
    OFFICE_MASTER_SHEET_NAME,
    RESERVATION_SHEET_NAME,
    HEALTH_SHEET_NAME,
    PULSE_SHEET_NAME,
    SLEEP_SHEET_NAME,
    TRAINER_SHARE_SHEET_NAME,
    SLEEP_MONTHLY_SHEET_NAME
  ];
  for (var i = order.length - 1; i >= 0; i--) {
    var s = ss.getSheetByName(order[i]);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(1);
    }
  }
  var em = ss.getSheetByName(ENTERPRISE_MASTER_SHEET_NAME);
  if (em) {
    ensureHeader_(em, ENTERPRISE_MASTER_HEADER);
    em.setFrozenRows(1);
    var n = ENTERPRISE_MASTER_HEADER.length;
    em.getRange(1, 1, 1, n).setFontWeight("bold").setBackground("#e8eef8");
    em.setColumnWidth(1, 140);
    em.setColumnWidth(2, 200);
    em.setColumnWidth(3, 220);
    em.setColumnWidth(4, 120);
    em.setColumnWidth(5, 110);
    em.setColumnWidth(6, 110);
    em.setColumnWidth(7, 100);
    em.setColumnWidth(8, 120);
    em.setColumnWidth(9, 120);
    em.setColumnWidth(10, 180);
    em.setColumnWidth(11, 220);
    em.setColumnWidth(12, 280);
    em.setColumnWidth(13, 200);
  }
  var om = ss.getSheetByName(OFFICE_MASTER_SHEET_NAME);
  if (om) {
    ensureHeader_(om, OFFICE_MASTER_HEADER);
    om.setFrozenRows(1);
    var on = OFFICE_MASTER_HEADER.length;
    om.getRange(1, 1, 1, on).setFontWeight("bold").setBackground("#eef6f0");
    om.setColumnWidth(1, 140);
    om.setColumnWidth(2, 120);
    om.setColumnWidth(3, 200);
    om.setColumnWidth(4, 180);
    om.setColumnWidth(5, 110);
    om.setColumnWidth(6, 200);
  }
}

function lookupEnterpriseDisplayName_(tenantId) {
  if (!tenantId) return "";
  try {
    var sheet = getSheet_(HEALTH_SPREADSHEET_ID, ENTERPRISE_MASTER_SHEET_NAME);
    ensureHeader_(sheet, ENTERPRISE_MASTER_HEADER);
    var last = sheet.getLastRow();
    if (last < 2) return "";
    var data = sheet.getRange(2, 1, last, 2).getValues();
    var t = String(tenantId).trim();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === t) {
        return String(data[i][1] != null ? data[i][1] : "").trim();
      }
    }
  } catch (e) {}
  return "";
}

function lookupOfficeDisplayName_(tenantId, officeId) {
  if (!tenantId || !officeId) return "";
  try {
    var sheet = getSheet_(HEALTH_SPREADSHEET_ID, OFFICE_MASTER_SHEET_NAME);
    ensureHeader_(sheet, OFFICE_MASTER_HEADER);
    var last = sheet.getLastRow();
    if (last < 2) return "";
    var data = sheet.getRange(2, 1, last, 3).getValues();
    var t = String(tenantId).trim();
    var o = String(officeId).trim();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === o && String(data[i][1]).trim() === t) {
        return String(data[i][2] != null ? data[i][2] : "").trim();
      }
    }
  } catch (e) {}
  return "";
}

/**
 * 指定年月の睡眠データを tenantId × officeId 別に集計し、「睡眠_月次集計」シートへ書き込む（officeId 空は同一企業内で1グループ）。
 * 同じ「年月」の既存行は置き換え（再実行可）。
 *
 * Google Apps Script エディタから手動実行するか、トリガーで月1回
 * runSleepMonthlyAggregationLastMonth を登録してください。
 *
 * @param {number} year 例 2026
 * @param {number} month 1〜12
 */
function runSleepMonthlyAggregation(year, month) {
  const ym = yearMonthKey_(year, month);
  const sheetIn = getSheet_(HEALTH_SPREADSHEET_ID, SLEEP_SHEET_NAME);
  ensureHeader_(sheetIn, SLEEP_HEADER);
  const lastRow = sheetIn.getLastRow();
  if (lastRow < 2) {
    writeMonthlyRowsForMonth_(ym, []);
    return;
  }
  const data = sheetIn.getRange(2, 1, lastRow, SLEEP_HEADER.length).getValues();
  const groups = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rd = normalizeRecordDate_(row[2]);
    if (!rd) continue;
    var parts = rd.split("-");
    if (parts.length !== 3) continue;
    var y = parseInt(parts[0], 10);
    var m = parseInt(parts[1], 10);
    if (y !== year || m !== month) continue;
    var tid = String(row[1] != null ? row[1] : "").trim();
    var oid = String(row[8] != null ? row[8] : "").trim();
    var gkey = tid + "\t" + (oid || "__none__");
    if (!groups[gkey]) groups[gkey] = [];
    groups[gkey].push({
      sleepMinutes: parseNum_(row[3]),
      latencyMin: parseNum_(row[4]),
      wakeups: parseNum_(row[5]),
      quality: parseNum_(row[6]),
      tags: String(row[7] != null ? row[7] : "")
    });
  }
  var groupKeys = Object.keys(groups).sort();
  var outRows = [];
  var runAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  for (var j = 0; j < groupKeys.length; j++) {
    var tKey = groupKeys[j];
    var tidPart = tKey;
    var oidPart = "";
    var tab = tKey.indexOf("\t");
    if (tab >= 0) {
      tidPart = tKey.slice(0, tab);
      oidPart = tKey.slice(tab + 1);
    }
    if (oidPart === "__none__") oidPart = "";
    var rows = groups[tKey];
    var n = rows.length;
    var ok = n >= SLEEP_AGG_MIN_N;
    var sumSm = 0;
    var sumLat = 0;
    var sumW = 0;
    var sumQ = 0;
    var sumScore = 0;
    var tagCounts = { alcohol: 0, caffeine: 0, exercise: 0, meal: 0, smoking: 0 };
    for (var k = 0; k < n; k++) {
      var r = rows[k];
      var sm = clampInt_(parseNum_(r.sleepMinutes), 0, 9999);
      var lat = clampInt_(parseNum_(r.latencyMin), 0, 240);
      var w = clampInt_(parseNum_(r.wakeups), 0, 999);
      var q = clampInt_(parseNum_(r.quality), 1, 5);
      sumSm += sm;
      sumLat += lat;
      sumW += w;
      sumQ += q;
      sumScore += computeSleepScore_(q, sm, w, lat);
      var tagParts = String(r.tags || "").split(",");
      for (var t = 0; t < tagParts.length; t++) {
        var key = String(tagParts[t]).trim().toLowerCase();
        if (tagCounts.hasOwnProperty(key)) tagCounts[key]++;
      }
    }
    var avgSm = ok ? Math.round((sumSm / n) * 10) / 10 : "";
    var avgW = ok ? Math.round((sumW / n) * 10) / 10 : "";
    var avgQ = ok ? Math.round((sumQ / n) * 10) / 10 : "";
    var avgLat = ok ? Math.round((sumLat / n) * 10) / 10 : "";
    var avgScore = ok ? Math.round((sumScore / n) * 10) / 10 : "";
    function pct(tagKey) {
      if (!ok || !n) return "";
      return Math.round((tagCounts[tagKey] / n) * 1000) / 10;
    }
    outRows.push([
      ym,
      tidPart,
      oidPart,
      lookupEnterpriseDisplayName_(tidPart),
      lookupOfficeDisplayName_(tidPart, oidPart),
      n,
      ok ? "OK" : "母体不足",
      avgScore,
      avgSm,
      avgW,
      avgLat,
      pct("alcohol"),
      pct("caffeine"),
      pct("exercise"),
      pct("meal"),
      pct("smoking"),
      runAt
    ]);
  }
  writeMonthlyRowsForMonth_(ym, outRows);
}

/** 前月分を集計（トリガー用） */
function runSleepMonthlyAggregationLastMonth() {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth(); // 0-based → 前月
  if (m === 0) {
    runSleepMonthlyAggregation(y - 1, 12);
  } else {
    runSleepMonthlyAggregation(y, m);
  }
}

function writeMonthlyRowsForMonth_(yearMonth, newRows) {
  var sheet = getSheet_(HEALTH_SPREADSHEET_ID, SLEEP_MONTHLY_SHEET_NAME);
  ensureHeader_(sheet, SLEEP_MONTHLY_HEADER);
  var lastRow = sheet.getLastRow();
  var ncols = SLEEP_MONTHLY_HEADER.length;
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow, ncols).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]) === yearMonth) {
        sheet.deleteRow(i + 2);
      }
    }
  }
  for (var j = 0; j < newRows.length; j++) {
    sheet.appendRow(newRows[j]);
  }
}

function yearMonthKey_(year, month) {
  return year + "-" + (month < 10 ? "0" : "") + month;
}

function normalizeRecordDate_(v) {
  var s = String(v == null ? "" : v).trim();
  if (!s) return "";
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return "";
  return m[1] + "-" + m[2] + "-" + m[3];
}

function parseNum_(v) {
  var n = parseFloat(String(v == null ? "" : v).replace(/,/g, ""));
  return isNaN(n) ? 0 : n;
}

function clampInt_(n, min, max) {
  var x = Math.round(Number(n));
  if (isNaN(x)) x = min;
  if (x < min) x = min;
  if (x > max) x = max;
  return x;
}

function sanitizeSleepTags_(raw) {
  var allowed = { alcohol: true, caffeine: true, exercise: true, meal: true, smoking: true };
  var parts = String(raw || "").split(/[,\s]+/);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var k = String(parts[i]).trim().toLowerCase();
    if (allowed[k] && out.indexOf(k) === -1) out.push(k);
  }
  return out.join(",");
}

/** 0〜100 の簡易スコア（企業向けレポートはこの平均などのみ利用） */
function computeSleepScore_(quality, sleepMin, wakeups, latencyMin) {
  var q = clampInt_(quality, 1, 5);
  var sm = Math.max(0, Number(sleepMin) || 0);
  var w = Math.max(0, Number(wakeups) || 0);
  var lat = Math.max(0, Number(latencyMin) || 0);
  var durPart = Math.min(1, sm / 420) * 35;
  var wakPart = Math.max(0, 25 - w * 5);
  var latPart = Math.max(0, 15 - lat / 4);
  var qPart = (q / 5) * 35;
  return Math.round(Math.min(100, durPart + wakPart + latPart + qPart));
}

function getBookedPrefer1_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, COL_PREFER1, lastRow - 1, 1).getValues();
  const tz = Session.getScriptTimeZone();

  return values
    .map(function (row) {
      const v = row[0];
      if (!v) return "";
      if (Object.prototype.toString.call(v) === "[object Date]") {
        return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm");
      }
      return String(v).trim();
    })
    .filter(function (v) { return v !== ""; });
}

function jsonResponse_(obj, statusCode) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);

  // Apps Script環境によっては setHeader が使えないため、存在確認してから設定
  if (typeof output.setHeader === "function") {
    output.setHeader("Access-Control-Allow-Origin", "*");
    output.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    output.setHeader("Access-Control-Allow-Headers", "Content-Type");
    output.setHeader("Vary", "Origin");
    if (statusCode) {
      output.setHeader("X-Status-Code", String(statusCode));
    }
  }

  return output;
}
