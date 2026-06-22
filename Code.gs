const STAFF_SS_ID  = "1UnBwhX8AjBY_sGXNpiYg--3BB2hgh99eu18oL1uOOts";
const STORE_SS_ID  = "1Ozyzi3WqYh7HkYYKBObZr8Mvsm941BQh4XL4w_qp-90";
const PUNCH_SS_ID  = "1NDgvWa7A25If18iklSVkU1F8HnKW3XGcq6O12QmEEx0";
const SHIFT_SS_ID  = "1NOd_wCbSV22mf3nZrt10HcK_SjzAlBc0ebVF59FgOTw";
const ATTEND_SS_ID = "1XW86MhvDMv5MYVv9aYvwa0ey7g1q7JyefGrcXK9UbPA"; // IDEA NOV 勤怠管理DB

const COL = {
  staffId:0,company:1,store:2,rank:3,empType:4,active:5,
  license:6,name:7,kanaLast:8,kanaFirst:9,gender:10,birth:11,
  origin:12,school:13,joinDate:14,transfer:15,retire:16,pin:17,
};

const PUNCH_HEADERS = [
  "打刻ID","打刻日時","日付","時刻","打刻種別","打刻種別ラベル",
  "店舗名","社員番号","氏名","役職","雇用形態","認証方法","テストモード","クライアントタイムゾーン"
];

const LEAVE_HEADERS = [
  "申請ID","申請日時","スタッフID","氏名","店舗名","休暇種別","休暇種別ラベル",
  "申請日","開始時刻","終了時刻","時間数","理由","ステータス","承認者","承認日時"
];

// ═══════════════════════════════════════════════════════════════
// 📡 doGet：GETリクエストのルーティング
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // ── 管理者認証 ──
  if (action === "adminLogin")  return adminLogin(e);
  if (action === "adminList")   return getAdminList(e);

  // ── シフト管理：保存（GETパラメータ方式でCORSを回避） ──
  if (action === "saveShift")    return saveShiftData(e);
  if (action === "saveSettings") return saveShiftSettings(e);

  // ── シフト管理：読込 ──
  if (action === "loadShift")    return loadShiftData(e);
  if (action === "loadSettings") return loadShiftSettings(e);

  // ── 勤怠管理 ──
  if (action === "punchLogs") {
    const from  = e.parameter.from  || "";
    const to    = e.parameter.to    || "";
    const store = e.parameter.store || "";
    const logs  = getPunchLogs(from, to, store);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, logs }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "leaveRequests") {
    const store  = e.parameter.store  || "";
    const status = e.parameter.status || "";
    const leaves = getLeaveRequests(store, status);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, leaves }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── デフォルト：スタッフ・店舗データ取得 ──
  try {
    const staffRows = SpreadsheetApp.openById(STAFF_SS_ID)
      .getSheets()[0].getDataRange().getValues();
    const storeRows = SpreadsheetApp.openById(STORE_SS_ID)
      .getSheets()[0].getDataRange().getValues();
    ensurePinColumn(STAFF_SS_ID);
    // セキュリティ #1・#2：PIN・PIIをクライアントへ返さない。
    // 列構成は維持（既存パーサ互換）しつつ、PIN列は「設定済みか」のフラグ(1/空)に置換し、
    // 生年月日・出身地・出身校・性別は空にして配信する。
    const safeStaff = staffRows.map((row, idx) => {
      if (idx === 0) return row; // ヘッダー行はそのまま
      const r = row.slice();
      r[COL.pin]      = (String(row[COL.pin] || "").trim() !== "") ? "1" : "";
      r[COL.birth]    = "";
      r[COL.origin]   = "";
      r[COL.school]   = "";
      r[COL.gender]   = "";
      // 日付は西暦に正規化（和暦混在の解消＋有給計算の不具合防止）
      r[COL.joinDate] = formatDate(row[COL.joinDate]);
      r[COL.retire]   = row[COL.retire] ? formatDate(row[COL.retire]) : "";
      return r;
    });
    const result = { ok: true, store: storeRows, staff: safeStaff };
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════════
// 📡 doPost：POSTリクエストのルーティング（勤怠管理用）
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse({ ok: false, error: "JSON parse error" });
    }

    if (payload.action === "verifyPin")      return verifyPinAction(payload);
    if (payload.action === "staffDirectory") return staffDirectory(payload);
    if (payload.action === "updatePunch")  return updatePunch(payload);
    if (payload.action === "deletePunch")  return deletePunch(payload);
    if (payload.action === "submitLeave")  return submitLeave(payload);
    if (payload.action === "approveLeave") return approveLeave(payload);
    if (payload.action === "rejectLeave")  return rejectLeave(payload);

    // ── 管理者・スタッフ管理（設定画面） ──
    if (payload.action === "addAdmin")     return addAdmin(payload);
    if (payload.action === "deleteAdmin")  return deleteAdmin(payload);
    if (payload.action === "addStaff")     return addStaff(payload);
    if (payload.action === "retireStaff")  return retireStaff(payload);

    if (!payload.staffId || !payload.punchType) {
      return jsonResponse({ ok: false, error: "必須フィールドが不足しています" });
    }
    // 打刻はサーバ側で必ず PIN照合＋在職確認を行う
    // （#3: testModeによる認証バイパスを廃止／#5: 在職チェックをサーバ側に追加）
    {
      const staffSheet = SpreadsheetApp.openById(STAFF_SS_ID).getSheets()[0];
      const staffData  = staffSheet.getDataRange().getValues();
      let found = false, pinOk = false, employed = false;
      for (let i = 1; i < staffData.length; i++) {
        const row = staffData[i];
        if (String(row[COL.staffId]).trim() === String(payload.staffId).trim()) {
          found = true;
          const storedPin = String(row[COL.pin] || "").trim();
          pinOk = (storedPin !== "" && storedPin === String(payload.pin || "").trim());
          const active = String(row[COL.active] || "").trim();
          const retire = String(row[COL.retire] || "").trim();
          const inactive = (active === "否" || active === "退職" || active === "×" || active === "0") ||
                           (retire !== "" && retire !== "0" && retire !== "-");
          employed = !inactive;
          break;
        }
      }
      if (!found)    return jsonResponse({ ok: false, error: "社員番号が見つかりません" });
      if (!employed) return jsonResponse({ ok: false, error: "在職中のスタッフではありません" });
      if (!pinOk)    return jsonResponse({ ok: false, error: "PIN認証失敗" });
    }
    const punchSheet = ensurePunchSheet();
    const existingData = punchSheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0] === payload.punchId) {
        return jsonResponse({ ok: false, error: "重複打刻（同一punchId）" });
      }
    }
    const rec = payload.recordedAt || {};
    const newRow = [
      payload.punchId,
      rec.localIso   || new Date().toISOString(),
      (rec.minute    || "").split(" ")[0] || "",
      (rec.minute    || "").split(" ")[1] || "",
      payload.punchType,
      payload.punchTypeLabel || "",
      payload.storeName      || "",
      payload.staffId        || "",
      payload.staffName      || "",
      payload.staffRank      || "",
      payload.staffEmpType   || "",
      (payload.auth && payload.auth.method)   || "pin",
      (payload.auth && payload.auth.testMode) ? "TRUE" : "FALSE",
      rec.tz || "",
    ];
    punchSheet.appendRow(newRow);
    return jsonResponse({ ok: true, message: "打刻を記録しました", punchId: payload.punchId });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// 🔐 PIN事前照合（打刻画面用 / #1・#2：PINをクライアントへ渡さずサーバ照合）
// payload: { action:"verifyPin", staffId, pin }
// 返却: { ok, verified, employed, pinSet }
// ═══════════════════════════════════════════════════════════════
function verifyPinAction(payload) {
  const data = SpreadsheetApp.openById(STAFF_SS_ID).getSheets()[0].getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[COL.staffId]).trim() === String(payload.staffId).trim()) {
      const storedPin = String(row[COL.pin] || "").trim();
      const active    = String(row[COL.active] || "").trim();
      const retire    = String(row[COL.retire] || "").trim();
      const employed  = !((active === "否" || active === "退職" || active === "×" || active === "0") ||
                          (retire !== "" && retire !== "0" && retire !== "-"));
      const verified  = storedPin !== "" && storedPin === String(payload.pin || "").trim();
      return jsonResponse({ ok: true, verified, employed, pinSet: storedPin !== "" });
    }
  }
  return jsonResponse({ ok: true, verified: false, employed: false, pinSet: false, error: "社員番号が見つかりません" });
}

// ═══════════════════════════════════════════════════════════════
// 🎟️ 管理者セッショントークン（個人情報など保護APIの認可）
// ═══════════════════════════════════════════════════════════════
// ログイン時に発行し CacheService に6時間保持。保護APIはトークン必須。
function makeAdminToken(admin) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache()
    .put("adm_" + token, JSON.stringify({ id: admin.id, level: Number(admin.level) }), 21600);
  return token;
}
// 有効なら {id,level} を返す。minLevel 指定時はその数値以下（=同等以上の権限）のみ許可。
function checkAdminToken(token, minLevel) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get("adm_" + String(token));
  if (!raw) return null;
  let a; try { a = JSON.parse(raw); } catch (e) { return null; }
  if (minLevel && Number(a.level) > minLevel) return null;
  return a;
}

// 人事労務：スタッフ個人情報の取得（#1。保護API＝SUPER/EXEC のみ・トークン必須）
function staffDirectory(payload) {
  const auth = checkAdminToken(payload.token, 2);
  if (!auth) return jsonResponse({ ok: false, error: "権限がありません（再ログインが必要な場合があります）" });
  const rows = SpreadsheetApp.openById(STAFF_SS_ID).getSheets()[0].getDataRange().getValues();
  const staff = rows.slice(1)
    .filter(r => String(r[COL.name] || "").trim() && r[COL.name] !== "氏名")
    .map(r => ({
      staffId:  r[COL.staffId], company: r[COL.company], store: r[COL.store],
      rank:     r[COL.rank],    empType: r[COL.empType], active: r[COL.active],
      name:     r[COL.name],    kanaLast: r[COL.kanaLast], kanaFirst: r[COL.kanaFirst],
      gender:   r[COL.gender],  birth: formatDate(r[COL.birth]), origin: r[COL.origin],
      school:   r[COL.school],  joinDate: formatDate(r[COL.joinDate]),
      retire:   r[COL.retire] ? formatDate(r[COL.retire]) : "",
      pinSet:   String(r[COL.pin] || "").trim() !== "",
    }));
  return jsonResponse({ ok: true, staff });
}

// ═══════════════════════════════════════════════════════════════
// 👤 管理者認証（IDEA NOV 勤怠管理DB / admin_master シート）
// ═══════════════════════════════════════════════════════════════

// 管理者マスタシートを自動作成
function ensureAdminSheet() {
  const ss    = SpreadsheetApp.openById(ATTEND_SS_ID);
  let sheet   = ss.getSheetByName("admin_master");
  if (!sheet) {
    sheet = ss.insertSheet("admin_master");
    const headers = ["管理者ID","パスワード","氏名","権限レベル","権限名","担当店舗","メモ"];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#2d3748").setFontColor("#ffffff");
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 80);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 200);
    sheet.setColumnWidth(7, 200);

    // セキュリティ #4：初期パスワードをコードに埋め込まない。
    // 初回作成時は SUPER 管理者の「枠」だけ作り、パスワードは手動でシートに設定する。
    // （公開リポジトリに既定の管理者ID/パスワードが残らないようにするため）
    sheet.appendRow(["admin", "", "脇田 将樹", 1, "SUPER", "全店舗", "初回：admin_masterシートにパスワードを手入力してください"]);
    sheet.getRange(2,1,1,7).setBackground("#fff5f5");
  }
  return sheet;
}

// 管理者ログイン
function adminLogin(e) {
  try {
    const adminId = e.parameter.adminId || "";
    const pass    = e.parameter.pass    || "";

    if (!adminId || !pass) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: "IDまたはパスワードが空です" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ensureAdminSheet();
    const rows  = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (String(row[0]).trim() === adminId && String(row[1]).trim() === pass) {
        const admin = {
          id:       row[0],
          name:     row[2],
          level:    Number(row[3]),
          role:     row[4],
          stores:   String(row[5]).trim(), // "全店舗" or "BASSA久米川店,BASSA新所沢店"
          memo:     row[6],
        };
        // セッショントークンを発行（個人情報など保護APIの認可に使用）
        const token = makeAdminToken(admin);
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, admin, token }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "IDまたはパスワードが違います" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 管理者一覧取得（SUPER権限のみ）
function getAdminList(e) {
  try {
    const sheet = ensureAdminSheet();
    const rows  = sheet.getDataRange().getValues();
    const admins = rows.slice(1).map(row => ({
      id:     row[0],
      name:   row[2],
      level:  Number(row[3]),
      role:   row[4],
      stores: row[5],
      memo:   row[6],
    }));
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, admins }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 管理者を新規登録（設定画面 / #5・#6）
// payload: { adminId, pass, name, level, role, stores, memo }
// stores = 担当店舗（カンマ区切り）＝この管理者が承認できる店舗の紐づけ（#6）
function addAdmin(payload) {
  const sheet = ensureAdminSheet();
  const rows  = sheet.getDataRange().getValues();
  const id    = String(payload.adminId || "").trim();
  if (!id || !String(payload.pass || "").trim()) {
    return jsonResponse({ ok: false, error: "管理者IDとパスワードは必須です" });
  }
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === id) {
      return jsonResponse({ ok: false, error: "同じ管理者IDが既に存在します" });
    }
  }
  sheet.appendRow([
    id,
    String(payload.pass),
    payload.name   || "",
    Number(payload.level) || 5,
    payload.role   || "STORE",
    payload.stores || "",
    payload.memo   || "",
  ]);
  return jsonResponse({ ok: true, message: "管理者を登録しました" });
}

// 管理者を削除（設定画面 / #5）
function deleteAdmin(payload) {
  const id = String(payload.adminId || "").trim();
  if (id === "admin") {
    return jsonResponse({ ok: false, error: "初期管理者（admin）は削除できません" });
  }
  const sheet = ensureAdminSheet();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === id) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true, message: "管理者を削除しました" });
    }
  }
  return jsonResponse({ ok: false, error: "対象の管理者が見つかりません" });
}

// スタッフを新規登録（設定画面 / #9）
// payload: { staffId, company, store, rank, empType, name, kanaLast, kanaFirst, birth, joinDate, pin }
function addStaff(payload) {
  const sheet   = SpreadsheetApp.openById(STAFF_SS_ID).getSheets()[0];
  const rows    = sheet.getDataRange().getValues();
  const staffId = String(payload.staffId || "").trim();
  if (!staffId || !String(payload.name || "").trim()) {
    return jsonResponse({ ok: false, error: "社員番号と氏名は必須です" });
  }
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.staffId]).trim() === staffId) {
      return jsonResponse({ ok: false, error: "同じ社員番号が既に存在します" });
    }
  }
  const width = Math.max(sheet.getLastColumn(), COL.pin + 1);
  const row   = new Array(width).fill("");
  row[COL.staffId]   = staffId;
  row[COL.company]   = payload.company   || "";
  row[COL.store]     = payload.store     || "";
  row[COL.rank]      = payload.rank      || "";
  row[COL.empType]   = payload.empType   || "";
  row[COL.active]    = "在職";
  row[COL.name]      = payload.name      || "";
  row[COL.kanaLast]  = payload.kanaLast  || "";
  row[COL.kanaFirst] = payload.kanaFirst || "";
  row[COL.birth]     = payload.birth     || "";
  row[COL.joinDate]  = payload.joinDate  || "";
  row[COL.pin]       = String(payload.pin || "");
  sheet.appendRow(row);
  return jsonResponse({ ok: true, message: "スタッフを登録しました" });
}

// スタッフを退職処理（設定画面 / #9）
// 行は削除せず在職列を「退職」にして退職日を記録（履歴・打刻ログ保全のためソフト削除）
function retireStaff(payload) {
  const sheet   = SpreadsheetApp.openById(STAFF_SS_ID).getSheets()[0];
  const rows    = sheet.getDataRange().getValues();
  const staffId = String(payload.staffId || "").trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][COL.staffId]).trim() === staffId) {
      sheet.getRange(i + 1, COL.active + 1).setValue("退職");
      sheet.getRange(i + 1, COL.retire + 1).setValue(payload.retireDate || formatDate(new Date()));
      return jsonResponse({ ok: true, message: "退職処理を行いました" });
    }
  }
  return jsonResponse({ ok: false, error: "対象スタッフが見つかりません" });
}

// ═══════════════════════════════════════════════════════════════
// 📅 シフト管理：データ保存・読込（GETパラメータ方式）
// ═══════════════════════════════════════════════════════════════

// シフトデータ保存
function saveShiftData(e) {
  try {
    const p         = e.parameter;
    const storeId   = p.storeId   || "";
    const storeName = p.storeName || "";
    const year      = p.year      || "";
    const month     = p.month     || "";
    const cells     = JSON.parse(p.cells || "{}");

    const ss        = SpreadsheetApp.openById(SHIFT_SS_ID);
    const sheetName = `shift_${year}_${String(month).padStart(2,"0")}`;
    let sheet       = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["store_id","store_name","cell_key","stamp","updated_at"]);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#e8f4fd");
    }

    // 既存の該当店舗データを削除
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(storeId)) rowsToDelete.push(i + 1);
    }
    rowsToDelete.forEach(r => sheet.deleteRow(r));

    // 新しいデータを追記
    const now  = new Date().toISOString();
    const rows = Object.entries(cells).map(([key, stamp]) => [storeId, storeName, key, stamp, now]);
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }

    return jsonResponse({ ok: true, message: `保存しました（${rows.length}件）` }, e);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, e);
  }
}

// シフトデータ読込
function loadShiftData(e) {
  try {
    const storeId   = e.parameter.storeId || "";
    const year      = e.parameter.year    || "";
    const month     = e.parameter.month   || "";
    const ss        = SpreadsheetApp.openById(SHIFT_SS_ID);
    const sheetName = `shift_${year}_${String(month).padStart(2,"0")}`;
    const sheet     = ss.getSheetByName(sheetName);

    if (!sheet) {
      return jsonResponse({ ok: true, cells: {} }, e);
    }

    const data  = sheet.getDataRange().getValues();
    const cells = {};
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(storeId)) {
        cells[data[i][2]] = data[i][3];
      }
    }
    return jsonResponse({ ok: true, cells }, e);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, e);
  }
}

// 店舗設定保存
function saveShiftSettings(e) {
  try {
    const p           = e.parameter;
    const storeId     = p.storeId     || "";
    const storeName   = p.storeName   || "";
    const minStaff    = JSON.parse(p.minStaff     || "{}");
    const storeSets   = JSON.parse(p.storeSettings|| "{}");
    const staffConfig = JSON.parse(p.staffConfig  || "{}");
    const extraRules  = JSON.parse(p.extraRules   || "{}");

    const ss    = SpreadsheetApp.openById(SHIFT_SS_ID);
    let sheet   = ss.getSheetByName("shift_settings");
    if (!sheet) {
      sheet = ss.insertSheet("shift_settings");
      sheet.appendRow(["store_id","store_name","setting_key","setting_value","updated_at"]);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#e8f5e9");
    }

    // 既存の該当店舗設定を削除
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(storeId)) rowsToDelete.push(i + 1);
    }
    rowsToDelete.forEach(r => sheet.deleteRow(r));

    const now  = new Date().toISOString();
    const rows = [
      [storeId, storeName, "minStaff",      JSON.stringify(minStaff),    now],
      [storeId, storeName, "storeSettings", JSON.stringify(storeSets),   now],
      [storeId, storeName, "staffConfig",   JSON.stringify(staffConfig), now],
      [storeId, storeName, "extraRules",    JSON.stringify(extraRules),  now],
    ];
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);

    return jsonResponse({ ok: true, message: "設定を保存しました" }, e);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, e);
  }
}

// 店舗設定読込
function loadShiftSettings(e) {
  try {
    const storeId = e.parameter.storeId || "";
    const ss      = SpreadsheetApp.openById(SHIFT_SS_ID);
    const sheet   = ss.getSheetByName("shift_settings");

    if (!sheet) {
      return jsonResponse({ ok: true, settings: {} }, e);
    }

    const data     = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(storeId)) {
        try { settings[data[i][2]] = JSON.parse(data[i][3]); }
        catch(ex) { settings[data[i][2]] = data[i][3]; }
      }
    }
    return jsonResponse({ ok: true, settings }, e);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, e);
  }
}

// ═══════════════════════════════════════════════════════════════
// 🕐 勤怠管理：打刻・休暇申請
// ═══════════════════════════════════════════════════════════════

function updatePunch(payload) {
  const sheet = SpreadsheetApp.openById(PUNCH_SS_ID).getSheetByName("打刻ログ");
  if (!sheet) return jsonResponse({ ok: false, error: "シートが見つかりません" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(payload.punchId)) {
      sheet.getRange(i + 1, 3).setValue(payload.date || rows[i][2]);
      sheet.getRange(i + 1, 4).setValue(payload.time || rows[i][3]);
      sheet.getRange(i + 1, 5).setValue(payload.punchType || rows[i][4]);
      sheet.getRange(i + 1, 6).setValue(payload.punchTypeLabel || rows[i][5]);
      return jsonResponse({ ok: true, message: "打刻を修正しました" });
    }
  }
  return jsonResponse({ ok: false, error: "対象の打刻が見つかりません" });
}

function deletePunch(payload) {
  const sheet = SpreadsheetApp.openById(PUNCH_SS_ID).getSheetByName("打刻ログ");
  if (!sheet) return jsonResponse({ ok: false, error: "シートが見つかりません" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(payload.punchId)) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ ok: true, message: "打刻を削除しました" });
    }
  }
  return jsonResponse({ ok: false, error: "対象の打刻が見つかりません" });
}

function submitLeave(payload) {
  const sheet = ensureLeaveSheet();
  const leaveId = "leave-" + new Date().getTime() + "-" + Math.random().toString(36).slice(2,6);
  const newRow = [
    leaveId, new Date().toISOString(),
    payload.staffId || "", payload.staffName || "", payload.storeName || "",
    payload.leaveType || "", payload.leaveLabel || "", payload.leaveDate || "",
    payload.timeFrom || "", payload.timeTo || "", payload.hours || 0,
    payload.reason || "", "pending", "", "",
  ];
  sheet.appendRow(newRow);
  return jsonResponse({ ok: true, message: "申請を受け付けました", leaveId });
}

function approveLeave(payload) {
  return updateLeaveStatus(payload.leaveId, "approved", payload.approvedBy || "管理者");
}

function rejectLeave(payload) {
  return updateLeaveStatus(payload.leaveId, "rejected", payload.approvedBy || "管理者");
}

function updateLeaveStatus(leaveId, status, approvedBy) {
  const sheet = SpreadsheetApp.openById(PUNCH_SS_ID).getSheetByName("休暇申請");
  if (!sheet) return jsonResponse({ ok: false, error: "シートが見つかりません" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(leaveId)) {
      sheet.getRange(i + 1, 13).setValue(status);
      sheet.getRange(i + 1, 14).setValue(approvedBy);
      sheet.getRange(i + 1, 15).setValue(new Date().toISOString());
      return jsonResponse({ ok: true, message: `申請を${status === "approved" ? "承認" : "却下"}しました` });
    }
  }
  return jsonResponse({ ok: false, error: "対象の申請が見つかりません" });
}

function getLeaveRequests(storeName, status) {
  const sheet = SpreadsheetApp.openById(PUNCH_SS_ID).getSheetByName("休暇申請");
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(row => {
      const storeOk  = !storeName || String(row[4]).includes(storeName);
      const statusOk = !status    || String(row[12]) === status;
      return storeOk && statusOk;
    })
    .map(row => ({
      leaveId: row[0], appliedAt: row[1], staffId: row[2], staffName: row[3],
      storeName: row[4], leaveType: row[5], leaveLabel: row[6], leaveDate: row[7],
      timeFrom: row[8], timeTo: row[9], hours: row[10], reason: row[11],
      status: row[12], approvedBy: row[13], approvedAt: row[14],
    }));
}

function ensureLeaveSheet() {
  const ss  = SpreadsheetApp.openById(PUNCH_SS_ID);
  let sheet = ss.getSheetByName("休暇申請");
  if (!sheet) {
    sheet = ss.insertSheet("休暇申請");
    sheet.appendRow(LEAVE_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,LEAVE_HEADERS.length).setFontWeight("bold").setBackground("#e8f5e9");
  }
  return sheet;
}

function formatTime(val) {
  if (val instanceof Date) {
    return `${String(val.getHours()).padStart(2,"0")}:${String(val.getMinutes()).padStart(2,"0")}`;
  }
  const s = String(val || "").trim();
  if (s.includes("T")) {
    const d = new Date(s);
    if (!isNaN(d)) return `${String(d.getHours()).padStart(2,"0")}:${String(d.getMinutes()).padStart(2,"0")}`;
  }
  return s;
}

function ymd_(d) {
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

// 日付を西暦 YYYY-MM-DD に正規化。和暦（漢字／S・H・R等の略記）も西暦へ変換する。
function formatDate(val) {
  if (val === null || val === undefined || val === "") return "";
  if (val instanceof Date) return isNaN(val) ? "" : ymd_(val);

  let s = String(val).trim();
  if (s === "" || s === "-" || s === "0") return "";
  // 全角数字 → 半角
  s = s.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

  // 和暦（元号 + 年[ + 月日]）。base + 年 = 西暦
  const ERAS = [
    { re: /(?:令和|令|R)\s*(元|\d+)/, base: 2018 },
    { re: /(?:平成|平|H)\s*(元|\d+)/, base: 1988 },
    { re: /(?:昭和|昭|S)\s*(元|\d+)/, base: 1925 },
    { re: /(?:大正|大|T)\s*(元|\d+)/, base: 1911 },
    { re: /(?:明治|明|M)\s*(元|\d+)/, base: 1867 },
  ];
  for (const e of ERAS) {
    const m = s.match(e.re);
    if (m) {
      const yr   = (m[1] === "元") ? 1 : Number(m[1]);
      const year = e.base + yr;
      const md   = s.match(/年\s*(\d{1,2})\s*月\s*(\d{1,2})/) ||
                   s.match(/[.\/\-]\s*(\d{1,2})\s*[.\/\-]\s*(\d{1,2})/);
      const mo = md ? Number(md[1]) : 1;
      const dd = md ? Number(md[2]) : 1;
      return `${year}-${String(mo).padStart(2,"0")}-${String(dd).padStart(2,"0")}`;
    }
  }

  // 西暦（Date解析可能なもの）
  const dt = new Date(s);
  if (!isNaN(dt)) return ymd_(dt);
  return s; // 解析不能は原文のまま
}

function getPunchLogs(fromDate, toDate, storeName) {
  const ss    = SpreadsheetApp.openById(PUNCH_SS_ID);
  const sheet = ss.getSheetByName("打刻ログ");
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  return rows.slice(1)
    .filter(row => {
      const dateStr = formatDate(row[2]);
      const dateOk  = (!fromDate || dateStr >= fromDate) && (!toDate || dateStr <= toDate);
      const storeOk = !storeName || String(row[6] || "").includes(storeName);
      return dateOk && storeOk;
    })
    .map(row => ({
      punchId: row[0], datetime: row[1], date: formatDate(row[2]), time: formatTime(row[3]),
      punchType: row[4], typeLabel: row[5], storeName: row[6],
      staffId: row[7], staffName: row[8], staffRank: row[9], empType: row[10],
    }));
}

function ensurePunchSheet() {
  const ss  = SpreadsheetApp.openById(PUNCH_SS_ID);
  let sheet = ss.getSheetByName("打刻ログ");
  if (!sheet) {
    sheet = ss.insertSheet("打刻ログ");
    sheet.appendRow(PUNCH_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,PUNCH_HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

function ensurePinColumn(ssId) {
  const sheet   = SpreadsheetApp.openById(ssId).getSheets()[0];
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const pinIdx  = headers.findIndex(h => String(h).trim() === "PIN");
  if (pinIdx === -1) {
    sheet.getRange(1, COL.pin + 1).setValue("PIN").setBackground("#fff9c4");
  }
}

function jsonResponse(obj, e) {
  const cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(`${cb}(${JSON.stringify(obj)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
