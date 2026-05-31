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
    const result = { ok: true, store: storeRows, staff: staffRows };
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

    if (payload.action === "updatePunch")  return updatePunch(payload);
    if (payload.action === "deletePunch")  return deletePunch(payload);
    if (payload.action === "submitLeave")  return submitLeave(payload);
    if (payload.action === "approveLeave") return approveLeave(payload);
    if (payload.action === "rejectLeave")  return rejectLeave(payload);

    if (!payload.staffId || !payload.punchType) {
      return jsonResponse({ ok: false, error: "必須フィールドが不足しています" });
    }
    if (!payload.auth || !payload.auth.testMode) {
      const staffSheet = SpreadsheetApp.openById(STAFF_SS_ID).getSheets()[0];
      const staffData  = staffSheet.getDataRange().getValues();
      let pinOk = false;
      for (let i = 1; i < staffData.length; i++) {
        const row = staffData[i];
        if (String(row[COL.staffId]).trim() === String(payload.staffId).trim()) {
          const storedPin = String(row[COL.pin] || "").trim();
          pinOk = (storedPin !== "" && storedPin === String(payload.pin || "").trim());
          break;
        }
      }
      if (!pinOk) return jsonResponse({ ok: false, error: "PIN認証失敗" });
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

    // 初期管理者（代表）を追加
    sheet.appendRow(["admin", "ideanow2026", "脇田 将樹", 1, "SUPER", "全店舗", "代表"]);
    sheet.getRange(2,1,1,7).setBackground("#f0fff4");
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
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, admin }))
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

function formatDate(val) {
  if (!val) return "";
  const d = new Date(val);
  if (isNaN(d)) return String(val).trim();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
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
