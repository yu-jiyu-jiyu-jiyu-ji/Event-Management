/**
 * うおの会 イベント管理システム — GAS Web API 基盤
 *
 * 【Script Properties に設定する値】
 * - SPREADSHEET_ID          : 対象スプレッドシート ID
 * - LINE_CHANNEL_ACCESS_TOKEN : LINE Messaging API チャネルアクセストークン
 * - LINE_ADMIN_USER_ID      : 通知先 LINE userId（魚谷さん等）
 *
 * 【スプレッドシート】
 * - シート名 master_customers  : 顧客マスタ
 * - シート名 event_histories    : イベント参加履歴
 *
 * 【重要】スプレッドシート → 拡張機能 → Apps Script に Code.js をすべて貼り付け
 * 初回は runFullFormSync() を1回実行
 */

// =============================================================================
// 定数
// =============================================================================

const SHEET_MASTER = "master_customers";
const SHEET_EVENTS = "event_histories";
const SHEET_RECEIPT_TEMPLATE = "receipt_template";
const SHEET_EVENT_SETTINGS = "event_settings";

/**
 * 領収書ひな型（receipt_template）の入力セル（A1形式）。
 * スプレッドシートのレイアウトを変えた場合はここだけ調整してください。
 */
const RECEIPT_TEMPLATE_CELLS = {
  eventDate: "U7",
  addressee: "E8",
  amount: "E11",
};

/** @enum {string} */
const CustomerStatus = {
  ACTIVE: "active",
  PENDING_MATCH: "pending_match",
};

/** @enum {string} */
const MatchPattern = {
  EXISTING_EMAIL: "existing_email", // パターンA
  PENDING_NAME: "pending_name", // パターンB
  NEW_CUSTOMER: "new_customer", // パターンC
};

const MASTER_SCHEMA = [
  { key: "line_user_id", label: "LINEユーザーID" },
  { key: "line_display_name", label: "LINE表示名" },
  { key: "user_name", label: "氏名" },
  { key: "company_name", label: "会社名" },
  { key: "position_category", label: "区分" },
  { key: "position_name", label: "役職" },
  { key: "email", label: "メールアドレス" },
  { key: "phone_number", label: "電話番号" },
  { key: "referrer", label: "紹介者" },
  { key: "status", label: "ステータス" },
  { key: "created_at", label: "登録日時" },
];

const EVENT_SCHEMA = [
  { key: "timestamp", label: "記録日時" },
  { key: "event_date", label: "イベント日" },
  { key: "user_name", label: "氏名" },
  { key: "name_reading", label: "読み方" },
  { key: "form_email", label: "申込メール" },
  { key: "receipt_required", label: "領収書" },
  { key: "receipt_name", label: "領収書宛名" },
  { key: "line_user_id", label: "LINEユーザーID" },
  { key: "payment_status", label: "支払い状況" },
  { key: "attendance_status", label: "来場状況" },
  { key: "receipt_dl_count", label: "領収書DL数" },
];

const MASTER_HEADERS = MASTER_SCHEMA.map((col) => col.label);
const EVENT_HEADERS = EVENT_SCHEMA.map((col) => col.label);

/** 日本語ヘッダー・旧英語ヘッダー → 内部キー */
const COLUMN_LABEL_TO_KEY_ = (function buildColumnMaps_() {
  const map = {};
  MASTER_SCHEMA.concat(EVENT_SCHEMA).forEach((col) => {
    map[col.label] = col.key;
    map[col.key] = col.key;
  });
  return map;
})();

function getSchemaForSheet_(sheet) {
  const name = sheet.getName();
  if (name === SHEET_EVENTS) {
    return EVENT_SCHEMA;
  }
  if (name === SHEET_MASTER) {
    return MASTER_SCHEMA;
  }
  return null;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} key
 * @param {Object[]|null} [schemaOpt]
 * @returns {number|undefined}
 */
function getColumnIndexForKey_(sheet, key, schemaOpt) {
  const headerMap = getHeaderMap_(sheet);
  const schema =
    schemaOpt ||
    getSchemaForSheet_(sheet) ||
    MASTER_SCHEMA.concat(EVENT_SCHEMA);
  const schemaCol = schema.find((c) => c.key === key);
  if (schemaCol && headerMap[schemaCol.label] !== undefined) {
    return headerMap[schemaCol.label];
  }
  if (headerMap[key] !== undefined) {
    return headerMap[key];
  }
  return undefined;
}

// =============================================================================
// HTTP エントリポイント
// =============================================================================

/**
 * Web アプリ POST 受信（Googleフォーム連携・フロントエンド API 共通）
 * @param {GoogleAppsScript.Events.DoPost} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  try {
    const payload = parseRequestBody_(e);
    const action = payload.action || "event_application";

    let result;
    switch (action) {
      case "event_application":
        result = handleEventApplication_(payload);
        break;
      case "checkin_lookup":
        result = handleCheckinLookup_(payload);
        break;
      case "checkin_link_by_name":
        result = handleCheckinLinkByName_(payload);
        break;
      case "checkin_link_email":
        result = handleCheckinLinkEmail_(payload);
        break;
      case "checkin_complete":
        result = handleCheckinComplete_(payload);
        break;
      case "receipt_list":
        result = handleReceiptList_(payload);
        break;
      case "receipt_lookup":
        result = handleReceiptLookup_(payload);
        break;
      case "receipt_generate":
        result = handleReceiptGenerate_(payload);
        break;
      case "health":
        result = { ok: true, message: "うおの会 GAS API is running." };
        break;
      default:
        throw new Error(`未対応の action です: ${action}`);
    }

    return jsonResponse_(200, result);
  } catch (error) {
    console.error(error);
    return jsonResponse_(400, {
      ok: false,
      error: error.message || String(error),
    });
  }
}

/**
 * GET でも API を呼べるように（Vercel → GAS の CORS 回避用）
 * @param {GoogleAppsScript.Events.DoGet} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doGet(e) {
  const action = e.parameter.action;
  if (!action) {
    return jsonResponse_(200, {
      ok: true,
      service: "uo-kai-event-api",
      version: "1.1.0",
      timestamp: new Date().toISOString(),
    });
  }

  try {
    let payload = {};
    if (e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    } else {
      payload = Object.assign({}, e.parameter);
      delete payload.action;
    }
    payload.action = action;

    let result;
    switch (action) {
      case "checkin_lookup":
        result = handleCheckinLookup_(payload);
        break;
      case "checkin_link_by_name":
        result = handleCheckinLinkByName_(payload);
        break;
      case "checkin_link_email":
        result = handleCheckinLinkEmail_(payload);
        break;
      case "checkin_complete":
        result = handleCheckinComplete_(payload);
        break;
      case "receipt_list":
        result = handleReceiptList_(payload);
        break;
      case "receipt_lookup":
        result = handleReceiptLookup_(payload);
        break;
      case "receipt_generate":
        result = handleReceiptGenerate_(payload);
        break;
      case "health":
        result = { ok: true, message: "うおの会 GAS API is running." };
        break;
      default:
        throw new Error(`GET 未対応の action です: ${action}`);
    }

    return jsonResponse_(200, result);
  } catch (error) {
    console.error(error);
    return jsonResponse_(400, {
      ok: false,
      error: error.message || String(error),
    });
  }
}

// =============================================================================
// 当日受付 LIFF API
// =============================================================================

/**
 * LINE ID で顧客マスタを検索
 * @param {Object} payload
 * @returns {Object}
 */
function handleCheckinLookup_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventDate = requireString_(payload.event_date, "event_date");

  const masterSheet = getSheet_(SHEET_MASTER);
  const eventSheet = getSheet_(SHEET_EVENTS);

  const masterIndex = loadMasterIndex_(masterSheet);
  const customer = masterIndex.byLineId[lineUserId] || null;

  if (!customer) {
    return {
      ok: true,
      found: false,
      needsName: true,
    };
  }

  const eventRows = loadEventRowsForDate_(eventSheet, eventDate);
  const eventHistory = findEventHistoryForCheckinFromRows_(eventRows, {
    eventDate,
    lineUserId,
    email: customer.email,
  });

  const lineDisplayName = sanitizeLineDisplayName_(payload.line_display_name || "");
  if (lineDisplayName) {
    updateMasterFields_(masterSheet, customer._rowIndex, {
      line_display_name: lineDisplayName,
    });
    customer.line_display_name = lineDisplayName;
  }

  return {
    ok: true,
    found: true,
    needsName: false,
    autoIdentified: true,
    customer: serializeCustomer_(customer),
    eventHistory: eventHistory ? serializeEvent_(eventHistory) : null,
    alreadyCheckedIn:
      eventHistory && String(eventHistory.attendance_status) === "済",
  };
}

/**
 * 当日参加者リストから氏名で検索し LINE ID を紐付け
 * 同姓同名が複数いる場合のみメール確認を要求
 * @param {Object} payload
 * @returns {Object}
 */
function handleCheckinLinkByName_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventDate = requireString_(payload.event_date, "event_date");
  const userName = sanitizeUserName_(payload.user_name);
  if (!userName) {
    throw new Error("氏名を入力してください。");
  }

  const emailInput = String(payload.email || "").trim();
  const masterSheet = getSheet_(SHEET_MASTER);
  const eventSheet = getSheet_(SHEET_EVENTS);
  const masterIndex = loadMasterIndex_(masterSheet);

  let matches = findEventParticipantsByName_(
    eventSheet,
    masterIndex,
    eventDate,
    userName,
  );

  if (matches.length === 0) {
    throw new Error(
      "本日の参加者リストに見つかりません。お申込み時の氏名をご確認ください。",
    );
  }

  const pendingMatches = matches.filter(
    (m) => String(m.event.attendance_status) !== "済",
  );
  matches = pendingMatches.length > 0 ? pendingMatches : matches;

  if (matches.length > 1 && !emailInput) {
    return {
      ok: true,
      linked: false,
      needsEmailConfirm: true,
      matchCount: matches.length,
      userName,
      message:
        "同姓同名の方がいらっしゃいます。お申込み時のメールアドレスを入力してください。",
    };
  }

  let selected;
  if (matches.length === 1) {
    selected = matches[0];
  } else {
    const email = sanitizeEmail_(emailInput);
    selected = matches.find(
      (m) => sanitizeEmailSafe_(m.email) === email,
    );
    if (!selected) {
      throw new Error("氏名とメールアドレスの組み合わせが見つかりません。");
    }
  }

  return linkCheckinParticipant_(
    lineUserId,
    eventDate,
    selected,
    userName,
    sanitizeLineDisplayName_(payload.line_display_name || ""),
  );
}

/**
 * 当日 event_histories から氏名一致の参加者を検索（一括読み込み）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventSheet
 * @param {Object} masterIndex
 * @param {string} eventDate
 * @param {string} userName
 * @returns {Object[]}
 */
function findEventParticipantsByName_(eventSheet, masterIndex, eventDate, userName) {
  const targetName = sanitizeUserName_(userName);
  const eventRows = loadEventRowsForDate_(eventSheet, eventDate);
  const matches = [];

  eventRows.forEach((event) => {
    const email = sanitizeEmailSafe_(event.form_email);
    const customer = email ? masterIndex.byEmail[email] || null : null;
    const displayName = getParticipantDisplayName_(event, customer);

    if (displayName !== targetName) {
      return;
    }

    matches.push({
      event,
      customer,
      email: email || (customer ? String(customer.email || "") : ""),
      displayName: customer
        ? String(customer.user_name || targetName)
        : String(event.receipt_name || targetName),
    });
  });

  return matches;
}

/**
 * 参加者の表示用氏名を取得（スペース除去済みで比較）
 * @param {Object} event
 * @param {Object|null} customer
 * @returns {string}
 */
function getParticipantDisplayName_(event, customer) {
  if (event.user_name) {
    return sanitizeUserName_(event.user_name);
  }
  if (customer && customer.user_name) {
    return sanitizeUserName_(customer.user_name);
  }
  if (event.receipt_name) {
    return sanitizeUserName_(event.receipt_name);
  }
  return "";
}

/**
 * 参加者を特定して master / event_histories に LINE ID を紐付け
 * @param {string} lineUserId
 * @param {string} eventDate
 * @param {Object} selected
 * @param {string} userName
 * @param {string} lineDisplayName
 * @returns {Object}
 */
function linkCheckinParticipant_(lineUserId, eventDate, selected, userName, lineDisplayName) {
  const masterSheet = getSheet_(SHEET_MASTER);
  const eventSheet = getSheet_(SHEET_EVENTS);
  const email = sanitizeEmail_(selected.email);

  let customer = selected.customer;
  if (!customer) {
    customer = findCustomerByEmail_(masterSheet, email);
  }

  const masterUpdates = {
    line_user_id: lineUserId,
    user_name: userName,
  };
  if (lineDisplayName) {
    masterUpdates.line_display_name = lineDisplayName;
  }

  if (!customer) {
    customer = appendMasterCustomer_(masterSheet, {
      line_user_id: lineUserId,
      line_display_name: lineDisplayName || "",
      user_name: userName,
      company_name: "",
      position_category: "",
      position_name: "",
      email: email,
      phone_number: "",
      referrer: "当日受付LIFF",
      status: CustomerStatus.ACTIVE,
    });
  } else {
    updateMasterFields_(masterSheet, customer._rowIndex, masterUpdates);
    customer = rowToObject_(masterSheet, customer._rowIndex);
  }

  let eventHistory = selected.event;
  if (!eventHistory.line_user_id) {
    updateEventFields_(eventSheet, eventHistory._rowIndex, {
      line_user_id: lineUserId,
    });
    eventHistory = rowToObject_(eventSheet, eventHistory._rowIndex);
  }

  const lineNotification = sendLineNotification_(
    buildCheckinLinkNotification_({
      pattern: MatchPattern.EXISTING_EMAIL,
      customer,
      eventDate,
      email,
      userName,
    }),
  );

  return {
    ok: true,
    linked: true,
    needsEmailConfirm: false,
    customer: serializeCustomer_(customer),
    eventHistory: serializeEvent_(eventHistory),
    alreadyCheckedIn: String(eventHistory.attendance_status) === "済",
    lineNotification,
  };
}

/**
 * メールアドレスで名寄せし LINE ID を紐付け（レガシー・フォールバック用）
 * @param {Object} payload
 * @returns {Object}
 */
function handleCheckinLinkEmail_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventDate = requireString_(payload.event_date, "event_date");
  const email = sanitizeEmail_(payload.email);
  const lineDisplayName = sanitizeLineDisplayName_(payload.line_display_name || payload.display_name || "");
  const displayName = sanitizeUserName_(payload.display_name || "");

  const masterSheet = getSheet_(SHEET_MASTER);
  const eventSheet = getSheet_(SHEET_EVENTS);

  let customer = findCustomerByEmail_(masterSheet, email);
  let pattern = MatchPattern.EXISTING_EMAIL;

  if (!customer) {
    const nameMatches = displayName
      ? findCustomersByName_(masterSheet, displayName)
      : [];

    if (nameMatches.length > 0) {
      pattern = MatchPattern.PENDING_NAME;
      customer = appendMasterCustomer_(masterSheet, {
        line_user_id: lineUserId,
        line_display_name: lineDisplayName,
        user_name: displayName || "未登録",
        company_name: "",
        position_category: "",
        position_name: "",
        email: email,
        phone_number: "",
        referrer: "当日受付LIFF",
        status: CustomerStatus.PENDING_MATCH,
      });
    } else {
      pattern = MatchPattern.NEW_CUSTOMER;
      customer = appendMasterCustomer_(masterSheet, {
        line_user_id: lineUserId,
        line_display_name: lineDisplayName,
        user_name: displayName || sanitizeUserName_(email.split("@")[0]),
        company_name: "",
        position_category: "",
        position_name: "",
        email: email,
        phone_number: "",
        referrer: "当日受付LIFF",
        status: CustomerStatus.ACTIVE,
      });
    }
  } else {
    const updates = { line_user_id: lineUserId };
    if (lineDisplayName) {
      updates.line_display_name = lineDisplayName;
    }
    updateMasterFields_(masterSheet, customer._rowIndex, updates);
    customer = rowToObject_(masterSheet, customer._rowIndex);
  }

  let eventHistory = findEventHistoryForCheckin_(eventSheet, {
    eventDate,
    email,
    lineUserId,
  });

  if (!eventHistory) {
    eventHistory = appendEventHistory_(eventSheet, {
      event_date: eventDate,
      user_name: customer.user_name || "",
      name_reading: "",
      form_email: email,
      receipt_required: "不要",
      receipt_name: "",
      line_user_id: lineUserId,
    });
  } else if (!eventHistory.line_user_id) {
    updateEventFields_(eventSheet, eventHistory._rowIndex, {
      line_user_id: lineUserId,
    });
    eventHistory = rowToObject_(eventSheet, eventHistory._rowIndex);
  }

  const lineNotification = sendLineNotification_(
    buildCheckinLinkNotification_({
      pattern,
      customer,
      eventDate,
      email,
    }),
  );

  return {
    ok: true,
    pattern,
    patternLabel: getPatternLabel_(pattern),
    customer: serializeCustomer_(customer),
    eventHistory: serializeEvent_(eventHistory),
    alreadyCheckedIn: String(eventHistory.attendance_status) === "済",
    lineNotification,
  };
}

/**
 * 受付完了：マスタ更新 + 来場ステータス「済」+ LINE 通知
 * @param {Object} payload
 * @returns {Object}
 */
function handleCheckinComplete_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventDate = requireString_(payload.event_date, "event_date");

  const masterSheet = getSheet_(SHEET_MASTER);
  const eventSheet = getSheet_(SHEET_EVENTS);

  const customer = findCustomerByLineId_(masterSheet, lineUserId);
  if (!customer) {
    throw new Error("顧客マスタに LINE ID が見つかりません。メール連携を行ってください。");
  }

  const masterUpdates = {};
  if (payload.company_name !== undefined) {
    masterUpdates.company_name = String(payload.company_name || "").trim();
  }
  if (payload.position_name !== undefined) {
    masterUpdates.position_name = String(payload.position_name || "").trim();
  }
  if (payload.position_category !== undefined) {
    masterUpdates.position_category = String(payload.position_category || "").trim();
  }
  const lineDisplayName = sanitizeLineDisplayName_(payload.line_display_name || "");
  if (lineDisplayName) {
    masterUpdates.line_display_name = lineDisplayName;
  }
  if (Object.keys(masterUpdates).length > 0) {
    updateMasterFields_(masterSheet, customer._rowIndex, masterUpdates);
    Object.assign(customer, masterUpdates);
  }

  let eventHistory = findEventHistoryForCheckin_(eventSheet, {
    eventDate,
    lineUserId,
    email: customer.email,
  });

  if (!eventHistory) {
    eventHistory = appendEventHistory_(eventSheet, {
      event_date: eventDate,
      user_name: customer.user_name || "",
      name_reading: "",
      form_email: customer.email,
      receipt_required: "不要",
      receipt_name: "",
      line_user_id: lineUserId,
    });
  }

  updateEventFields_(eventSheet, eventHistory._rowIndex, {
    line_user_id: lineUserId,
    attendance_status: "済",
  });
  eventHistory = rowToObject_(eventSheet, eventHistory._rowIndex);

  const lineNotification = sendLineNotification_(
    buildCheckinCompleteNotification_({
      customer,
      eventDate,
      eventHistory,
    }),
  );

  return {
    ok: true,
    customer: serializeCustomer_(customer),
    eventHistory: serializeEvent_(eventHistory),
    lineNotification,
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} lineUserId
 * @returns {Object|null}
 */
// =============================================================================
// 領収書 LIFF API
// =============================================================================

function todayYmdJst_() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd");
}

function formatEventDateJapanese_(yyyymmdd) {
  const s = String(yyyymmdd || "").trim();
  if (s.length !== 8) {
    return s;
  }
  const y = s.slice(0, 4);
  const m = Number(s.slice(4, 6));
  const d = Number(s.slice(6, 8));
  return y + "\u5e74" + m + "\u6708" + d + "\u65e5";
}

function formatReceiptAmount_(yen) {
  const n = Number(yen);
  if (isNaN(n) || n < 0) {
    throw new Error("\u53c2\u52a0\u8cbb\u304c\u6b63\u3057\u304f\u8a2d\u5b9a\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002");
  }
  return "\u00a5" + n.toLocaleString("ja-JP");
}

function formatReceiptAddressee_(receiptName) {
  const n = String(receiptName || "").trim();
  if (n) {
    return n + "\u3000\u69d8";
  }
  return "\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u3000\u69d8";
}

function getEventSettingsSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_EVENT_SETTINGS);
  if (!sheet) {
    throw new Error("event_settings \u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002");
  }
  return sheet;
}

function findEventSettings_(eventDate) {
  const sheet = getEventSettingsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  const target = String(eventDate).trim();
  for (let row = 2; row <= lastRow; row++) {
    const record = rowToObject_(sheet, row);
    const rowDate = String(record["\u30a4\u30d9\u30f3\u30c8\u65e5"] || record.event_date || "").trim();
    if (rowDate === target) {
      return {
        event_date: rowDate,
        participation_fee: Number(record["\u53c2\u52a0\u8cbb"] || record.participation_fee || 0),
        receipt_available_from: String(
          record["\u9818\u53ce\u66f8\u767a\u884c\u958b\u59cb\u65e5"] || record.receipt_available_from || "",
        ).trim(),
        event_name: String(record["\u30a4\u30d9\u30f3\u30c8\u540d"] || record.event_name || "").trim(),
      };
    }
  }
  return null;
}

function getReceiptAvailableFrom_(settings) {
  const from = String(settings.receipt_available_from || "").trim();
  if (from && from.length === 8) {
    return from;
  }
  const d = String(settings.event_date || "").trim();
  if (d.length !== 8) {
    return "";
  }
  const dt = new Date(Number(d.slice(0, 4)), Number(d.slice(4, 6)) - 1, Number(d.slice(6, 8)));
  dt.setDate(dt.getDate() + 1);
  return Utilities.formatDate(dt, "Asia/Tokyo", "yyyyMMdd");
}

function isReceiptRequiredRow_(row) {
  return normalizeReceiptRequired_(row.receipt_required) === "\u8981";
}

function isAttendanceCheckedIn_(row) {
  return String(row.attendance_status || "").trim() === "\u6e08";
}

function isReceiptIssuanceOpen_(settings) {
  const from = getReceiptAvailableFrom_(settings);
  if (!from) {
    return false;
  }
  return todayYmdJst_() >= from;
}

function findEventHistoryForReceipt_(eventSheet, eventDate, lineUserId) {
  const rows = loadEventRowsForDate_(eventSheet, eventDate);
  return (
    rows.find((row) => String(row.line_user_id || "").trim() === lineUserId) || null
  );
}

function buildReceiptEligibility_(eventRow, settings) {
  if (!eventRow) {
    return { eligible: false, reason: "\u53c2\u52a0\u5c65\u6b74\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002" };
  }
  if (!isReceiptRequiredRow_(eventRow)) {
    return { eligible: false, reason: "\u9818\u53ce\u66f8\u306e\u7533\u8acb\u304c\u3042\u308a\u307e\u305b\u3093\u3002" };
  }
  if (!isAttendanceCheckedIn_(eventRow)) {
    return { eligible: false, reason: "\u53d7\u4ed8\u304c\u672a\u5b8c\u4e86\u3067\u3059\u3002" };
  }
  if (!settings) {
    return { eligible: false, reason: "event_settings \u306b\u8a2d\u5b9a\u304c\u3042\u308a\u307e\u305b\u3093\u3002" };
  }
  if (!isReceiptIssuanceOpen_(settings)) {
    const from = getReceiptAvailableFrom_(settings);
    return {
      eligible: false,
      reason: "\u9818\u53ce\u66f8\u306f " + formatEventDateJapanese_(from) + " \u4ee5\u964d\u306b\u767a\u884c\u3067\u304d\u307e\u3059\u3002",
      available_from: from,
    };
  }
  return {
    eligible: true,
    event_date: settings.event_date,
    event_name: settings.event_name || "\u3046\u304a\u306e\u4f1a",
    amount_display: formatReceiptAmount_(settings.participation_fee),
    addressee_preview: formatReceiptAddressee_(eventRow.receipt_name),
    receipt_dl_count: Number(eventRow.receipt_dl_count || 0),
    available_from: getReceiptAvailableFrom_(settings),
  };
}

function handleReceiptList_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventSheet = getSheet_(SHEET_EVENTS);
  const lastRow = eventSheet.getLastRow();
  const eligible = [];

  if (lastRow >= 2) {
    for (let row = 2; row <= lastRow; row++) {
      const eventRow = rowToObject_(eventSheet, row);
      if (String(eventRow.line_user_id || "").trim() !== lineUserId) {
        continue;
      }
      const settings = findEventSettings_(eventRow.event_date);
      const info = buildReceiptEligibility_(eventRow, settings);
      if (info.eligible) {
        eligible.push(info);
      }
    }
  }

  eligible.sort((a, b) => String(b.event_date).localeCompare(String(a.event_date)));
  return { ok: true, eligible: eligible, today: todayYmdJst_() };
}

function handleReceiptLookup_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventDate = requireString_(payload.event_date, "event_date");
  const eventSheet = getSheet_(SHEET_EVENTS);
  const eventRow = findEventHistoryForReceipt_(eventSheet, eventDate, lineUserId);
  const settings = findEventSettings_(eventDate);
  const info = buildReceiptEligibility_(eventRow, settings);
  return { ok: true, eligible: info.eligible, receipt: info };
}

function exportSheetAsPdf_(spreadsheetId, sheetGid) {
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?format=pdf&size=A4&portrait=true&fitw=true" +
    "&sheetnames=false&printtitle=false&pagenumbers=false" +
    "&gridlines=false&fzr=false&gid=" +
    sheetGid;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("PDF\u51fa\u529b\u306b\u5931\u6557\u3057\u307e\u3057\u305f (HTTP " + code + ")");
  }
  return response.getBlob().setName("receipt.pdf");
}

function buildReceiptPdfBase64_(settings, eventRow) {
  const ss = getSpreadsheet_();
  const template = ss.getSheetByName(SHEET_RECEIPT_TEMPLATE);
  if (!template) {
    throw new Error("receipt_template \u30b7\u30fc\u30c8\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093\u3002");
  }

  const workName = "_receipt_work_" + Date.now();
  const work = template.copyTo(ss).setName(workName);
  try {
    work.getRange(RECEIPT_TEMPLATE_CELLS.eventDate).setValue(
      formatEventDateJapanese_(settings.event_date),
    );
    work.getRange(RECEIPT_TEMPLATE_CELLS.addressee).setValue(
      formatReceiptAddressee_(eventRow.receipt_name),
    );
    work.getRange(RECEIPT_TEMPLATE_CELLS.amount).setValue(
      formatReceiptAmount_(settings.participation_fee),
    );
    SpreadsheetApp.flush();
    const blob = exportSheetAsPdf_(ss.getId(), work.getSheetId());
    return Utilities.base64Encode(blob.getBytes());
  } finally {
    ss.deleteSheet(work);
  }
}

function incrementReceiptDlCount_(eventSheet, rowIndex) {
  const row = rowToObject_(eventSheet, rowIndex);
  const current = Number(row.receipt_dl_count || 0);
  updateEventFields_(eventSheet, rowIndex, { receipt_dl_count: current + 1 });
}

function createReceiptPdfPayload_(payload) {
  const lineUserId = requireString_(payload.line_user_id, "line_user_id");
  const eventDate = requireString_(payload.event_date, "event_date");
  const eventSheet = getSheet_(SHEET_EVENTS);
  const eventRow = findEventHistoryForReceipt_(eventSheet, eventDate, lineUserId);
  const settings = findEventSettings_(eventDate);
  const info = buildReceiptEligibility_(eventRow, settings);
  if (!info.eligible) {
    throw new Error(info.reason || "\u9818\u53ce\u66f8\u3092\u767a\u884c\u3067\u304d\u307e\u305b\u3093\u3002");
  }

  const pdfBase64 = buildReceiptPdfBase64_(settings, eventRow);
  incrementReceiptDlCount_(eventSheet, eventRow._rowIndex);

  return {
    filename: "\u9818\u53ce\u66f8_" + eventDate + ".pdf",
    pdf_base64: pdfBase64,
    receipt: info,
  };
}

function handleReceiptGenerate_(payload) {
  const created = createReceiptPdfPayload_(payload);
  return {
    ok: true,
    filename: created.filename,
    mime_type: "application/pdf",
    pdf_base64: created.pdf_base64,
    receipt: created.receipt,
  };
}

function findCustomerByLineId_(sheet, lineUserId) {
  const col = getColumnIndexForKey_(sheet, "line_user_id");
  if (col === undefined) {
    throw new Error("master_customers に LINEユーザーID 列がありません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === lineUserId) {
      return rowToObject_(sheet, i + 2);
    }
  }
  return null;
}

function findEventHistoryForCheckin_(sheet, criteria) {
  const eventRows = loadEventRowsForDate_(sheet, criteria.eventDate);
  return findEventHistoryForCheckinFromRows_(eventRows, criteria);
}

/**
 * 読み込み済みの当日行から参加履歴を検索
 * @param {Object[]} eventRows
 * @param {Object} criteria
 * @returns {Object|null}
 */
function findEventHistoryForCheckinFromRows_(eventRows, criteria) {
  const matches = eventRows.filter((record) => {
    const emailMatch =
      criteria.email &&
      sanitizeEmailSafe_(record.form_email) === sanitizeEmailSafe_(criteria.email);
    const lineMatch =
      criteria.lineUserId &&
      String(record.line_user_id).trim() === criteria.lineUserId;
    return emailMatch || lineMatch;
  });

  if (matches.length === 0) {
    return null;
  }

  const pending = matches.find((r) => String(r.attendance_status) !== "済");
  return pending || matches[matches.length - 1];
}

/**
 * master_customers を一括読み込みしてインデックス化
 * @param {GoogleAppsScript.Spreadsheet.Sheet} masterSheet
 * @returns {{ byLineId: Object, byEmail: Object }}
 */
function loadMasterIndex_(masterSheet) {
  const headerMap = getHeaderMap_(masterSheet);
  const lastRow = masterSheet.getLastRow();
  const lastCol = masterSheet.getLastColumn();
  const byLineId = {};
  const byEmail = {};

  if (lastRow < 2) {
    return { byLineId, byEmail };
  }

  const numRows = lastRow - 1;
  if (numRows < 1) {
    return { byLineId, byEmail };
  }
  const values = masterSheet.getRange(2, 1, numRows, lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    const obj = rowValuesToObject_(headerMap, values[i], i + 2);
    const lineId = String(obj.line_user_id || "").trim();
    const email = sanitizeEmailSafe_(obj.email);
    if (lineId) {
      byLineId[lineId] = obj;
    }
    if (email) {
      byEmail[email] = obj;
    }
  }

  return { byLineId, byEmail };
}

/** 当日参加者リストの CacheService TTL（秒） */
const EVENT_ROWS_CACHE_TTL_SEC = 300;

/**
 * @param {string} eventDate
 * @returns {string}
 */
function eventRowsCacheKey_(eventDate) {
  return "event_rows_v1_" + String(eventDate);
}

/**
 * @param {Object} row
 * @returns {Object}
 */
function serializeEventRowForCache_(row) {
  const copy = Object.assign({}, row);
  if (copy.timestamp instanceof Date) {
    copy.timestamp = copy.timestamp.toISOString();
  }
  return copy;
}

/**
 * @param {Object} row
 * @returns {Object}
 */
function deserializeEventRowFromCache_(row) {
  const copy = Object.assign({}, row);
  if (copy.timestamp && typeof copy.timestamp === "string") {
    copy.timestamp = new Date(copy.timestamp);
  }
  return copy;
}

/**
 * @param {string} eventDate
 */
function invalidateEventRowsCache_(eventDate) {
  const key = String(eventDate || "").trim();
  if (!key) {
    return;
  }
  try {
    CacheService.getScriptCache().remove(eventRowsCacheKey_(key));
  } catch (ignore) {}
}

/**
 * 指定日の event_histories（CacheService 付き）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventSheet
 * @param {string} eventDate
 * @returns {Object[]}
 */
function loadEventRowsForDate_(eventSheet, eventDate) {
  const targetDate = String(eventDate || "").trim();
  if (!targetDate) {
    return [];
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = eventRowsCacheKey_(targetDate);
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached).map(deserializeEventRowFromCache_);
    } catch (parseError) {
      console.warn("event rows cache parse failed:", parseError);
    }
  }

  const rows = loadEventRowsForDateFromSheet_(eventSheet, targetDate);
  try {
    const payload = JSON.stringify(rows.map(serializeEventRowForCache_));
    if (payload.length < 90000) {
      cache.put(cacheKey, payload, EVENT_ROWS_CACHE_TTL_SEC);
    }
  } catch (cacheError) {
    console.warn("event rows cache put failed:", cacheError);
  }
  return rows;
}

/**
 * 指定日の event_histories をシートから一括読み込み（キャッシュなし）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventSheet
 * @param {string} eventDate
 * @returns {Object[]}
 */
function loadEventRowsForDateFromSheet_(eventSheet, eventDate) {
  const headerMap = getHeaderMap_(eventSheet);
  const lastRow = eventSheet.getLastRow();
  const lastCol = eventSheet.getLastColumn();
  if (lastRow < 2) {
    return [];
  }

  const values = eventSheet.getRange(2, 1, lastRow, lastCol).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const obj = rowValuesToObject_(headerMap, values[i], i + 2);
    if (String(obj.event_date) === String(eventDate)) {
      rows.push(obj);
    }
  }
  return rows;
}

/**
 * @param {Object<string, number>} headerMap
 * @param {Array} values
 * @param {number} rowIndex
 * @returns {Object}
 */
function rowValuesToObject_(headerMap, values, rowIndex) {
  const obj = { _rowIndex: rowIndex };
  Object.keys(headerMap).forEach((header) => {
    const key = COLUMN_LABEL_TO_KEY_[header] || header;
    obj[key] = values[headerMap[header]];
  });
  return obj;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @param {Object} fields
 */
function updateMasterFields_(sheet, rowIndex, fields) {
  Object.keys(fields).forEach((key) => {
    const col = getColumnIndexForKey_(sheet, key);
    if (col !== undefined) {
      sheet.getRange(rowIndex, col + 1).setValue(fields[key]);
    }
  });
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @param {Object} fields
 */
function updateEventFields_(sheet, rowIndex, fields) {
  let eventDateToInvalidate = "";
  if (sheet.getName() === SHEET_EVENTS) {
    const row = rowToObject_(sheet, rowIndex);
    eventDateToInvalidate = String(
      fields.event_date || row.event_date || "",
    ).trim();
  }

  Object.keys(fields).forEach((key) => {
    const col = getColumnIndexForKey_(sheet, key);
    if (col !== undefined) {
      sheet.getRange(rowIndex, col + 1).setValue(fields[key]);
    }
  });

  if (eventDateToInvalidate) {
    invalidateEventRowsCache_(eventDateToInvalidate);
  }
}

/**
 * @param {Object} customer
 * @returns {Object}
 */
function serializeCustomer_(customer) {
  return {
    line_user_id: customer.line_user_id || "",
    line_display_name: customer.line_display_name || "",
    user_name: customer.user_name || "",
    company_name: customer.company_name || "",
    position_category: customer.position_category || "",
    position_name: customer.position_name || "",
    email: customer.email || "",
    phone_number: customer.phone_number || "",
    status: customer.status || "",
  };
}

/**
 * @param {Object} event
 * @returns {Object}
 */
function serializeEvent_(event) {
  return {
    event_date: String(event.event_date || ""),
    user_name: event.user_name || "",
    name_reading: event.name_reading || "",
    form_email: event.form_email || "",
    attendance_status: String(event.attendance_status || ""),
    payment_status: String(event.payment_status || ""),
    line_user_id: event.line_user_id || "",
  };
}

/**
 * @param {Object} ctx
 * @returns {string}
 */
function buildCheckinLinkNotification_(ctx) {
  return [
    "【うおの会】当日受付：参加者確認",
    "",
    `イベント日: ${ctx.eventDate}`,
    `氏名: ${ctx.userName || ctx.customer.user_name}`,
    `LINE表示名: ${ctx.customer.line_display_name || "—"}`,
    `メール: ${ctx.email}`,
    `会社: ${ctx.customer.company_name || "—"}`,
    `status: ${ctx.customer.status}`,
  ].join("\n");
}

/**
 * @param {Object} ctx
 * @returns {string}
 */
function buildCheckinCompleteNotification_(ctx) {
  return [
    "【うおの会】✅ 受付完了",
    "",
    `イベント日: ${ctx.eventDate}`,
    `氏名: ${ctx.customer.user_name}`,
    `LINE表示名: ${ctx.customer.line_display_name || "—"}`,
    `会社: ${ctx.customer.company_name || "—"}`,
    `役職: ${ctx.customer.position_name || "—"}`,
    `メール: ${ctx.customer.email}`,
    `来場ステータス: 済`,
  ].join("\n");
}

// =============================================================================
// コア業務ロジック：イベント申込 → 名寄せ
// =============================================================================

/**
 * Googleフォーム等からのイベント申込を処理
 * @param {Object} payload
 * @returns {Object}
 */
function handleEventApplication_(payload) {
  const eventDate = requireString_(payload.event_date, "event_date");
  const sanitized = sanitizeApplicationInput_(payload);

  const masterSheet = getSheet_(SHEET_MASTER);
  const eventSheet = getSheet_(SHEET_EVENTS);

  ensureAllEventHeaders_(eventSheet);
  ensureAllMasterHeaders_(masterSheet);

  const data = {
    user_name: sanitized.user_name,
    name_reading: sanitized.name_reading || "",
    email: sanitized.email,
    event_date: eventDate,
    receipt_required: sanitized.receipt_required,
    receipt_name: sanitized.receipt_name || "",
  };

  const existingByEmail = findCustomerByEmail_(masterSheet, sanitized.email);
  const existingByName = findCustomersByName_(masterSheet, sanitized.user_name);

  let pattern;
  let masterRow = existingByEmail;

  if (existingByEmail) {
    pattern = MatchPattern.EXISTING_EMAIL;
  } else if (existingByName.length > 0) {
    pattern = MatchPattern.PENDING_NAME;
    masterRow = writeMasterRow_(masterSheet, Object.assign({}, data, {
      referrer: sanitized.referrer || "フォーム申込",
    }));
    updateMasterFields_(masterSheet, masterRow._rowIndex, {
      status: CustomerStatus.PENDING_MATCH,
      company_name: sanitized.company_name,
      position_category: sanitized.position_category,
      position_name: sanitized.position_name,
      phone_number: sanitized.phone_number,
      referrer: sanitized.referrer,
    });
    masterRow = rowToObject_(masterSheet, masterRow._rowIndex);
  } else {
    pattern = MatchPattern.NEW_CUSTOMER;
    masterRow = writeMasterRow_(masterSheet, data);
    if (sanitized.company_name || sanitized.phone_number || sanitized.referrer) {
      updateMasterFields_(masterSheet, masterRow._rowIndex, {
        company_name: sanitized.company_name,
        position_category: sanitized.position_category,
        position_name: sanitized.position_name,
        phone_number: sanitized.phone_number,
        referrer: sanitized.referrer || "フォーム申込",
      });
      masterRow = rowToObject_(masterSheet, masterRow._rowIndex);
    }
  }

  const existingEvent = findEventByEmailAndDate_(
    eventSheet,
    sanitized.email,
    eventDate,
  );
  let eventRow;
  let duplicate = false;

  if (existingEvent) {
    duplicate = true;
    fillEmptyEventCells_(eventSheet, existingEvent._rowIndex, data);
    eventRow = rowToObject_(eventSheet, existingEvent._rowIndex);
  } else {
    eventRow = writeEventRow_(eventSheet, Object.assign({}, data, {
      line_user_id: existingByEmail ? existingByEmail.line_user_id : sanitized.line_user_id || "",
    }));
  }

  SpreadsheetApp.flush();

  const notificationPayload = buildNotificationMessage_({
    pattern,
    sanitized,
    eventDate,
    existingByEmail,
    existingByName,
    masterRow,
    eventRow,
  });

  const lineNotification = sendLineNotification_(notificationPayload.text);

  return {
    ok: true,
    pattern,
    patternLabel: getPatternLabel_(pattern),
    customer: existingByEmail || masterRow,
    eventHistory: eventRow,
    lineNotification,
    pendingMatchCandidates:
      pattern === MatchPattern.PENDING_NAME
        ? existingByName.map((c) => ({
            email: c.email,
            user_name: c.user_name,
            company_name: c.company_name,
            status: c.status,
          }))
        : [],
  };
}

// =============================================================================
// データサニタイズ
// =============================================================================

/**
 * 申込 payload 全体をサニタイズ
 * @param {Object} payload
 * @returns {Object}
 */
function sanitizeApplicationInput_(payload) {
  return {
    line_user_id: String(payload.line_user_id || "").trim(),
    user_name: sanitizeUserName_(payload.user_name),
    name_reading: sanitizeNameReading_(payload.name_reading),
    company_name: String(payload.company_name || "").trim(),
    position_category: String(payload.position_category || "").trim(),
    position_name: String(payload.position_name || "").trim(),
    email: sanitizeEmail_(payload.email),
    phone_number: sanitizePhoneNumber_(payload.phone_number),
    referrer: String(payload.referrer || "").trim(),
    receipt_required: String(payload.receipt_required || "不要").trim(),
    receipt_name: String(payload.receipt_name || "").trim(),
  };
}

/**
 * 氏名：全角・半角スペースをすべて削除
 * @param {string} name
 * @returns {string}
 */
function sanitizeUserName_(name) {
  return toHalfWidth_(String(name || ""))
    .replace(/[\s\u3000]+/g, "")
    .trim();
}

/**
 * 読み方：半角化のみ（スペース・かなは保持）
 * @param {string} reading
 * @returns {string}
 */
function sanitizeNameReading_(reading) {
  return toHalfWidth_(String(reading || "")).trim();
}

/**
 * LINE プロフィール表示名（スペース・絵文字はそのまま保持）
 * @param {string} name
 * @returns {string}
 */
function sanitizeLineDisplayName_(name) {
  return String(name || "").trim();
}

/**
 * メール：半角化 → 小文字化
 * @param {string} email
 * @returns {string}
 */
function sanitizeEmail_(email) {
  const normalized = toHalfWidth_(String(email || "")).trim().toLowerCase();
  if (!normalized) {
    throw new Error("email は必須です。");
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) {
    throw new Error("email の形式が正しくありません。");
  }
  return normalized;
}

/**
 * 電話番号：半角化し、数字以外を除去（記号・ハイフン対応）
 * @param {string} phone
 * @returns {string}
 */
function sanitizePhoneNumber_(phone) {
  const half = toHalfWidth_(String(phone || ""));
  return half.replace(/[^\d]/g, "");
}

/**
 * 全角英数・記号を半角に変換（スペース含む）
 * @param {string} str
 * @returns {string}
 */
function toHalfWidth_(str) {
  return String(str)
    .replace(/[\uFF01-\uFF5E]/g, (ch) =>
      String.fromCharCode(ch.charCodeAt(0) - 0xfee0),
    )
    .replace(/\u3000/g, " ");
}

/**
 * 領収書の要否を「要」/「不要」に正規化
 * @param {*} value
 * @returns {string}
 */
function normalizeReceiptRequired_(value) {
  const v = toHalfWidth_(String(value || "")).trim().toLowerCase();
  if (!v) {
    return "不要";
  }
  if (
    v === "要" ||
    v === "必要" ||
    v === "はい" ||
    v === "yes" ||
    v === "y" ||
    v.indexOf("必要") >= 0 ||
    v.indexOf("ほしい") >= 0
  ) {
    return "要";
  }
  return "不要";
}

/**
 * イベント日を YYYYMMDD 形式に正規化
 * 例: 「【受付中】7/7(火) 19:00~ うおの会」→ 20260707
 * @param {*} value
 * @returns {string}
 */
function normalizeEventDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    return `${y}${m}${d}`;
  }

  const str = toHalfWidth_(String(value || "")).trim();
  if (!str) {
    throw new Error(
      "行の参加イベントが空です。フォームで開催日を含む選択肢を選んでください。",
    );
  }

  const digitsOnly = str.replace(/[^\d]/g, "");
  if (digitsOnly.length === 8) {
    return digitsOnly;
  }

  const ymdMatch = str.match(/(\d{4})[年\/\-.](\d{1,2})[月\/\-.](\d{1,2})/);
  if (ymdMatch) {
    const y = ymdMatch[1];
    const m = String(ymdMatch[2]).padStart(2, "0");
    const d = String(ymdMatch[3]).padStart(2, "0");
    return `${y}${m}${d}`;
  }

  const mdMatch = str.match(/(\d{1,2})\/(\d{1,2})/);
  if (mdMatch) {
    const props = PropertiesService.getScriptProperties();
    const year =
      props.getProperty("DEFAULT_EVENT_YEAR") ||
      String(new Date().getFullYear());
    const m = String(mdMatch[1]).padStart(2, "0");
    const d = String(mdMatch[2]).padStart(2, "0");
    return `${year}${m}${d}`;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, "0");
    const d = String(parsed.getDate()).padStart(2, "0");
    return `${y}${m}${d}`;
  }

  throw new Error(
    "参加イベントからイベント日を読み取れません（例: 7/7 または 20260707）: " + str,
  );
}

// =============================================================================
// スプレッドシート操作
// =============================================================================

/**
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      return active;
    }
  } catch (error) {
  }
  throw new Error(
    "Script Property SPREADSHEET_ID が未設定です。スプレッドシートに紐づく GAS から実行するか、ID を設定してください。",
  );
}

/**
 * 1行分のセル値を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex 1始まり
 * @returns {*[]}
 */
function getRowValues_(sheet, rowIndex) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return [];
  }
  return sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
}

/**
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getSheet_(sheetName) {
  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    const headersBySheet = {
      [SHEET_MASTER]: MASTER_HEADERS,
      [SHEET_EVENTS]: EVENT_HEADERS,
    };
    const headers = headersBySheet[sheetName];
    if (!headers) {
      throw new Error(`シート "${sheetName}" が見つかりません。`);
    }
    sheet = spreadsheet.insertSheet(sheetName);
    ensureHeaders_(sheet, headers);
  } else if (sheetName === SHEET_MASTER) {
    ensureAllMasterHeaders_(sheet);
  } else if (sheetName === SHEET_EVENTS) {
    ensureAllEventHeaders_(sheet);
  }

  return sheet;
}

/**
 * ヘッダー行から列インデックス（0始まり）マップを生成
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object<string, number>}
 */
function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error(`シート "${sheet.getName()}" にヘッダーがありません。`);
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) {
      map[String(header).trim()] = index;
    }
  });
  return map;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex 1始まり
 * @returns {Object}
 */
function rowToObject_(sheet, rowIndex) {
  const headerMap = getHeaderMap_(sheet);
  const values = getRowValues_(sheet, rowIndex);
  const obj = { _rowIndex: rowIndex };
  Object.keys(headerMap).forEach((header) => {
    const key = COLUMN_LABEL_TO_KEY_[header] || header;
    obj[key] = values[headerMap[header]];
  });
  return obj;
}

/**
 * メール完全一致で顧客を1件取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} email
 * @returns {Object|null}
 */
function findCustomerByEmail_(sheet, email) {
  const emailCol = getColumnIndexForKey_(sheet, "email");
  if (emailCol === undefined) {
    throw new Error("master_customers にメールアドレス列がありません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  for (let row = 2; row <= lastRow; row++) {
    const values = getRowValues_(sheet, row);
    const cellEmail = sanitizeEmailSafe_(values[emailCol]);
    if (cellEmail && cellEmail === email) {
      return rowToObject_(sheet, row);
    }
  }
  return null;
}

/**
 * サニタイズ済み氏名の完全一致で顧客一覧を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} userName
 * @returns {Object[]}
 */
function findCustomersByName_(sheet, userName) {
  const nameCol = getColumnIndexForKey_(sheet, "user_name");
  if (nameCol === undefined) {
    throw new Error("master_customers に氏名列がありません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const matches = [];
  for (let row = 2; row <= lastRow; row++) {
    const values = getRowValues_(sheet, row);
    const cellName = sanitizeUserName_(values[nameCol]);
    if (cellName && cellName === userName) {
      matches.push(rowToObject_(sheet, row));
    }
  }
  return matches;
}

/**
 * 既存行の email 比較用（空や不正値は null）
 * @param {*} value
 * @returns {string|null}
 */
function sanitizeEmailSafe_(value) {
  try {
    if (!value) {
      return null;
    }
    return sanitizeEmail_(value);
  } catch (error) {
    return null;
  }
}

/**
 * スキーマのキーに対応する列へ1行分を書き込む（列順ズレ防止）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object[]} schema
 * @param {Object} values
 * @returns {number} 書き込んだ行番号
 */
function appendRowBySchema_(sheet, schema, values) {
  const rowIndex = Math.max(sheet.getLastRow(), 1) + 1;

  schema.forEach((col) => {
    if (values[col.key] === undefined) {
      return;
    }
    let colIdx = getColumnIndexForKey_(sheet, col.key, schema);
    if (colIdx === undefined) {
      colIdx = sheet.getLastColumn();
      sheet.getRange(1, colIdx + 1).setValue(col.label);
    }
    sheet.getRange(rowIndex, colIdx + 1).setValue(values[col.key]);
  });

  return rowIndex;
}

/**
 * 列が無ければ末尾にヘッダーを追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} col
 * @returns {number} 列インデックス（0始まり）
 */
function ensureColumnForKey_(sheet, col) {
  let colIdx = getColumnIndexForKey_(sheet, col.key);
  if (colIdx === undefined) {
    colIdx = sheet.getLastColumn();
    sheet.getRange(1, colIdx + 1).setValue(col.label);
  }
  return colIdx;
}

/**
 * 顧客マスタに1行追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 * @returns {Object}
 */
function appendMasterCustomer_(sheet, data) {
  ensureAllMasterHeaders_(sheet);

  const values = {
    line_user_id: data.line_user_id || "",
    line_display_name: data.line_display_name || "",
    user_name: data.user_name || "",
    company_name: data.company_name || "",
    position_category: data.position_category || "",
    position_name: data.position_name || "",
    email: data.email || "",
    phone_number: data.phone_number || "",
    referrer: data.referrer || "",
    status: data.status || CustomerStatus.ACTIVE,
    created_at: new Date(),
  };

  const newRowIndex = appendRowBySchema_(sheet, MASTER_SCHEMA, values);
  return rowToObject_(sheet, newRowIndex);
}

/**
 * イベント参加履歴に1行追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 * @returns {Object}
 */
function appendEventHistory_(sheet, data) {
  ensureAllEventHeaders_(sheet);

  const values = {
    timestamp: new Date(),
    event_date: data.event_date,
    user_name: data.user_name || "",
    name_reading: data.name_reading || "",
    form_email: data.form_email,
    receipt_required: data.receipt_required || "不要",
    receipt_name: data.receipt_name || "",
    line_user_id: data.line_user_id || "",
    payment_status: "未",
    attendance_status: "未",
    receipt_dl_count: 0,
  };

  const newRowIndex = appendRowBySchema_(sheet, EVENT_SCHEMA, values);
  if (data.event_date) {
    invalidateEventRowsCache_(String(data.event_date));
  }
  return rowToObject_(sheet, newRowIndex);
}

/**
 * ヘッダー行が無ければ作成
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} expectedHeaders
 */
function ensureHeaders_(sheet, expectedHeaders) {
  if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
    return;
  }
  sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
}

/**
 * master_customers に不足しているヘッダー列を末尾に追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureAllMasterHeaders_(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
    return;
  }

  const headerMap = getHeaderMap_(sheet);
  MASTER_SCHEMA.forEach((col) => {
    const hasLabel = headerMap[col.label] !== undefined;
    const hasLegacy = headerMap[col.key] !== undefined;
    if (!hasLabel && !hasLegacy) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col.label);
    }
  });
}

/**
 * event_histories に不足しているヘッダー列を末尾に追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureAllEventHeaders_(sheet) {
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, EVENT_HEADERS.length).setValues([EVENT_HEADERS]);
    return;
  }

  const headerMap = getHeaderMap_(sheet);
  EVENT_SCHEMA.forEach((col) => {
    const hasLabel = headerMap[col.label] !== undefined;
    const hasLegacy = headerMap[col.key] !== undefined;
    if (!hasLabel && !hasLegacy) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col.label);
    }
  });
}

// =============================================================================
// LINE 通知
// =============================================================================

/**
 * 魚谷さん（管理者）の LINE へ Push 通知
 * 失敗しても業務処理は継続する（受付を止めない）
 * @param {string} message
 * @returns {{ sent: boolean, error?: string }}
 */
function sendLineNotification_(message) {
  const token = PropertiesService.getScriptProperties().getProperty(
    "LINE_CHANNEL_ACCESS_TOKEN",
  );
  const userId = PropertiesService.getScriptProperties().getProperty(
    "LINE_ADMIN_USER_ID",
  );

  if (!token || !userId) {
    const msg = "LINE_CHANNEL_ACCESS_TOKEN または LINE_ADMIN_USER_ID が未設定";
    console.warn("LINE 通知スキップ:", msg);
    return { sent: false, error: msg };
  }

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: userId,
    messages: [{ type: "text", text: message }],
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const body = response.getContentText();

  if (status < 200 || status >= 300) {
    console.error("LINE Push 失敗:", status, body);
    return {
      sent: false,
      error: `LINE Push 失敗 (${status}): ${body}`,
    };
  }

  return { sent: true };
}

/**
 * @param {Object} ctx
 * @returns {{ text: string }}
 */
function buildNotificationMessage_(ctx) {
  const label = getPatternLabel_(ctx.pattern);
  const lines = [
    "【うおの会】イベント申込を受信しました",
    "",
    `■ 処理パターン: ${label}`,
    `■ イベント日: ${ctx.eventDate}`,
    "",
    "■ 申込内容",
    `氏名: ${ctx.sanitized.user_name}`,
    `読み方: ${ctx.sanitized.name_reading || "—"}`,
    `メール: ${ctx.sanitized.email}`,
    `会社: ${ctx.sanitized.company_name || "—"}`,
    `区分: ${ctx.sanitized.position_category || "—"}`,
    `役職: ${ctx.sanitized.position_name || "—"}`,
    `電話: ${ctx.sanitized.phone_number || "—"}`,
    `紹介者: ${ctx.sanitized.referrer || "—"}`,
    `領収書: ${ctx.sanitized.receipt_required}`,
    `領収書宛名: ${ctx.sanitized.receipt_name || "—"}`,
  ];

  if (ctx.pattern === MatchPattern.EXISTING_EMAIL && ctx.existingByEmail) {
    lines.push("", "■ 既存顧客（メール一致）");
    lines.push(`status: ${ctx.existingByEmail.status}`);
    lines.push(`line_user_id: ${ctx.existingByEmail.line_user_id || "未連携"}`);
  }

  if (ctx.pattern === MatchPattern.PENDING_NAME && ctx.existingByName.length > 0) {
    lines.push("", "■ 同一氏名の既存顧客（要確認）");
    ctx.existingByName.slice(0, 5).forEach((c, i) => {
      lines.push(
        `${i + 1}. ${c.user_name} / ${c.email} / ${c.company_name || "—"} / ${c.status}`,
      );
    });
    lines.push("", "→ 新規行は status=pending_match で追加しました。");
  }

  if (ctx.pattern === MatchPattern.NEW_CUSTOMER) {
    lines.push("", "■ 完全新規顧客として master に active で追加しました。");
  }

  lines.push("", `event_histories 行: ${ctx.eventRow._rowIndex}`);

  return { text: lines.join("\n") };
}

/**
 * @param {string} pattern
 * @returns {string}
 */
function getPatternLabel_(pattern) {
  switch (pattern) {
    case MatchPattern.EXISTING_EMAIL:
      return "パターンA：既存顧客（メール一致）";
    case MatchPattern.PENDING_NAME:
      return "パターンB：承認待ち（氏名一致・メール不一致）";
    case MatchPattern.NEW_CUSTOMER:
      return "パターンC：完全新規";
    default:
      return pattern;
  }
}

// =============================================================================
// ユーティリティ
// =============================================================================

/**
 * @param {GoogleAppsScript.Events.DoPost} e
 * @returns {Object}
 */
function parseRequestBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("POST ボディが空です。");
  }

  const contentType = e.postData.type || "";
  const raw = e.postData.contents;

  if (contentType.indexOf("application/json") >= 0) {
    return JSON.parse(raw);
  }

  // Googleフォーム等 form-urlencoded 対応
  if (contentType.indexOf("application/x-www-form-urlencoded") >= 0) {
    const params = {};
    raw.split("&").forEach((pair) => {
      const [key, value] = pair.split("=").map(decodeURIComponent);
      params[key] = value;
    });
    return params;
  }

  // フォールバック: JSON としてパース試行
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error("リクエストボディを解析できません。");
  }
}

/**
 * @param {number} status
 * @param {Object} body
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function jsonResponse_(status, body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(
    ContentService.MimeType.JSON,
  );
}

/**
 * @param {*} value
 * @param {string} fieldName
 * @returns {string}
 */
function requireString_(value, fieldName) {
  const str = String(value || "").trim();
  if (!str) {
    throw new Error(`${fieldName} は必須です。`);
  }
  return str;
}

// =============================================================================
// Googleフォーム連携
// =============================================================================

/** フォーム設問タイトル → payload キーのエイリアス */
const FORM_FIELD_ALIASES = {
  event_date: [
    "参加希望日",
    "イベント日",
    "開催日",
    "event_date",
    "参加日",
    "うおの会",
    "参加する回",
    "参加イベント",
  ],
  user_name: [
    "氏名",
    "お名前",
    "名前",
    "user_name",
    "おなまえ",
    "ネーム",
  ],
  name_reading: [
    "読み方",
    "ふりがな",
    "フリガナ",
    "よみがな",
    "name_reading",
    "お名前（ふりがな）",
  ],
  email: [
    "メールアドレス",
    "メール",
    "email",
    "Email",
    "Eメール",
    "連絡先",
  ],
  company_name: ["会社名", "company_name", "御社名", "所属"],
  position_category: [
    "区分",
    "立場",
    "position_category",
    "職種",
    "属性",
  ],
  position_name: ["役職", "役職名", "position_name", "肩書き"],
  phone_number: ["電話番号", "電話", "phone_number", "携帯番号"],
  referrer: ["紹介者", "referrer", "ご紹介者", "紹介者名"],
  receipt_required: [
    "領収書",
    "領収書の要不要",
    "領収書の有無",
    "receipt_required",
    "領収書希望",
    "領収書は必要",
  ],
  receipt_name: [
    "領収書宛名",
    "receipt_name",
    "領収書の宛名",
  ],
};

/**
 * =============================================================================
 * フォーム → スプレッドシート転記（シンプル版）
 * =============================================================================
 *
 * 【やること】フォーム回答1行を読み取り、2シートに書く
 *   - event_histories  … 毎回追加（参加者リスト）
 *   - master_customers … 新規メールの人だけ追加
 *
 * 【初回】GASエディタで runFullFormSync() を1回実行
 * 【以降】フォーム送信で onFormSubmit が自動実行
 * 【手動】processAllFormResponses() または スプレッドシートメニュー「うおの会」
 */

// =============================================================================
// エントリポイント
// =============================================================================

/** フォーム送信トリガー（関数: onFormSubmit / イベント: フォーム送信時） */
function onFormSubmit(e) {
  transferFormResponse_(e);
}

/** 初回セットアップ＋全行転記 */
function runFullFormSync() {
  setupSpreadsheetHeaders();
  reinstallFormSubmitTrigger();
  return processAllFormResponses();
}

/** 全フォーム回答を転記 */
function processAllFormResponses() {
  const sheet = getFormResponseSheet_();
  const lastRow = sheet.getLastRow();
  const results = [];

  if (lastRow < 2) {
    const msg = "フォーム回答がありません。";
    Logger.log(msg);
    alertIfUi_(msg);
    return results;
  }

  for (let row = 2; row <= lastRow; row++) {
    try {
      const result = transferFormRow_(sheet, row);
      results.push({ row: row, ok: true, name: result.data.user_name, duplicate: result.duplicate });
    } catch (error) {
      results.push({ row: row, ok: false, error: error.message || String(error) });
    }
  }

  const summary = summarizeFormSync_(results);
  Logger.log(summary);
  alertIfUi_("フォーム転記結果\n\n" + summary);
  return results;
}

/** 最終行だけ転記（テスト用） */
function processLatestFormResponse() {
  const sheet = getFormResponseSheet_();
  const row = sheet.getLastRow();
  if (row < 2) {
    throw new Error("フォーム回答がありません。");
  }
  const result = transferFormRow_(sheet, row);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/** トリガー設定 */
function installFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.some((t) => t.getHandlerFunction() === "onFormSubmit")) {
    Logger.log("トリガー設定済み");
    return;
  }
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(getSpreadsheet_())
    .onFormSubmit()
    .create();
  Logger.log("トリガーを設定しました");
}

function uninstallFormSubmitTrigger() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function reinstallFormSubmitTrigger() {
  uninstallFormSubmitTrigger();
  installFormSubmitTrigger();
}

/** スプレッドシートメニュー */

// =============================================================================
// 転記の本体
// =============================================================================

function transferFormResponse_(e) {
  if (!e || !e.range) {
    throw new Error("フォームイベントが不正です");
  }
  try {
    const result = transferFormRow_(e.range.getSheet(), e.range.getRow());
    Logger.log("転記OK: " + JSON.stringify(result));
    return result;
  } catch (error) {
    const msg = error.message || String(error);
    Logger.log("転記エラー: " + msg);
    try {
      if (typeof sendLineNotification_ === "function") {
        sendLineNotification_("【うおの会】フォーム転記エラー\n\n" + msg);
      }
    } catch (ignore) {}
    throw error;
  }
}

function transferFormRow_(formSheet, rowIndex) {
  const data = readFormRow_(formSheet, rowIndex);

  const eventSheet = getSheet_(SHEET_EVENTS);
  const masterSheet = getSheet_(SHEET_MASTER);

  ensureAllEventHeaders_(eventSheet);
  ensureAllMasterHeaders_(masterSheet);

  const existing = findEventByEmailAndDate_(eventSheet, data.email, data.event_date);
  let eventRow;
  let duplicate = false;

  if (existing) {
    duplicate = true;
    fillEmptyEventCells_(eventSheet, existing._rowIndex, data);
    eventRow = rowToObject_(eventSheet, existing._rowIndex);
  } else {
    eventRow = writeEventRow_(eventSheet, data);
  }

  let masterRow = findCustomerByEmail_(masterSheet, data.email);
  if (!masterRow) {
    masterRow = writeMasterRow_(masterSheet, data);
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    duplicate: duplicate,
    data: data,
    eventRow: eventRow,
    masterRow: masterRow,
  };
}

// =============================================================================
// フォーム1行の読み取り
// =============================================================================

function readFormRow_(sheet, rowIndex) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    throw new Error("フォーム回答シートに列がありません");
  }

  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map((h) => String(h || "").trim());
  const cells = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  let userName = "";
  let nameReading = "";
  let eventRaw = "";
  let receiptRequired = "不要";
  let receiptName = "";
  const emails = [];

  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    const raw = cells[i];
    const value =
      raw instanceof Date
        ? raw
        : String(raw === undefined || raw === null ? "" : raw).trim();
    if (!header || value === "" || value === null || value === undefined) {
      continue;
    }

    if (header === "氏名" || (header.indexOf("氏名") >= 0 && header.indexOf("読み") < 0)) {
      userName = String(value).trim();
    } else if (header === "読み方" || header.indexOf("読み") >= 0) {
      nameReading = String(value).trim();
    } else if (header.indexOf("参加イベント") >= 0 || header.indexOf("参加") >= 0) {
      eventRaw = value;
    } else if (header.indexOf("領収書の有無") >= 0 || header === "領収書") {
      receiptRequired = String(value).trim();
    } else if (header.indexOf("宛名") >= 0) {
      receiptName = String(value).trim();
    }

    const valueForEmail = String(value).trim();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(valueForEmail)) {
      emails.push(valueForEmail);
    }
  }

  const email = emails.length >= 2 ? emails[emails.length - 1] : emails[0] || "";

  if (!userName) {
    throw new Error("行" + rowIndex + ": 氏名が空です");
  }
  if (!email) {
    throw new Error("行" + rowIndex + ": メールアドレスが空です");
  }
  if (!eventRaw) {
    throw new Error("行" + rowIndex + ": 参加イベントが空です");
  }

  return {
    user_name: sanitizeUserName_(userName),
    name_reading: nameReading,
    email: sanitizeEmail_(email),
    event_date: normalizeEventDate_(eventRaw),
    receipt_required: normalizeReceiptRequired_(receiptRequired),
    receipt_name: receiptName,
  };
}

// =============================================================================
// シートへの書き込み（日本語ヘッダー名で直接書く）
// =============================================================================

function writeEventRow_(sheet, data) {
  const rowNum = Math.max(sheet.getLastRow(), 1) + 1;
  const labels = {
    記録日時: new Date(),
    イベント日: data.event_date,
    氏名: data.user_name,
    読み方: data.name_reading || "",
    申込メール: data.email,
    領収書: data.receipt_required || "不要",
    領収書宛名: data.receipt_name || "",
    LINEユーザーID: data.line_user_id || "",
    支払い状況: "未",
    来場状況: "未",
    領収書DL数: 0,
  };
  writeCellsByLabels_(sheet, rowNum, labels);
  if (data.event_date) {
    invalidateEventRowsCache_(String(data.event_date));
  }
  return rowToObject_(sheet, rowNum);
}

function writeMasterRow_(sheet, data) {
  const rowNum = Math.max(sheet.getLastRow(), 1) + 1;
  const labels = {
    LINEユーザーID: "",
    LINE表示名: "",
    氏名: data.user_name,
    会社名: "",
    区分: "",
    役職: "",
    メールアドレス: data.email,
    電話番号: "",
    紹介者: "フォーム申込",
    ステータス: "active",
    登録日時: new Date(),
  };
  writeCellsByLabels_(sheet, rowNum, labels);
  return rowToObject_(sheet, rowNum);
}

function writeCellsByLabels_(sheet, rowNum, labels) {
  const headerMap = getHeaderMap_(sheet);
  Object.keys(labels).forEach((label) => {
    if (headerMap[label] !== undefined) {
      sheet.getRange(rowNum, headerMap[label] + 1).setValue(labels[label]);
    }
  });
}

function fillEmptyEventCells_(sheet, rowNum, data) {
  const row = rowToObject_(sheet, rowNum);
  const labels = {};
  if (!row.user_name && data.user_name) {
    labels["氏名"] = data.user_name;
  }
  if (!row.name_reading && data.name_reading) {
    labels["読み方"] = data.name_reading;
  }
  if (!row.receipt_name && data.receipt_name) {
    labels["領収書宛名"] = data.receipt_name;
  }
  writeCellsByLabels_(sheet, rowNum, labels);
  if (Object.keys(labels).length > 0 && row.event_date) {
    invalidateEventRowsCache_(String(row.event_date));
  }
}

function findEventByEmailAndDate_(sheet, email, eventDate) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const targetEmail = sanitizeEmail_(email);
  const targetDate = String(eventDate);

  for (let row = 2; row <= lastRow; row++) {
    const record = rowToObject_(sheet, row);
    const recordEmail = sanitizeEmailSafe_(record.form_email);
    const recordDate = String(record.event_date || "").trim();
    if (recordEmail === targetEmail && recordDate === targetDate) {
      return record;
    }
  }
  return null;
}

// =============================================================================
// ユーティリティ
// =============================================================================

function summarizeFormSync_(results) {
  let ok = 0;
  let dup = 0;
  let err = 0;
  const errors = [];

  results.forEach((r) => {
    if (!r.ok) {
      err++;
      errors.push("行" + r.row + ": " + r.error);
    } else if (r.duplicate) {
      dup++;
    } else {
      ok++;
    }
  });

  const lines = [
    "新規転記: " + ok + " 件",
    "重複スキップ: " + dup + " 件",
    "エラー: " + err + " 件",
  ];
  if (errors.length > 0) {
    lines.push("", errors.join("\n"));
  }
  return lines.join("\n");
}

function alertIfUi_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("うおの会")
    .addItem("フォーム回答を一括転記", "processAllFormResponses")
    .addItem("初回セットアップ＋転記", "runFullFormSync")
    .addSeparator()
    .addItem("設定を診断", "diagnoseFormSubmitSetup")
    .addItem("トリガーを再設定", "reinstallFormSubmitTrigger")
    .addToUi();
}

function getFormResponseSheet_() {
  const props = PropertiesService.getScriptProperties();
  const sheetName = props.getProperty("FORM_RESPONSE_SHEET_NAME");
  const spreadsheet = getSpreadsheet_();

  if (sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`フォーム回答シート "${sheetName}" が見つかりません。`);
    }
    return sheet;
  }

  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (
      name.indexOf("フォームの回答") >= 0 ||
      name.indexOf("Form Responses") >= 0 ||
      name.indexOf("Form_Responses") >= 0
    ) {
      return sheets[i];
    }
  }

  throw new Error(
    "フォーム回答シートが見つかりません。Script Property FORM_RESPONSE_SHEET_NAME を設定してください。",
  );
}

function diagnoseFormSubmitSetup() {
  const lines = ["=== フォーム連携診断 ==="];
  try {
    const spreadsheet = getSpreadsheet_();
    lines.push("スプレッドシート: " + spreadsheet.getName());
    const formSheet = getFormResponseSheet_();
    lines.push("フォーム回答シート: " + formSheet.getName());
    lines.push("回答行数: " + Math.max(formSheet.getLastRow() - 1, 0));
    lines.push(
      "トリガー数: " +
        ScriptApp.getProjectTriggers().filter(function (t) {
          return t.getHandlerFunction() === "onFormSubmit";
        }).length,
    );
    if (formSheet.getLastRow() >= 2) {
      const row = formSheet.getLastRow();
      const data = readFormRow_(formSheet, row);
      lines.push("--- 最終行テスト ---");
      lines.push("氏名: " + data.user_name);
      lines.push("読み方: " + (data.name_reading || "-"));
      lines.push("メール: " + data.email);
      lines.push("イベント日: " + data.event_date);
    }
  } catch (e) {
    lines.push("エラー: " + (e.message || String(e)));
  }
  const report = lines.join("\n");
  Logger.log(report);
  alertIfUi_(report);
  return report;
}

function processFormResponseRow(rowNumber) {
  const sheet = getFormResponseSheet_();
  return transferFormRow_(sheet, rowNumber);
}

/** シートのヘッダー行を初期化 */
function setupSpreadsheetHeaders() {
  ensureAllMasterHeaders_(getSheet_(SHEET_MASTER));
  ensureAllEventHeaders_(getSheet_(SHEET_EVENTS));
  Logger.log("ヘッダー行を初期化しました。");
}

/**
 * 名寄せロジックの動作確認（テストデータ）
 * 実行前に SPREADSHEET_ID を設定し、テスト用シートを使用してください。
 */
function testEventApplicationNewCustomer() {
  const result = handleEventApplication_({
    action: "event_application",
    event_date: "20260707",
    user_name: "テスト 太郎",
    email: "test.taro@example.com",
    company_name: "テスト株式会社",
    position_category: "経営者",
    position_name: "代表",
    phone_number: "090-1234-5678",
    referrer: "紹介者A",
    receipt_required: "要",
    receipt_name: "テスト太郎",
  });
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * LINE Push 通知の疎通確認（GAS エディタから手動実行）
 * 失敗時はログに HTTP ステータスとレスポンス本文が出ます。
 */
function testLineNotification() {
  const result = sendLineNotification_(
    "【うおの会】LINE 通知テスト\n\nこのメッセージが届けば設定は OK です。",
  );
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * フォーム転記用の日付・領収書正規化テスト（GAS エディタから手動実行）
 */
function testFormFieldNormalization() {
  const eventSamples = [
    "【受付中】7/7(火) 19:00~ うおの会",
    "20260707",
    "2026/7/7",
    new Date(2026, 6, 7),
  ];
  eventSamples.forEach((sample) => {
    Logger.log(
      JSON.stringify(sample) + " => " + normalizeEventDate_(sample),
    );
  });
  ["必要", "要", "不要", "はい", ""].forEach((sample) => {
    Logger.log(
      JSON.stringify(sample) + " => " + normalizeReceiptRequired_(sample),
    );
  });
}
