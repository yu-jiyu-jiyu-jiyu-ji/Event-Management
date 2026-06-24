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
 */

// =============================================================================
// 定数
// =============================================================================

const SHEET_MASTER = "master_customers";
const SHEET_EVENTS = "event_histories";

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

const MASTER_HEADERS = [
  "line_user_id",
  "line_display_name",
  "user_name",
  "company_name",
  "position_category",
  "position_name",
  "email",
  "phone_number",
  "referrer",
  "status",
  "created_at",
];

const EVENT_HEADERS = [
  "timestamp",
  "event_date",
  "form_email",
  "receipt_required",
  "receipt_name",
  "line_user_id",
  "payment_status",
  "attendance_status",
  "receipt_dl_count",
];

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
function findCustomerByLineId_(sheet, lineUserId) {
  const headerMap = getHeaderMap_(sheet);
  const col = headerMap.line_user_id;
  if (col === undefined) {
    throw new Error("master_customers に line_user_id 列がありません。");
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

  const values = masterSheet.getRange(2, 1, lastRow, lastCol).getValues();
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

/**
 * 指定日の event_histories を一括読み込み
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventSheet
 * @param {string} eventDate
 * @returns {Object[]}
 */
function loadEventRowsForDate_(eventSheet, eventDate) {
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
  Object.keys(headerMap).forEach((key) => {
    obj[key] = values[headerMap[key]];
  });
  return obj;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @param {Object} fields
 */
function updateMasterFields_(sheet, rowIndex, fields) {
  const headerMap = getHeaderMap_(sheet);
  Object.keys(fields).forEach((key) => {
    if (headerMap[key] !== undefined) {
      sheet.getRange(rowIndex, headerMap[key] + 1).setValue(fields[key]);
    }
  });
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @param {Object} fields
 */
function updateEventFields_(sheet, rowIndex, fields) {
  const headerMap = getHeaderMap_(sheet);
  Object.keys(fields).forEach((key) => {
    if (headerMap[key] !== undefined) {
      sheet.getRange(rowIndex, headerMap[key] + 1).setValue(fields[key]);
    }
  });
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

  const existingByEmail = findCustomerByEmail_(masterSheet, sanitized.email);
  const existingByName = findCustomersByName_(masterSheet, sanitized.user_name);

  let pattern;
  let masterRow = null;

  if (existingByEmail) {
    // 【パターンA】メール完全一致 → 既存顧客。マスタは更新しない
    pattern = MatchPattern.EXISTING_EMAIL;
  } else if (existingByName.length > 0) {
    // 【パターンB】氏名一致・メール不一致 → 承認待ちでマスタ追加
    pattern = MatchPattern.PENDING_NAME;
    masterRow = appendMasterCustomer_(masterSheet, {
      line_user_id: sanitized.line_user_id || "",
      line_display_name: "",
      user_name: sanitized.user_name,
      company_name: sanitized.company_name,
      position_category: sanitized.position_category,
      position_name: sanitized.position_name,
      email: sanitized.email,
      phone_number: sanitized.phone_number,
      referrer: sanitized.referrer,
      status: CustomerStatus.PENDING_MATCH,
    });
  } else {
    // 【パターンC】完全新規
    pattern = MatchPattern.NEW_CUSTOMER;
    masterRow = appendMasterCustomer_(masterSheet, {
      line_user_id: sanitized.line_user_id || "",
      line_display_name: "",
      user_name: sanitized.user_name,
      company_name: sanitized.company_name,
      position_category: sanitized.position_category,
      position_name: sanitized.position_name,
      email: sanitized.email,
      phone_number: sanitized.phone_number,
      referrer: sanitized.referrer,
      status: CustomerStatus.ACTIVE,
    });
  }

  const existingEvent = findEventHistoryForCheckin_(eventSheet, {
    eventDate: eventDate,
    email: sanitized.email,
  });
  if (existingEvent) {
    return {
      ok: true,
      duplicate: true,
      pattern,
      patternLabel: getPatternLabel_(pattern),
      message: "同一メール・同一イベント日の申込が既に登録されています。",
      customer: existingByEmail || masterRow || findCustomerByEmail_(masterSheet, sanitized.email),
      eventHistory: existingEvent,
    };
  }

  const eventRow = appendEventHistory_(eventSheet, {
    event_date: eventDate,
    form_email: sanitized.email,
    receipt_required: sanitized.receipt_required,
    receipt_name: sanitized.receipt_name,
    line_user_id: existingByEmail
      ? existingByEmail.line_user_id
      : sanitized.line_user_id || "",
  });

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

// =============================================================================
// スプレッドシート操作
// =============================================================================

/**
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) {
    throw new Error("Script Property SPREADSHEET_ID が未設定です。");
  }
  return SpreadsheetApp.openById(id);
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
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  const obj = { _rowIndex: rowIndex };
  Object.keys(headerMap).forEach((key) => {
    obj[key] = values[headerMap[key]];
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
  const headerMap = getHeaderMap_(sheet);
  const emailCol = headerMap.email;
  if (emailCol === undefined) {
    throw new Error("master_customers に email 列がありません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }

  const emails = sheet.getRange(2, emailCol + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    const cellEmail = sanitizeEmailSafe_(emails[i][0]);
    if (cellEmail && cellEmail === email) {
      return rowToObject_(sheet, i + 2);
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
  const headerMap = getHeaderMap_(sheet);
  const nameCol = headerMap.user_name;
  if (nameCol === undefined) {
    throw new Error("master_customers に user_name 列がありません。");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const names = sheet.getRange(2, nameCol + 1, lastRow - 1, 1).getValues();
  const matches = [];
  for (let i = 0; i < names.length; i++) {
    const cellName = sanitizeUserName_(names[i][0]);
    if (cellName && cellName === userName) {
      matches.push(rowToObject_(sheet, i + 2));
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
 * 顧客マスタに1行追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 * @returns {Object}
 */
function appendMasterCustomer_(sheet, data) {
  ensureAllMasterHeaders_(sheet);

  const row = [
    data.line_user_id || "",
    data.line_display_name || "",
    data.user_name,
    data.company_name || "",
    data.position_category || "",
    data.position_name || "",
    data.email,
    data.phone_number || "",
    data.referrer || "",
    data.status || CustomerStatus.ACTIVE,
    new Date(),
  ];

  sheet.appendRow(row);
  const newRowIndex = sheet.getLastRow();
  return rowToObject_(sheet, newRowIndex);
}

/**
 * イベント参加履歴に1行追加
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} data
 * @returns {Object}
 */
function appendEventHistory_(sheet, data) {
  ensureHeaders_(sheet, EVENT_HEADERS);

  const row = [
    new Date(),
    data.event_date,
    data.form_email,
    data.receipt_required || "不要",
    data.receipt_name || "",
    data.line_user_id || "",
    "未",
    "未",
    0,
  ];

  sheet.appendRow(row);
  const newRowIndex = sheet.getLastRow();
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
  MASTER_HEADERS.forEach((header) => {
    if (headerMap[header] === undefined) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(header);
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
  ],
  user_name: ["氏名", "お名前", "名前", "user_name", "おなまえ"],
  email: [
    "メールアドレス",
    "メール",
    "email",
    "Email",
    "Eメール",
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
    "receipt_required",
    "領収書希望",
  ],
  receipt_name: ["領収書宛名", "receipt_name", "領収書の宛名"],
};

/**
 * フォーム送信時トリガー（スプレッドシート連携フォーム用）
 * @param {Object} e
 */
function onFormSubmit(e) {
  try {
    const payload = parseFormSubmitToPayload_(e);
    const result = handleEventApplication_(payload);
    Logger.log(
      "フォーム申込処理完了: " + JSON.stringify(result, null, 2),
    );
  } catch (error) {
    console.error(error);
    sendLineNotification_(
      [
        "【うおの会】⚠️ フォーム処理エラー",
        "",
        error.message || String(error),
        "",
        "フォーム回答シートの設問名・必須項目をご確認ください。",
      ].join("\n"),
    );
    throw error;
  }
}

/**
 * フォーム送信イベントを event_application payload に変換
 * @param {Object} e
 * @returns {Object}
 */
function parseFormSubmitToPayload_(e) {
  if (!e) {
    throw new Error("フォーム送信イベントが空です。");
  }

  if (e.namedValues) {
    return mapNamedValuesToPayload_(e.namedValues);
  }

  if (e.range) {
    return mapSheetRowToPayload_(e.range.getSheet(), e.range.getRow());
  }

  throw new Error("フォーム回答データを取得できませんでした。");
}

/**
 * @param {Object<string, string[]>} namedValues
 * @returns {Object}
 */
function mapNamedValuesToPayload_(namedValues) {
  const payload = {};
  Object.keys(FORM_FIELD_ALIASES).forEach((fieldKey) => {
    const value = resolveFormField_(namedValues, FORM_FIELD_ALIASES[fieldKey]);
    if (value !== "") {
      payload[fieldKey] = value;
    }
  });
  return finalizeFormPayload_(payload);
}

/**
 * 回答シートの1行から payload を生成
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowIndex
 * @returns {Object}
 */
function mapSheetRowToPayload_(sheet, rowIndex) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map((h) => String(h || "").trim());
  const values = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const namedValues = {};
  headers.forEach((header, index) => {
    if (header) {
      namedValues[header] = [String(values[index] || "")];
    }
  });

  return mapNamedValuesToPayload_(namedValues);
}

/**
 * @param {Object<string, string[]>} namedValues
 * @param {string[]} aliases
 * @returns {string}
 */
function resolveFormField_(namedValues, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i];
    if (namedValues[alias] && namedValues[alias][0] !== undefined) {
      return String(namedValues[alias][0]).trim();
    }
  }

  const keys = Object.keys(namedValues);
  for (let i = 0; i < aliases.length; i++) {
    const alias = aliases[i].toLowerCase();
    for (let j = 0; j < keys.length; j++) {
      const key = keys[j];
      if (key.toLowerCase().indexOf(alias) >= 0) {
        return String(namedValues[key][0] || "").trim();
      }
    }
  }

  return "";
}

/**
 * @param {Object} payload
 * @returns {Object}
 */
function finalizeFormPayload_(payload) {
  const props = PropertiesService.getScriptProperties();

  if (!payload.event_date) {
    const defaultDate =
      props.getProperty("DEFAULT_EVENT_DATE") ||
      props.getProperty("FORM_EVENT_DATE") ||
      "";
    if (defaultDate) {
      payload.event_date = defaultDate;
    }
  }

  payload.event_date = normalizeEventDate_(payload.event_date);

  if (!payload.user_name) {
    throw new Error("フォームに「氏名」が見つかりません。");
  }
  if (!payload.email) {
    throw new Error("フォームに「メールアドレス」が見つかりません。");
  }

  if (!payload.receipt_required) {
    payload.receipt_required = "不要";
  }

  return payload;
}

/**
 * イベント日を YYYYMMDD 形式に正規化
 * @param {*} value
 * @returns {string}
 */
function normalizeEventDate_(value) {
  const str = String(value || "").trim();
  if (!str) {
    throw new Error(
      "イベント日が未設定です。フォームに「参加希望日」を追加するか、Script Property DEFAULT_EVENT_DATE を設定してください。",
    );
  }

  const digits = toHalfWidth_(str).replace(/[^\d]/g, "");
  if (digits.length === 8) {
    return digits;
  }

  const match = str.match(/(\d{4})[年\/\-.](\d{1,2})[月\/\-.](\d{1,2})/);
  if (match) {
    const y = match[1];
    const m = String(match[2]).padStart(2, "0");
    const d = String(match[3]).padStart(2, "0");
    return `${y}${m}${d}`;
  }

  throw new Error(`イベント日の形式が正しくありません: ${value}`);
}

/**
 * フォーム送信トリガーをインストール（GAS エディタから1回実行）
 */
function installFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(
    (trigger) => trigger.getHandlerFunction() === "onFormSubmit",
  );
  if (exists) {
    Logger.log("onFormSubmit トリガーは既に設定済みです。");
    return;
  }

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(getSpreadsheet_())
    .onFormSubmit()
    .create();

  Logger.log("onFormSubmit トリガーをインストールしました。");
}

/**
 * フォーム送信トリガーを削除
 */
function uninstallFormSubmitTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  Logger.log(`onFormSubmit トリガーを ${removed} 件削除しました。`);
}

/**
 * 回答シートの最終行を手動処理（テスト・再処理用）
 */
function processLatestFormResponse() {
  const sheet = getFormResponseSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("フォーム回答がありません。");
  }
  const payload = mapSheetRowToPayload_(sheet, lastRow);
  const result = handleEventApplication_(payload);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * フォーム回答シートを取得
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
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
      name.indexOf("Form Responses") >= 0
    ) {
      return sheets[i];
    }
  }

  throw new Error(
    "フォーム回答シートが見つかりません。Script Property FORM_RESPONSE_SHEET_NAME を設定してください。",
  );
}

// =============================================================================
// セットアップ・テスト用（エディタから手動実行可）
// =============================================================================

/**
 * スプレッドシートにヘッダー行を初期化
 */
function setupSpreadsheetHeaders() {
  ensureHeaders_(getSheet_(SHEET_MASTER), MASTER_HEADERS);
  ensureHeaders_(getSheet_(SHEET_EVENTS), EVENT_HEADERS);
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
