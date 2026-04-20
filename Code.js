// ==========================================
// 1. 定数・設定・ユーティリティ (基盤)
// ==========================================
// 【重要】スプレッドシートIDをここに貼り付ける
const SPREADSHEET_ID = "1QSkz0k0g9N58eonBuUH4Jn421p1o9-MAnay5JcvC5qM";
const SHORT_TERM_FOLDER_ID = "1GIdpZS8m49nFdKv0cGT_E-L98ivrVNvg";
const LONG_TERM_FOLDER_ID = "1J9YzLcVLDPH9JL_RnIjtyDpw-fFDeA9N";
const UPLOAD_FOLDER_ID = "1nm64X4tSsS-sOeYp393fmanKRavi8ZTi";
const DIST_FOLDER_ID = "1gtCQ2BQtcSvAUMuTESQxsKRAAyUWZmz6";
const MASTER_PASSWORD = "t8E!C5rA2kFs";
const INTERNAL_REG_CODE = "FES7413";
const EXTERNAL_REG_CODE = "FES0000";
const SHEET_NAME = {
  INFO: "備品情報",       // A.備品情報
  HISTORY: "操作履歴",    // B.操作履歴
  RESERVATION: "予約情報", // C.予約情報
  TAG: "タグ情報",        // D.タグ情報
  USER_MGMT: "利用者管理", // E.利用者管理
  PROJECTS: "企画データ",  // Phase 1-3
  SETTINGS: "システム設定",
  FILES: "配布資料",
  CONSTANTS: "入力規制一覧",
  INQUIRIES: "問い合わせ",
  NOTIFICATIONS: "通知",
  SALES_CONFIG: "店舗設定",
  SALES_TRANSACTIONS: "売上台帳",
  SALES_EXPENSES: "経費台帳",
  FAQ: "FAQ",
  FES_EQUIP: "学祭備品設定",
  ATTENDANCE: "出欠管理",
};
const CACHE_EXPIRATION = 21600;
const CACHE_CHUNK_SIZE = 90 * 1024;
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function getSheet(sheetName) {
  return getSpreadsheet().getSheetByName(sheetName);
}
/**
 * 安全なJSONパース関数
 */
function safeJsonParse(jsonString, defaultValue = {}) {
  try {
    return jsonString ? JSON.parse(jsonString) : defaultValue;
  } catch (e) {
    console.error("JSON Parse Error: " + e, jsonString);
    return defaultValue;
  }
}
/**
 * 排他制御ラッパー関数
 */
function withLock(callback) {
  const lock = LockService.getScriptLock();
  try {
    const success = lock.tryLock(3000); // 3秒待機
    if (!success) {
      throw new Error("サーバーが混み合っています。しばらく待ってから再試行してください。");
    }
    return callback();
  } catch (e) {
    console.error("Lock Error: " + e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}
/**
 * 大容量キャッシュ用ヘルパー関数
 */
function setLargeCache(key, valueStr) {
  const cache = CacheService.getScriptCache();
  const chunks = [];
  for (let i = 0; i < valueStr.length; i += CACHE_CHUNK_SIZE) {
    chunks.push(valueStr.substring(i, i + CACHE_CHUNK_SIZE));
  }
  const chunkInfo = { numChunks: chunks.length };
  cache.put(key, JSON.stringify(chunkInfo), CACHE_EXPIRATION);
  chunks.forEach((chunk, index) => {
    cache.put(key + "_chunk_" + index, chunk, CACHE_EXPIRATION);
  });
}
function getLargeCache(key) {
  const cache = CacheService.getScriptCache();
  const infoStr = cache.get(key);
  if (!infoStr) return null;
  
  const info = JSON.parse(infoStr);
  let fullStr = "";
  for (let i = 0; i < info.numChunks; i++) {
    const chunk = cache.get(key + "_chunk_" + i);
    if (!chunk) return null;
    fullStr += chunk;
  }
  return fullStr;
}
function clearLargeCache(key) {
  const cache = CacheService.getScriptCache();
  const infoStr = cache.get(key);
  if (!infoStr) return;
  
  const info = JSON.parse(infoStr);
  cache.remove(key);
  for (let i = 0; i < info.numChunks; i++) {
    cache.remove(key + "_chunk_" + i);
  }
}
function forceClearAllCaches() {
  const cacheKeys = [
    "RESERVATION_BASE_DATA",
    "INQUIRIES_BASE_DATA",
    "NOTICES_BASE_DATA",
    "FILES_DATA",
    "USER_LIST_DATA",
    "SETTINGS_DATA",
    "SYSTEM_SETTINGS_DATA",
    "FESTIVAL_EQUIP_LIST_DATA",
    "ITEMS_LIST_DATA",
    "OP_HISTORY_DATA",
    "ADMIN_PROJECT_LIST",
    "AVAILABLE_SHOPS_DATA",
    "FAQ_LIST_DATA",
    "CONSTANTS_BASE_DATA"
  ];
  cacheKeys.forEach(key => {
    try {
      clearLargeCache(key);
    } catch (e) {
      console.warn(key + " の削除に失敗: " + e.message);
    }
  });
  console.log("すべてのキャッシュを強制削除しました。");
}

// ==========================================
// 2. ユーザー認証・基本取得系
// ==========================================
function getUserInfo() {
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) {
    return { role: 'ゲスト', name: 'ゲスト', dept: '未登録' };
  }
  const userSheet = getSheet(SHEET_NAME.USER_MGMT);
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[3] && row[3].toString().toLowerCase() === userEmail.toLowerCase()) {
      const lastName = row[1] ? row[1].toString().trim() : '';
      const firstName = row[2] ? row[2].toString().trim() : '';
      const role = row[5] ? row[5].toString().trim() : '';
      const deptType = row[6] ? row[6].toString().trim() : ''; // G列: 所属区分
      const deptName = row[7] ? row[7].toString().trim() : ''; // H列: 所属名
      return {
        role: role || 'ゲスト',
        name: (lastName + ' ' + firstName).trim() || '登録ユーザー',
        dept: deptName || deptType || '未登録'
      };
    }
  }
  return { role: 'ゲスト', name: '未登録ユーザー', dept: '未登録' };
}
function getUserId() {
  return Session.getActiveUser().getEmail();
}
function doGet(e) {
  const userInfo = getUserInfo();
  const userRole = userInfo.role;
  const userName = userInfo.name;
  const userDept = userInfo.dept;
  const versionInfo = "version 1.1.6";
  const appUrl = ScriptApp.getService().getUrl();
  const viewMode = e.parameter.view;
  const isAdminView = (viewMode === 'external');
  const isMaintenance = PropertiesService.getScriptProperties().getProperty("MAINTENANCE_MODE") === "true";
  if (isMaintenance && !isAdminView && userRole !== 'オーナー' && userRole !== '管理者') {
    return HtmlService.createHtmlOutput(`
      <div style="font-family:sans-serif; text-align:center; padding:50px;">
        <h1>🚧 メンテナンス中 🚧</h1>
        <p>現在、システムメンテナンスを行っております。</p>
        <p>終了までしばらくお待ちください。</p>
      </div>`)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setTitle("FesTrack-メンテナンス中");
  }
  const ICON_ID = "1VdotAk7j7moO5AUqJXTyMncoo4r8oTFB";
  let iconData = "";
  try {
    const blob = DriveApp.getFileById(ICON_ID).getBlob();
    iconData = `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch (e) {
    iconData = "https://ssl.gstatic.com/docs/doclist/images/infinite_arrow_favicon_5.ico";
  }
  let templateFile = 'SimpleApp';
  let pageTitle = 'FesTrack';
  if (isAdminView) {
    templateFile = 'ExternalApp';
    pageTitle = 'FesTrack Manager';
  } else if (userRole === 'オーナー' || userRole === '管理者') {
    templateFile = 'AdminApp';
    pageTitle = 'FesTrack Manager';
  } else if (userRole === '外部利用者') {
    templateFile = 'ExternalApp';
    pageTitle = 'FesTrack';
  }
  const template = HtmlService.createTemplateFromFile(templateFile);
  template.userRole = userRole;
  template.userName = userName;
  template.userDept = userDept;
  template.versionInfo = versionInfo;
  template.appUrl = appUrl;
  template.iconData = iconData;
  template.isAdminView = isAdminView;
  const htmlOutput = template.evaluate();
  htmlOutput
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://drive.google.com/uc?id=1VdotAk7j7moO5AUqJXTyMncoo4r8oTFB&.png');
  return htmlOutput;
}
function generateHistoryId() {
  const timestamp = Date.now().toString(36);
  const randomStr = Math.random().toString(36).substring(2, 5);
  return "H-" + timestamp + randomStr;
}
function recordOperationHistory({targetId, type, from = "-", to = "-", remarks = "-"}) {
  if (!targetId || !type) {
    console.error("履歴記録エラー: targetId または type が不足しています", {targetId, type});
    return; 
  }
  const sheet = getSheet(SHEET_NAME.HISTORY);
  const now = new Date();
  const historyId = generateHistoryId();
  const executor = Session.getActiveUser().getEmail() || "Unknown User";
  // A:履歴ID, B:対象ID, C:種別, D:日時, E:移動元, F:移動先, G:実行者, H:特記事項
  sheet.appendRow([
    historyId,
    targetId,
    type,
    now,
    from,
    to,
    executor,
    remarks
  ]);
}

// ==========================================
// 3. 備品管理・操作系 (LockService導入 & Batch処理)
// ==========================================
function processCheckout(itemIds, organizationName) {
  if (!itemIds || itemIds.length === 0 || !organizationName) {
    throw new Error("データが不足しています。");
  }
  return withLock(() => {
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const historySheet = getSheet(SHEET_NAME.HISTORY);
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    const infoData = infoSheet.getDataRange().getValues();
    const idMap = new Map();
    infoData.forEach((row, index) => {
      idMap.set(row[0].toString(), index); // A列: ID
    });
    const historyRowsToAdd = [];
    const userEmail = Session.getActiveUser().getEmail();
    const stateCol = 3;   // D列
    const locationCol = 4; // E列
    const timestampCol = 5; // F列
    const orgCol = 6;      // G列
    const initLocCol = 7;  // H列
    for (const id of itemIds) {
      const rowIndex = idMap.get(id.toString());
      if (rowIndex !== undefined) {
        infoData[rowIndex][stateCol] = "貸出中";
        infoData[rowIndex][locationCol] = organizationName;
        infoData[rowIndex][timestampCol] = now;
        infoData[rowIndex][orgCol] = organizationName;
        historyRowsToAdd.push([
          generateHistoryId(), // A: 履歴ID
          id,                       // B: 対象ID
          "貸出",                    // C: 操作種別
          timestampString,          // D: 操作日時
          infoData[rowIndex][initLocCol], // E: 移動元(初期場所)
          organizationName,         // F: 移動先(貸出先)
          userEmail,                // G: 実行者
          ""                        // H: 特記事項
        ]);
      } else {
        throw new Error("備品ID " + id + " が見つかりませんでした。");
      }
    }
    infoSheet.getDataRange().setValues(infoData);
    if (historyRowsToAdd.length > 0) {
      historySheet.getRange(
        historySheet.getLastRow() + 1,
        1,
        historyRowsToAdd.length,
        historyRowsToAdd[0].length
      ).setValues(historyRowsToAdd);
    }
    clearLargeCache("ITEMS_LIST_DATA");
    clearLargeCache("OP_HISTORY_DATA");
    return itemIds.length + " 件の貸出処理が完了しました。";
  });
}
function processReturn(itemIds) {
  if (!itemIds || itemIds.length === 0) throw new Error("データが不足しています。");
  return withLock(() => {
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const historySheet = getSheet(SHEET_NAME.HISTORY);
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    const infoData = infoSheet.getDataRange().getValues();
    const idMap = new Map();
    infoData.forEach((row, index) => idMap.set(row[0].toString(), index));
    const historyRowsToAdd = [];
    const userEmail = Session.getActiveUser().getEmail();
    const stateCol = 3;
    const locationCol = 4;
    const timestampCol = 5;
    const orgCol = 6;
    const initialLocationCol = 7;
    for (const id of itemIds) {
      const rowIndex = idMap.get(id.toString());
      if (rowIndex !== undefined) {
        const previousLocation = infoData[rowIndex][locationCol];
        const initLocation = infoData[rowIndex][initialLocationCol];
        infoData[rowIndex][stateCol] = "利用可能";
        infoData[rowIndex][locationCol] = initLocation;
        infoData[rowIndex][orgCol] = "";
        infoData[rowIndex][timestampCol] = now;
        historyRowsToAdd.push([
          generateHistoryId(), // A
          id,                       // B
          "返却",                    // C
          timestampString,          // D
          previousLocation,         // E: 返却前の場所
          initLocation,             // F: 返却後の定位置
          userEmail,                // G
          ""                        // H
        ]);
      }
    }
    infoSheet.getDataRange().setValues(infoData);
    if (historyRowsToAdd.length > 0) {
      historySheet.getRange(
        historySheet.getLastRow() + 1,
        1,
        historyRowsToAdd.length,
        historyRowsToAdd[0].length
      ).setValues(historyRowsToAdd);
    }
    clearLargeCache("ITEMS_LIST_DATA");
    clearLargeCache("OP_HISTORY_DATA");
    return itemIds.length + " 件の返却処理が完了しました。";
  });
}
function processMove(itemIds, newLocation) {
  if (!itemIds || itemIds.length === 0) throw new Error("備品が選択されていません。");
  if (!newLocation || newLocation.trim() === "") throw new Error("移動先の場所が入力されていません。");
  return withLock(() => {
    const trimmedLocation = newLocation.trim();
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const historySheet = getSheet(SHEET_NAME.HISTORY);
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    const infoData = infoSheet.getDataRange().getValues();
    const idMap = new Map();
    infoData.forEach((row, index) => idMap.set(row[0].toString(), index));
    const historyRowsToAdd = [];
    let moveCount = 0;
    let returnCount = 0;
    const userEmail = Session.getActiveUser().getEmail();
    const stateCol = 3;
    const locationCol = 4;
    const timestampCol = 5;
    const orgCol = 6;
    const initialLocationCol = 7;
    for (const id of itemIds) {
      const rowIndex = idMap.get(id.toString());
      if (rowIndex !== undefined) {
        const previousLocation = infoData[rowIndex][locationCol];
        const itemInitialLocation = infoData[rowIndex][initialLocationCol];
        let operationType, finalLocation;
        if (trimmedLocation === itemInitialLocation) {
          operationType = "返却";
          finalLocation = itemInitialLocation;
          returnCount++;
        } else {
          operationType = "移動";
          finalLocation = trimmedLocation;
          moveCount++;
        }
        infoData[rowIndex][stateCol] = "利用可能";
        infoData[rowIndex][locationCol] = finalLocation;
        infoData[rowIndex][orgCol] = "";
        infoData[rowIndex][timestampCol] = now;
        historyRowsToAdd.push([
          generateHistoryId(), // A
          id,                       // B
          operationType,            // C
          timestampString,          // D
          previousLocation,         // E
          finalLocation,            // F
          userEmail,                // G
          ""                        // H
        ]);
      }
    }
    infoSheet.getDataRange().setValues(infoData);
    if (historyRowsToAdd.length > 0) {
      historySheet.getRange(
        historySheet.getLastRow() + 1,
        1,
        historyRowsToAdd.length,
        historyRowsToAdd[0].length
      ).setValues(historyRowsToAdd);
    }
    clearLargeCache("ITEMS_LIST_DATA");
    clearLargeCache("OP_HISTORY_DATA");
    return `${moveCount} 件の移動、${returnCount} 件の返却処理が完了しました。`;
  });
}
function addNewItem(itemData) {
  return withLock(() => {
    if (!itemData?.id || !itemData?.name) throw new Error("必須項目が不足しています。");
    const newItemId = itemData.id.trim().toUpperCase();
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const idValues = infoSheet.getRange("A2:A" + infoSheet.getLastRow()).getValues();
    const idSet = new Set(idValues.map(r => r[0].toString()));
    if (idSet.has(newItemId)) {
      throw new Error(`備品ID「${newItemId}」は既に使用されています。`);
    }
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    const adminEmail = Session.getActiveUser().getEmail();
    infoSheet.appendRow([
      newItemId, itemData.name, itemData.category, "利用可能",
      itemData.location, timestampString, "", itemData.location, itemData.remarks || ""
    ]);
    const tagSheet = getSheet(SHEET_NAME.TAG);
    tagSheet.appendRow([
      newItemId, itemData.name, itemData.printLot || "N/A",
      timestampString, itemData.attachLocation || "", "正常"
    ]);
    const historySheet = getSheet(SHEET_NAME.HISTORY);
    historySheet.appendRow([
      generateHistoryId(), // A
      newItemId,                // B
      "新規登録",               // C
      timestampString,          // D
      "-",                      // E
      itemData.location,        // F: 登録先
      adminEmail,               // G
      itemData.remarks || ""    // H: 特記事項
    ]);
    return `備品ID「${newItemId}」を正常に登録しました。`;
  });
}
function updateItemDetails(itemData) {
  return withLock(() => {
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const tagSheet = getSheet(SHEET_NAME.TAG);
    const infoRow = parseInt(itemData.infoRow, 10);
    const tagRow = parseInt(itemData.tagRow, 10);
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    infoSheet.getRange(infoRow, 2, 1, 4).setValues([[itemData.name, itemData.category, itemData.status, itemData.location]]);
    infoSheet.getRange(infoRow, 8, 1, 2).setValues([[itemData.initialLocation, itemData.notes]]);
    infoSheet.getRange(infoRow, 6).setValue(now);
    if (tagRow > 0) {
      tagSheet.getRange(tagRow, 2, 1, 1).setValue(itemData.name);
      tagSheet.getRange(tagRow, 3, 1, 1).setValue(itemData.printLot);
      tagSheet.getRange(tagRow, 5, 1, 2).setValues([[itemData.attachLocation, itemData.tagStatus]]);
    }
    getSheet(SHEET_NAME.HISTORY).appendRow([
      generateHistoryId(),   // A
      itemData.itemId,            // B
      "情報更新",                 // C
      timestampString,            // D
      "管理画面",                 // E
      "管理画面",                 // F
      Session.getActiveUser().getEmail(), // G
      "備品マスタ情報の更新"      // H
    ]);
    clearLargeCache("ITEMS_LIST_DATA");
    clearLargeCache("OP_HISTORY_DATA");
    return `備品ID「${itemData.itemId}」を更新しました。`;
  });
}
function checkItemAvailability(itemId) {
  try {
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const lastRow = infoSheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "データがありません", id: itemId };
    const data = infoSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const itemMap = new Map();
    data.forEach(row => {
      itemMap.set(row[0].toString(), row[3]); // ID -> Status
    });
    const status = itemMap.get(itemId.toString());
    if (status) {
      if (status === "利用可能") {
        return { success: true, message: "OK", id: itemId };
      } else {
        return { success: false, message: "この備品は現在「" + status + "」です。", id: itemId };
      }
    }
    return { success: false, message: "備品IDがマスタに存在しません。", id: itemId };
  } catch (e) {
    Logger.log("checkItemAvailabilityエラー: " + e);
    return { success: false, message: "ステータス確認中にエラーが発生しました。", id: itemId };
  }
}

// ==========================================
// 4. マスタデータ取得 (Cache推奨だが簡易化)
// ==========================================
function getBaseConstantsData() {
  const cacheKey = "CONSTANTS_BASE_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.CONSTANTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  setLargeCache(cacheKey, JSON.stringify(data));
  return data;
}
function getCategoryList() {
  return getMasterListColumn(0); // A列
}
function getOrganizationList() {
  try {
    const data = getBaseConstantsData();
    if (data.length < 2) return [];
    const results = [];
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][1] || "").trim(); // B列(団体名)
      const id = String(data[i][5] || "").trim();   // F列(団体ID)
      if (name) {
        results.push({ id: id || name, name: name });
      }
    }
    return results;
  } catch (e) {
    console.error("Org list load error: " + e);
    return [];
  }
}
function getLocationList() {
  return getMasterListColumn(2); // C列
}
function getMasterListColumn(colIndex) {
  try {
    const data = getBaseConstantsData();
    if (data.length < 2) return [];
    const results = [];
    for (let i = 1; i < data.length; i++) {
      const val = data[i][colIndex];
      if (val) results.push(val);
    }
    return results;
  } catch (e) {
    console.error("Master list load error: " + e);
    return [];
  }
}

// ==========================================
// 5. 予約・申請系 (withLock適用)
// ==========================================
function submitReservation(d) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.RESERVATION);
    const user = Session.getActiveUser().getEmail();
    const userInfo = getUserInfo();
    const shortId = Utilities.getUuid().slice(0, 13);
    const newRow = [
      shortId,
      d.organizationName,
      d.category,
      d.quantity,
      new Date(d.startTime),
      new Date(d.endTime),
      "審査中",
      userInfo.name,
      user,
      "",
      d.usagePurpose,
      "",
      "",
      ""
    ];
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, newRow.length).setValues([newRow]);
    clearLargeCache("RESERVATION_BASE_DATA");
    return "予約申請を受け付けました。";
  });
}
function submitAppeal(reservationId, reason) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.RESERVATION);
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === reservationId) {
        const currentStatus = rows[i][6];
        if (currentStatus !== "却下") {
          throw new Error("異議申立は「却下」ステータスの場合のみ可能です。");
        }
        sheet.getRange(i + 1, 7).setValue("再審中");
        sheet.getRange(i + 1, 13).setValue(reason);
        clearLargeCache("RESERVATION_BASE_DATA");
        return "異議申立を行いました。管理者の再審査をお待ちください。";
      }
    }
    throw new Error("予約IDが見つかりません。");
  });
}
function cancelReservation(reservationId) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.RESERVATION);
    const rows = sheet.getDataRange().getValues();
    const user = Session.getActiveUser().getEmail();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === reservationId) {
        if (rows[i][8] !== user) throw new Error("権限がありません。");
        sheet.deleteRow(i + 1);
        clearLargeCache("RESERVATION_BASE_DATA");
        return "予約申請を取り消しました。";
      }
    }
    throw new Error("予約IDが見つかりません。");
  });
}

// ==========================================
// 6. ユーザー情報管理 (Batch Update)
// ==========================================
function getCurrentUserInfoDetails() {
  const userInfo = getUserInfo();
  if (userInfo.role === 'ゲスト') {
    return { name: "ゲスト", email: "N/A", departmentType: "N/A", departmentName: "N/A" };
  }
  const userEmail = Session.getActiveUser().getEmail();
  const userSheet = getSheet(SHEET_NAME.USER_MGMT);
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[3] && row[3].toString().toLowerCase() === userEmail.toLowerCase()) {
      return {
        name: (row[1] + ' ' + row[2]).trim(),
        email: row[3],
        departmentType: row[6] || "（未登録）",
        departmentName: row[7] || "（未登録）"
      };
    }
  }
  return { name: "エラー", email: userEmail, departmentType: "取得失敗", departmentName: "取得失敗" };
}
function updateCurrentUserInfo(newData) {
  return withLock(() => {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("ユーザーセッションが見つかりません。");
    const userSheet = getSheet(SHEET_NAME.USER_MGMT);
    const userList = userSheet.getDataRange().getValues();
    let foundRowIndex = -1;
    for (let i = 1; i < userList.length; i++) {
      if (userList[i][3] && userList[i][3].toString().toLowerCase() === userEmail.toLowerCase()) {
        foundRowIndex = i;
        break;
      }
    }
    if (foundRowIndex === -1) throw new Error("更新対象のユーザーがシートに見つかりません。");
    const sheetRow = foundRowIndex + 1;
    // B, C列 (2列分)
    userSheet.getRange(sheetRow, 2, 1, 2).setValues([[newData.lastName, newData.firstName]]);
    // G, H列 (2列分)
    userSheet.getRange(sheetRow, 7, 1, 2).setValues([[newData.departmentType, newData.departmentName]]);
    clearLargeCache("USER_LIST_DATA");
    return "会員情報を更新しました。";
  });
}
function deleteCurrentUser(password = null) {
  return withLock(() => {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("ユーザー情報を取得できませんでした。");
    const userSheet = getSheet(SHEET_NAME.USER_MGMT);
    const lastRow = userSheet.getLastRow();
    if (lastRow < 2) throw new Error("ユーザーデータがありません。");
    const range = userSheet.getRange(2, 4, lastRow - 1, 1); // D列(Email)を取得
    const values = range.getValues();
    let targetRow = -1;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString().toLowerCase() === userEmail.toLowerCase()) {
        targetRow = i + 2; // 行番号補正
        break;
      }
    }
    if (targetRow === -1) throw new Error("ユーザーが見つかりません。");
    const currentRole = userSheet.getRange(targetRow, 6).getValue();
    if (currentRole === "オーナー") {
      if (password !== MASTER_PASSWORD) {
        throw new Error("オーナー権限のアカウント削除には正しい管理者パスワードが必要です。");
      }
    }
    userSheet.getRange(targetRow, 6).setValue("ゲスト"); // F列をゲストに変更
    userSheet.getRange(targetRow, 10).setValue(new Date()); // J列に退会日時を記録
    clearLargeCache("USER_LIST_DATA");
    return "利用登録を解除しました。";
  });
}

// ==========================================
// 7. 新規登録・アイテム管理 (Batch & Lock)
// ==========================================
function swapItemTag(oldTagId, newTagId) {
  return withLock(() => {
    const tagSheet = getSheet(SHEET_NAME.TAG);
    let targetRow = null;
    let found = tagSheet.getRange("C:C").createTextFinder(oldTagId).matchEntireCell(true).findNext(); // PrintLot検索
    if (!found) {
       found = tagSheet.getRange("A:A").createTextFinder(oldTagId).matchEntireCell(true).findNext();
    }
    if (!found) throw new Error("交換元のタグまたは備品が見つかりません。");
    targetRow = found.getRow();
    const newFound = tagSheet.getRange("C:C").createTextFinder(newTagId).matchEntireCell(true).findNext();
    if (newFound) throw new Error("新しいタグIDは既に他の備品で使用されています。");
    tagSheet.getRange(targetRow, 3).setValue(newTagId);
    const itemName = tagSheet.getRange(targetRow, 2).getValue();
    clearLargeCache("ITEMS_LIST_DATA");
    clearLargeCache("OP_HISTORY_DATA");
    return `「${itemName}」のタグを更新しました。\n新タグID: ${newTagId}`;
  });
}
function registerNewUser(userData) {
  return withLock(() => {
    // 1. 認証コードチェック
    const inputCode = userData.authCode;
    let isValid = false;
    if (userData.departmentType === "学祭委員" && inputCode === INTERNAL_REG_CODE) isValid = true;
    if (userData.departmentType === "外部団体" && inputCode === EXTERNAL_REG_CODE) isValid = true;
    if (!isValid) throw new Error("認証コードが正しくありません。");
    // 2. メールアドレス取得
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) throw new Error("Googleアカウントのメールアドレスを取得できませんでした。");
    const userSheet = getSheet(SHEET_NAME.USER_MGMT);
    const found = userSheet.getRange("D:D").createTextFinder(userEmail).matchEntireCell(true).findNext();
    if (found) {
      const row = found.getRow();
      const currentRole = userSheet.getRange(row, 6).getValue();
      if (currentRole === "ゲスト") {
        userSheet.getRange(row, 2).setValue(userData.lastName);       // B列
        userSheet.getRange(row, 3).setValue(userData.firstName);      // C列
        userSheet.getRange(row, 5).setValue(new Date());              // E列
        userSheet.getRange(row, 6).setValue(userData.role);           // F列 (Guestから正規のロールへ戻る)
        userSheet.getRange(row, 7).setValue(userData.departmentType); // G列
        userSheet.getRange(row, 8).setValue(userData.departmentName); // H列
        return "アカウントを再有効化しました。登録が完了しました。";
      } else {
        throw new Error("このメールアドレスは既に登録されています。");
      }
    } else {
      const newRow = [
        "",
        userData.lastName,       // B列
        userData.firstName,      // C列
        userEmail,               // D列
        new Date(),              // E列
        userData.role,           // F列
        userData.departmentType, // G列
        userData.departmentName  // H列
      ];
      userSheet.appendRow(newRow);
      clearLargeCache("USER_LIST_DATA");
      return "登録が完了しました。";
    }
  });
}

// ==========================================
// 8. 検索・参照系
// ==========================================
function getItemDetails(itemId) {
  const infoSheet = getSheet(SHEET_NAME.INFO);
  const tagSheet = getSheet(SHEET_NAME.TAG);
  const infoFinder = infoSheet.getRange("A:A").createTextFinder(itemId).matchEntireCell(true).findNext();
  if (!infoFinder) throw new Error("備品情報が見つかりません。");
  const infoRowIndex = infoFinder.getRow();
  const itemInfo = infoSheet.getRange(infoRowIndex, 1, 1, infoSheet.getLastColumn()).getValues()[0];
  const tagFinder = tagSheet.getRange("A:A").createTextFinder(itemId).matchEntireCell(true).findNext();
  const tagRowIndex = tagFinder ? tagFinder.getRow() : -1;
  const itemTag = tagRowIndex > 0 ? tagSheet.getRange(tagRowIndex, 1, 1, tagSheet.getLastColumn()).getValues()[0] : [];
  return {
    itemId: itemInfo[0],
    name: itemInfo[1],
    category: itemInfo[2],
    status: itemInfo[3],
    location: itemInfo[4],
    initialLocation: itemInfo[7],
    notes: itemInfo[8],
    printLot: itemTag[2] || "",
    attachLocation: itemTag[4] || "",
    tagStatus: itemTag[5] || "",
    infoRow: infoRowIndex,
    tagRow: tagRowIndex
  };
}
function searchItems(searchText) {
  if (!searchText || searchText.trim().length < 2) throw new Error("検索キーワードは2文字以上必要です。");
  const lowerSearchText = searchText.toLowerCase().trim();
  const infoSheet = getSheet(SHEET_NAME.INFO);
  const data = infoSheet.getDataRange().getValues();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[0]).toLowerCase().includes(lowerSearchText) || String(row[1]).toLowerCase().includes(lowerSearchText)) {
      results.push({
        id: row[0], name: row[1], category: row[2], initialLocation: row[7]
      });
    }
  }
  if (results.length > 50) throw new Error("候補が多すぎます(50件以上)。");
  return results;
}
function findItemIdByTag(tagId) {
  const tagSheet = getSheet(SHEET_NAME.TAG);
  const infoSheet = getSheet(SHEET_NAME.INFO);
  const foundTag = tagSheet.getRange("C:C").createTextFinder(tagId).matchEntireCell(true).findNext();
  if (foundTag) {
    const tagRow = foundTag.getRow();
    const itemId = tagSheet.getRange(tagRow, 1).getValue();
    const itemName = tagSheet.getRange(tagRow, 2).getValue();
    const foundInfo = infoSheet.getRange("A:A").createTextFinder(itemId).matchEntireCell(true).findNext();
    let itemStatus = "不明";
    if (foundInfo) {
      itemStatus = infoSheet.getRange(foundInfo.getRow(), 4).getValue(); 
    }
    return { success: true, itemId: itemId, name: itemName, status: itemStatus, message: "OK" };
  }
  const foundId = infoSheet.getRange("A:A").createTextFinder(tagId).matchEntireCell(true).findNext();
  if (foundId) {
     const row = foundId.getRow();
     const name = infoSheet.getRange(row, 2).getValue();
     const itemStatus = infoSheet.getRange(row, 4).getValue(); 
     return { success: true, itemId: tagId, name: name, status: itemStatus, message: "ID直接入力" };
  }
  return { success: false, message: "未登録のタグです" };
}

// ==========================================
// 9. Dashboard & Visualization
// ==========================================
function getDashboardData() {
  const items = getAllItemsList();
  const stats = { total: 0, onLoan: 0, available: 0 };
  const breakdown = {};
  items.forEach(item => {
    const cat = item.category;
    const st = item.status;
    if (cat && st && st !== "紛失" && st !== "廃棄" && st !== "管理外") {
      stats.total++;
      if (st === "貸出中") stats.onLoan++;
      else if (st === "利用可能") stats.available++;
      if (st === "貸出中" || st === "利用可能") {
        if (!breakdown[st]) breakdown[st] = {};
        if (!breakdown[st][cat]) breakdown[st][cat] = 0;
        breakdown[st][cat]++;
      }
    }
  });
  const users = getUserList(); // キャッシュ化済みのユーザーリスト取得
  const userStats = { total: users.length, admin: 0, normal: 0, external: 0 };
  users.forEach(u => {
    const role = u.role;
    if (role === "オーナー" || role === "管理者") userStats.admin++;
    else if (role === "外部利用者") userStats.external++;
    else userStats.normal++; // それ以外（利用者など）
  });
  const reserveData = getBaseReservationData();
  let pendingCount = 0;
  if (reserveData && reserveData.length > 0) {
    reserveData.forEach(r => {
      const status = r[6]; // G列: ステータス
      if (status === "審査中" || status === "再審中") pendingCount++;
    });
  }
  const inqData = getBaseInquiryData();
  let unreadMsgCount = 0;
  if (inqData && inqData.length > 1) {
    for (let i = 1; i < inqData.length; i++) {
      if (inqData[i][7] === '未読') unreadMsgCount++; // H列: ステータス
    }
  }
  return { stats, breakdown, pendingCount, unreadMsgCount, userStats };
}
function getApprovedReservations() {
  const data = getBaseReservationData();
  if (!data || data.length === 0) return [];
  const idx = { status: 6, cat: 2, qty: 3, start: 4, end: 5, org: 1 };
  const dailyTotals = {};
  data.forEach(r => {
    if (r[idx.status] === "承認" || r[idx.status] === "条件付") {
      const cat = r[idx.cat];
      const qty = Number(r[idx.qty]);
      const org = r[idx.org];
      let current = new Date(r[idx.start]);
      const end = new Date(r[idx.end]);
      current.setHours(0, 0, 0, 0);
      const loopEnd = new Date(end); 
      loopEnd.setHours(0, 0, 0, 0);
      while (current <= loopEnd) {
        const dStr = Utilities.formatDate(current, "JST", "yyyy-MM-dd");
        if (!dailyTotals[dStr]) dailyTotals[dStr] = {};
        if (!dailyTotals[dStr][cat]) dailyTotals[dStr][cat] = { total: 0, details: [] };
        dailyTotals[dStr][cat].total += qty;
        dailyTotals[dStr][cat].details.push(`${org}(${qty})`);
        current.setDate(current.getDate() + 1);
      }
    }
  });
  const events = [];
  for (const date in dailyTotals) {
    for (const cat in dailyTotals[date]) {
      const item = dailyTotals[date][cat];
      events.push({
        title: `${cat}: ${item.total}`,
        start: date,
        allDay: true,
        description: item.details.join('\n')
      });
    }
  }
  return events;
}

// ==========================================
// 10. External Dashboard & Phases
// ==========================================
function getExternalDashboardData() {
  const userEmail = Session.getActiveUser().getEmail();
  const userInfo = getUserInfo();
  let settingsData = JSON.parse(getLargeCache("SETTINGS_DATA") || "null");
  if (!settingsData) {
    settingsData = getSheet(SHEET_NAME.SETTINGS).getDataRange().getValues();
    setLargeCache("SETTINGS_DATA", JSON.stringify(settingsData));
  }
  const now = new Date();
  const projectSheet = getSheet(SHEET_NAME.PROJECTS);
  let p1Status = null, p2Status = null, p3Status = null;
  if (projectSheet) {
    const found = projectSheet.getRange("F:F").createTextFinder(userEmail).matchEntireCell(true).findNext();
    if (found) {
      const row = found.getRow();
      const statuses = projectSheet.getRange(row, 2, 1, 20).getValues()[0];
      p1Status = statuses[0];  
      p2Status = statuses[13]; 
      p3Status = statuses[19]; 
    }
  }
  const phases = {};
  let currentPhase = null;
  let nextPhase = null;
  for (let i = 1; i < settingsData.length; i++) {
    const key = settingsData[i][0];
    const start = new Date(settingsData[i][2]);
    const end = new Date(settingsData[i][3]);
    end.setHours(23, 59, 59, 999);
    const isOpen = (now >= start && now <= end);
    let myStatus = null;
    if (key === 'PHASE_1') myStatus = p1Status;
    else if (key === 'PHASE_2') myStatus = p2Status;
    else if (key === 'PHASE_3') myStatus = p3Status;
    phases[key] = {
      name: settingsData[i][1],
      start: Utilities.formatDate(start, "JST", "MM/dd"),
      end: Utilities.formatDate(end, "JST", "MM/dd"),
      isOpen: isOpen,
      isClosed: (now > end),
      userStatus: myStatus
    };
    if (isOpen) currentPhase = key;
    if ((now < start) && !nextPhase) nextPhase = key;
  }
  let files = JSON.parse(getLargeCache("FILES_DATA") || "null");
  if (!files) {
    const filesSheet = getSheet(SHEET_NAME.FILES);
    if (filesSheet && filesSheet.getLastRow() > 1) {
      const fileData = filesSheet.getRange(2, 1, filesSheet.getLastRow() - 1, 6).getValues();
      files = fileData.map(r => ({
        title: r[1], target: r[2], fileId: r[3],
        date: Utilities.formatDate(new Date(r[4]), "JST", "yyyy/MM/dd")
      }));
      setLargeCache("FILES_DATA", JSON.stringify(files));
    } else {
      files = [];
    }
  }
  const filteredFiles = files.filter(f => f.target === "全体" || f.target === userInfo.role);
  return { userName: userInfo.name, phases, currentPhase, nextPhase, files: filteredFiles };
}
function checkDuplicateRep(email, currentProjectId) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  if (!sheet || !email) return false;
  const founds = sheet.getRange("F:G").createTextFinder(email).matchEntireCell(true).findAll();
  for (let i = 0; i < founds.length; i++) {
    const row = founds[i].getRow();
    const rowId = sheet.getRange(row, 1).getValue();
    if (currentProjectId && String(rowId) === String(currentProjectId)) continue;
    return true; 
  }
  return false;
}
// ==========================================
// 企画データ一括取得（フロントエンド用）
// ==========================================
function getMyProjectData() {
  const userEmail = Session.getActiveUser().getEmail();
  const row = findProjectRowByEmail(userEmail);
  if (row) {
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    const data = sheet.getRange(row, 1, 1, 24).getValues()[0];
    const projectId = data[0];
    const approvedEquips = getApprovedEquipment(projectId);
    const hasApprovedEquip = approvedEquips.length > 0;
    let content = safeJsonParse(data[21]);
    let isContractValid = false;
    let contract = content.contractInfo;
    if (contract && contract.snapshot) {
      const currentStr = JSON.stringify(approvedEquips.sort((a,b) => a.name > b.name ? 1 : -1));
      const snapshotStr = JSON.stringify(contract.snapshot.sort((a,b) => a.name > b.name ? 1 : -1));
      isContractValid = (currentStr === snapshotStr);
    } else if (contract) {
      isContractValid = true; // 古いデータ互換用
    }
    const mainEmail = data[5];
    const subEmail = data[6];
    let currentUserRole = 'NONE';
    if (userEmail === mainEmail) {
      currentUserRole = 'MAIN';
    } else if (userEmail === subEmail) {
      currentUserRole = 'SUB';
    }
    return {
      isNew: false,
      // --- 基本情報 (Phase 1共通) ---
      projectId: data[0],
      status: data[1],
      projectType: data[2],
      projectName: data[3],
      orgName: data[4],
      mainRepEmail: mainEmail,
      subRepEmail: subEmail,
      mainRepInfo: safeJsonParse(data[7]),
      subRepInfo: safeJsonParse(data[8]),
      basicInfo: safeJsonParse(data[9]),
      locationInfo: safeJsonParse(data[12]),
      roomInfo: safeJsonParse(data[13]),
      phase1Comment: data[11], // L列
      // --- Phase 2 固有情報 ---
      p2Status: data[14],                   // O列
      proposalFileId: data[15],             // P列
      relatedApps: safeJsonParse(data[16]), // Q列
      adInfo: safeJsonParse(data[17]),      // R列
      vehicleInfo: safeJsonParse(data[18]), // S列
      equipmentInfo: safeJsonParse(data[19]),// T列
      phase2Comment: data[22],              // W列
      // --- Phase 3 固有情報 ---
      p3Status: data[20],                   // U列
      content: content,                     // V列
      phase3Comment: data[23],              // X列
      hasApprovedEquip: hasApprovedEquip,   // 備品の承認状況
      approvedEquipList: approvedEquips, // 通信削減のためリストごと返す
      contractInfo: contract,
      isContractValid: isContractValid,
      currentUserRole: currentUserRole
    };
  }
  const myInfo = getUserInfo();
  if (myInfo.name !== "ゲスト") return { isNew: true, mainRepName: myInfo.name };
  return null;
}
function savePhase1Data(data) {
  console.log("[Debug] savePhase1Data started. User:", Session.getActiveUser().getEmail());
  return withLock(() => {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const projectSheet = ss.getSheetByName(SHEET_NAME.PROJECTS);
    const userSheet = ss.getSheetByName(SHEET_NAME.USER_MGMT);
    const userEmail = Session.getActiveUser().getEmail();
    if (!projectSheet) throw new Error("Error: 企画データシートが見つかりません");
    const finder = projectSheet.getRange("F:G").createTextFinder(userEmail).matchEntireCell(true);
    const foundCell = finder.findNext(); 
    let isNew = (foundCell === null);
    let targetRow = isNew ? projectSheet.getLastRow() + 1 : foundCell.getRow();
    let projectId = isNew ? (data.projectId || Utilities.getUuid()) : projectSheet.getRange(targetRow, 1).getValue();
    const step = data.step || 'basic';
    const isDraft = data.isDraft || false;
    const timestamp = new Date();
    if (step === 'basic') {
      if (data.subRepEmail && data.subRepEmail.trim() !== "") {
        if (data.subRepEmail.trim().toLowerCase() === userEmail.toLowerCase()) {
          throw new Error("正責任者と副責任者に同じメールアドレスを設定することはできません。");
        }
      }
      const checkDuplicate = (email) => {
        if (!email) return false;
        const f = projectSheet.getRange("F:G").createTextFinder(email).matchEntireCell(true).findNext();
        return f && f.getRow() !== targetRow;
      };
      if (checkDuplicate(userEmail)) {
        throw new Error("あなたは既に他の企画の正責任者または副責任者として登録されています。\n1人につき1つの企画のみ担当できます。");
      }
      if (checkDuplicate(data.subRepEmail)) {
        throw new Error(`指定された副責任者 (${data.subRepEmail}) は、\n既に他の企画の正責任者または副責任者として登録されています。`);
      }
      if (data.orgName && userSheet) {
        const updateOrg = (email) => {
          if (!email) return;
          const f = userSheet.getRange("D:D").createTextFinder(email).matchEntireCell(true).findNext();
          if (f) userSheet.getRange(f.getRow(), 8).setValue(data.orgName);
        };
        updateOrg(userEmail);
        updateOrg(data.subRepEmail);
        
        clearLargeCache("USER_LIST_DATA"); // ユーザー情報変更によるキャッシュクリア
      }
      let p1Status = isDraft ? "一時保存" : "提出中(未完)";
      const rowData1_11 = [
        projectId,                               // A: ID
        p1Status,                                // B: P1 Status
        data.projectType,                        // C: 形態
        data.projectName,                        // D: 企画名
        data.orgName,                            // E: 団体名
        userEmail,                               // F: 正責任者
        data.subRepEmail,                        // G: 副責任者
        JSON.stringify(data.mainRepInfo || {}),  // H: 正情報
        JSON.stringify(data.subRepInfo || {}),   // I: 副情報
        JSON.stringify(data.basicInfo || {}),    // J: 基本情報
        timestamp                                // K: タイムスタンプ
      ];
      projectSheet.getRange(targetRow, 1, 1, 11).setValues([rowData1_11]);
      if (isNew) {
        projectSheet.getRange(targetRow, 15).setValue("未着手"); // O列
        projectSheet.getRange(targetRow, 21).setValue("未着手"); // U列
      }
      clearLargeCache("ADMIN_PROJECT_LIST"); // 企画データ変更によるキャッシュクリア
      return isDraft ? "一時保存しました。" : "基本情報を保存しました。";
    }
    else if (step === 'location') {
      if (isNew) throw new Error("エラー: 先に基本情報を登録してください。");
      projectSheet.getRange(targetRow, 13).setValue(JSON.stringify(data.locationInfo || {})); 
      projectSheet.getRange(targetRow, 11).setValue(timestamp);
      updatePhase1Status(targetRow);
      clearLargeCache("ADMIN_PROJECT_LIST"); // キャッシュクリア
      return "場所・時間申請を登録しました。";
    }
    else if (step === 'room') {
      if (isNew) throw new Error("エラー: 先に基本情報を登録してください。");
      projectSheet.getRange(targetRow, 14).setValue(JSON.stringify(data.roomInfo || {})); 
      projectSheet.getRange(targetRow, 11).setValue(timestamp);
      if (!isDraft) projectSheet.getRange(targetRow, 2).setValue("提出済");
      updatePhase1Status(targetRow);
      clearLargeCache("ADMIN_PROJECT_LIST"); // キャッシュクリア
      return "控室申請を登録しました。";
    }
  });
}
// ==========================================
// 11. バックアップ・その他
// ==========================================
function createBackup() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const time = Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmm");
    const filename = `Backup_FesTrack_${time}`;
    let folder;
    try {
      folder = DriveApp.getFolderById(SHORT_TERM_FOLDER_ID);
    } catch (e) {
      folder = DriveApp.getRootFolder();
    }
    DriveApp.getFileById(ss.getId()).makeCopy(filename, folder);
    return "バックアップを作成しました: " + filename;
  } catch (e) {
    return "バックアップ失敗: " + e.message;
  }
}
function fetchDataUrl(fileId) {
  if (!fileId || fileId.includes("ファイルID")) return "";
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return `data:${blob.getContentType()};base64,${Utilities.base64Encode(blob.getBytes())}`;
  } catch (e) {
    return "";
  }
}
function isMaintenanceMode() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty("MAINTENANCE_MODE") === "true";
}
function setMaintenanceMode(enable) {
  return withLock(() => {
    const props = PropertiesService.getScriptProperties();
    props.setProperty("MAINTENANCE_MODE", enable ? "true" : "false");
    return enable ? "メンテナンスモードを有効にしました。" : "メンテナンスモードを解除しました。";
  });
}
function deleteOldShortTermBackups() {
  const folder = DriveApp.getFolderById(SHORT_TERM_FOLDER_ID);
  const files = folder.getFiles();
  const threshold = new Date();
  threshold.setDate(threshold.getDate() - 30);
  while (files.hasNext()) {
    const file = files.next();
    if (file.getDateCreated() < threshold) {
      file.setTrashed(true); // ゴミ箱へ移動
    }
  }
}
function createMonthlyLongTermBackup() {
  const folder = DriveApp.getFolderById(LONG_TERM_FOLDER_ID);
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM");
  const fileName = `【長期保存】FesTrack_Backup_${dateStr}`;
  const originalFile = DriveApp.getFileById(SPREADSHEET_ID);
  const copiedFile = originalFile.makeCopy(fileName, folder);
  const backupSs = SpreadsheetApp.openById(copiedFile.getId());
  maskPersonalData(backupSs);
}
function maskPersonalData(ss) {
  const maskText = "***";
  const maskEmail = "masked@example.com";
  // --- 1. 単純な列のマスキング ---
  // 操作履歴: G列(7) = Email
  processSimpleMask(ss, SHEET_NAME.HISTORY, [7], maskEmail);
  // 予約情報: B列(2), H列(8) = 氏名 / I列(9), J列(10) = Email
  processSimpleMask(ss, SHEET_NAME.RESERVATION, [2, 8, 14], maskText);
  processSimpleMask(ss, SHEET_NAME.RESERVATION, [9, 10], maskEmail);
  // 利用者管理: B列(2)=姓, C列(3)=名 / D列(4)=Email
  processSimpleMask(ss, SHEET_NAME.USER_MGMT, [2, 3], maskText);
  processSimpleMask(ss, SHEET_NAME.USER_MGMT, [4], maskEmail);
  // システム設定: G列(7) = 氏名
  processSimpleMask(ss, SHEET_NAME.SETTINGS, [7], maskText);
  // 問い合わせ: C列(3) = 氏名 / B列(2) = Email
  processSimpleMask(ss, SHEET_NAME.INQUIRIES, [3], maskText);
  processSimpleMask(ss, SHEET_NAME.INQUIRIES, [2], maskEmail);
  // 売上台帳: D列(4) = 学籍番号
  processSimpleMask(ss, SHEET_NAME.SALES_TRANSACTIONS, [4], maskText);
  // --- 2. 複雑なJSONを含むマスキング (企画データ) ---
  const projectSheet = ss.getSheetByName(SHEET_NAME.PROJECTS);
  if (projectSheet) {
    const lastRow = projectSheet.getLastRow();
    if (lastRow > 1) {
      const range = projectSheet.getRange(2, 6, lastRow - 1, 17); 
      const values = range.getValues();
      const maskedValues = values.map(row => {
        row[0] = maskEmail;
        row[1] = maskEmail;
        row[2] = redactJson(row[2], ["studentId", "name", "tel"]);
        row[3] = redactJson(row[3], ["studentId", "name", "tel"]);
        row[16] = redactProjectVJson(row[16]);
        return row;
      });
      projectSheet.getRange(2, 6, lastRow - 1, 17).setValues(maskedValues);
    }
  }
}
function processSimpleMask(ss, sheetName, colIndices, replacement) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  colIndices.forEach(col => {
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const values = range.getValues().map(() => [replacement]);
    range.setValues(values);
  });
}
function redactJson(jsonStr, keysToMask) {
  if (!jsonStr) return jsonStr;
  try {
    const data = JSON.parse(jsonStr);
    keysToMask.forEach(key => { if (data[key]) data[key] = "***"; });
    return JSON.stringify(data);
  } catch (e) { return jsonStr; }
}
function redactProjectVJson(jsonStr) {
  if (!jsonStr) return jsonStr;
  try {
    const data = JSON.parse(jsonStr);
    if (data.contractInfo && data.contractInfo.signedBy) {
      data.contractInfo.signedBy = "***";
    }
    if (data.staff) {
      ["clean", "cook", "prep", "members"].forEach(group => {
        if (Array.isArray(data.staff[group])) {
          data.staff[group] = data.staff[group].map(person => {
            if (person.name) person.name = "***";
            if (person.id) person.id = "***";
            if (person.studentId) person.studentId = "***";
            return person;
          });
        }
      });
    }
    return JSON.stringify(data);
  } catch (e) { return jsonStr; }
}
function cleanupOperationalData() {
  const ss = getSpreadsheet();
  const now = new Date();
  const threeYearsAgo = new Date(now.getFullYear() - 3, now.getMonth(), now.getDate());
  const twoYearsAgo = new Date(now.getFullYear() - 2, now.getMonth(), now.getDate());
  const twoYearsAndOneDayAgo = new Date(now.getFullYear() - 2, now.getMonth(), now.getDate() - 1);
  const oneYearAgo = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate());
  // 1. 操作履歴のクリーンアップ（D列: 操作日時 が3年以上前なら行削除）
  const historySheet = ss.getSheetByName(SHEET_NAME.HISTORY);
  if (historySheet && historySheet.getLastRow() > 1) {
    const hData = historySheet.getRange(2, 4, historySheet.getLastRow() - 1, 1).getValues();
    for (let i = hData.length - 1; i >= 0; i--) {
      const rowDate = new Date(hData[i][0]);
      if (!isNaN(rowDate.getTime()) && rowDate < threeYearsAgo) {
        historySheet.deleteRow(i + 2); // 見出し行(+1)と配列インデックス(+1)の補正
      }
    }
  }
  // 2. 利用者管理のクリーンアップ
  const userSheet = ss.getSheetByName(SHEET_NAME.USER_MGMT);
  if (userSheet && userSheet.getLastRow() > 1) {
    const uData = userSheet.getDataRange().getValues();
    for (let i = uData.length - 1; i >= 1; i--) {
      const row = uData[i];
      const rowIndex = i + 1;
      const role = row[5];          // F列
      const lastAccessVal = row[8]; // I列
      const forcedExitVal = row[9]; // J列
      const lastAccess = lastAccessVal ? new Date(lastAccessVal) : null;
      const forcedExit = forcedExitVal ? new Date(forcedExitVal) : null;
      // --- 行の削除判定 ---
      let shouldDelete = false;
      if (forcedExit && !isNaN(forcedExit.getTime()) && forcedExit < oneYearAgo) {
        // 条件: 強制退会から1年以上経過
        shouldDelete = true;
      } else if (!forcedExitVal && lastAccess && !isNaN(lastAccess.getTime()) && lastAccess < twoYearsAndOneDayAgo) {
        // 条件: 強制退会が空欄 かつ 最終アクセスから2年1日以上経過
        shouldDelete = true;
      }
      if (shouldDelete) {
        userSheet.deleteRow(rowIndex);
        continue; // 削除した場合は下の処理を行わずに次の行へ
      }
      // --- ゲスト降格・強制退会日時の記録判定 ---
      if (lastAccess && !isNaN(lastAccess.getTime()) && lastAccess < twoYearsAgo && role !== "ゲスト") {
        userSheet.getRange(rowIndex, 6).setValue("ゲスト"); // F列をゲストに
        userSheet.getRange(rowIndex, 10).setValue(now);     // J列に現在時刻を記録
      }
    }
  }
}

// ==========================================
// 12. 予約データの取得・操作 (Batch & Lock)
// ==========================================
function getMyReservations() {
  const userEmail = Session.getActiveUser().getEmail();
  if (!userEmail) return [];
  const myOrgs = [];
  const projectSheet = getSheet(SHEET_NAME.PROJECTS);
  if (projectSheet) {
    const founds = projectSheet.getRange("F:G").createTextFinder(userEmail).matchEntireCell(true).findAll();
    founds.forEach(f => {
      const orgName = projectSheet.getRange(f.getRow(), 5).getValue(); // E列: 団体名
      if (orgName && !myOrgs.includes(orgName)) {
        myOrgs.push(orgName);
      }
    });
  }
  const data = getBaseReservationData();
  const idx = { id: 0, org: 1, cat: 2, qty: 3, start: 4, end: 5, status: 6, email: 8, comment: 11, reason: 12 };
  return data
    .filter(r => {
      const isMyRequest = (r[idx.email] === userEmail);
      const isMyOrgRequest = (r[idx.org] && myOrgs.includes(r[idx.org]));
      return isMyRequest || isMyOrgRequest;
    })
    .map(r => ({
      reservationId: r[idx.id],
      organizationName: r[idx.org],
      category: r[idx.cat],
      quantity: r[idx.qty],
      startTime: new Date(r[idx.start]).toISOString(),
      endTime: new Date(r[idx.end]).toISOString(),
      status: r[idx.status],
      adminComment: r[idx.comment],
      hasAppealed: (r[idx.reason] && r[idx.reason].toString().trim() !== "")
    }));
}
function getPendingReservations() {
  const data = getBaseReservationData(); // キャッシュから取得
  const now = new Date();
  const idx = { id: 0, org: 1, cat: 2, qty: 3, start: 4, end: 5, status: 6, email: 8, purpose: 10, reason: 12 };
  const list = [];
  const updates = []; 
  data.forEach((r, i) => {
    const status = r[idx.status];
    const endTime = new Date(r[idx.end]);
    const sheetRow = i + 2; 
    if (endTime < now && (status === "審査中" || status === "再審中")) {
      const newStatus = "却下";
      const comment = (status === "再審中") ? "予約期間超過による再審却下（自動処理）" : "予約期間超過による却下（自動処理）";
      updates.push({ row: sheetRow, status: newStatus, comment: comment, approver: "システム自動処理" });
    }
    else if (status === "審査中" || status === "再審中") {
      list.push({
        reservationId: r[idx.id],
        organizationName: r[idx.org],
        category: r[idx.cat],
        quantity: r[idx.qty],
        startTime: new Date(r[idx.start]).toISOString(),
        endTime: endTime.toISOString(),
        applicantEmail: r[idx.email],
        status: status,
        usagePurpose: r[idx.purpose] || "",
        appealReason: (status === "再審中") ? (r[idx.reason] || "") : ""
      });
    }
  });
  if (updates.length > 0) {
    withLock(() => {
      const sheet = getSheet(SHEET_NAME.RESERVATION);
      updates.forEach(u => {
        sheet.getRange(u.row, 7).setValue(u.status);   // G列
        sheet.getRange(u.row, 12).setValue(u.comment); // L列
        sheet.getRange(u.row, 15).setValue(u.approver); // O列 (システム自動処理)
      });
      clearLargeCache("RESERVATION_BASE_DATA");
    });
  }
  return list;
}
function processReservationDecision(data) {
  const userInfo = getUserInfo();
  const approverName = userInfo.name;
  const adminEmail = Session.getActiveUser().getEmail();
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.RESERVATION);
    const found = sheet.getRange("A:A").createTextFinder(data.reservationId).matchEntireCell(true).findNext();
    if (!found) throw new Error("予約IDが見つかりません。");
    const row = found.getRow();
    const currentStatus = sheet.getRange(row, 7).getValue(); // G列
    let newStatus = "";
    let finalComment = data.comment || "";
    if (data.action === "approve") {
      newStatus = "承認";
    } else if (data.action === "conditional") {
      newStatus = "条件付"; 
      finalComment += "\n\n承認内容に異議がある場合は、再度詳細な理由を添えて新規申請をするか、ご意見・ご要望などを通じて学祭実行委員会総務局までご相談ください。";
    } else if (data.action === "reject") {
      if (currentStatus === "再審中") {
        newStatus = "再審却下";
      } else {
        newStatus = "却下";
        finalComment += "\n\nなお、異議申立は一度のみ可能です。";
      }
    }
    const updateMain = [[data.quantity, new Date(data.startTime), new Date(data.endTime), newStatus]];
    sheet.getRange(row, 4, 1, 4).setValues(updateMain);
    sheet.getRange(row, 10).setValue(adminEmail);   // J列: 承認者メールアドレス
    sheet.getRange(row, 12).setValue(finalComment); // L列: 最終コメント（定型文付き）
    sheet.getRange(row, 15).setValue(approverName); // O列: 承認者氏名（★追加）
    clearLargeCache("RESERVATION_BASE_DATA"); // ★キャッシュクリア
    return `ステータスを「${newStatus}」に更新しました。`;
  });
}
function revertReservationStatus(id) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.RESERVATION);
    const found = sheet.getRange("A:A").createTextFinder(id).matchEntireCell(true).findNext();
    if (!found) throw new Error("IDが見つかりません");
    sheet.getRange(found.getRow(), 7).setValue("審査中");
    clearLargeCache("RESERVATION_BASE_DATA");
    return "ステータスを「審査中」に戻しました。";
  });
}
function cancelAppeal(reservationId) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.RESERVATION);
    const found = sheet.getRange("A:A").createTextFinder(reservationId).matchEntireCell(true).findNext();
    if (!found) throw new Error("予約IDが見つかりません。");
    const row = found.getRow();
    const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
    const status = rowData[6]; // G列
    const applicant = rowData[8]; // I列
    const user = Session.getActiveUser().getEmail();
    if (applicant !== user) throw new Error("権限がありません。");
    if (status !== "再審中") throw new Error("再審中のものしか取り下げられません。");
    sheet.getRange(row, 7).setValue("却下");
    sheet.getRange(row, 13).setValue(""); // 理由クリア
    clearLargeCache("RESERVATION_BASE_DATA");
    return "再審請求を取り下げました。";
  });
}

// ==========================================
// 13. マスタ・設定管理 (Batch & TextFinder)
// ==========================================
function getBaseReservationData() {
  const cacheKey = "RESERVATION_BASE_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.RESERVATION);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); 
  setLargeCache(cacheKey, JSON.stringify(data));
  return data;
}
function getReservationList() {
  const data = getBaseReservationData();
  const idx = { id: 0, org: 1, cat: 2, qty: 3, start: 4, end: 5, status: 6, name: 7, email: 8, admin: 9, adminName: 14 };
  return data
    .filter(r => r[idx.id] && r[idx.start] && r[idx.end]) 
    .map(r => ({
      reservationId: r[idx.id],
      organizationName: r[idx.org],
      category: r[idx.cat],
      quantity: r[idx.qty],
      startTime: new Date(r[idx.start]).toISOString(),
      endTime: new Date(r[idx.end]).toISOString(),
      applicantName: r[idx.name],
      applicantEmail: r[idx.email],
      adminName: r[idx.adminName],
      adminEmail: r[idx.admin],
      status: r[idx.status]
    }))
    .sort((a, b) => new Date(b.startTime) - new Date(a.startTime));
}
function addMasterListItem(listName, itemValue) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.CONSTANTS);
    let colIndex, startColChar, valToSearch, orgId;
    if (listName === 'category') { 
      colIndex = 1; startColChar = "A"; valToSearch = itemValue; 
    } else if (listName === 'organization') { 
      const name = data.name;
      const typeCode = data.typeCode;
      const prefix = `ORG-${typeCode}-`;
      const sheetData = sheet.getDataRange().getValues();
      let maxNum = 0;
      for (let i = 1; i < sheetData.length; i++) {
        const existingId = String(sheetData[i][5] || "").trim();
        if (existingId.startsWith(prefix)) {
          const numPart = parseInt(existingId.replace(prefix, ''), 10);
          if (!isNaN(numPart) && numPart > maxNum) {
            maxNum = numPart;
          }
        }
      }
      const nextNumStr = String(maxNum + 1).padStart(3, '0');
      const newId = `${prefix}${nextNumStr}`;
      sheet.appendRow(["", name, "", "", "", newId]);
      return `団体「${name}」を「${newId}」として追加しました。`;
    } else if (listName === 'location') { 
      colIndex = 3; startColChar = "C"; valToSearch = itemValue; 
    } else {
      throw new Error("無効なリスト名です: " + String(listName));
    }
    if (!valToSearch || String(valToSearch).trim() === "") {
      throw new Error("値が空です。");
    }
    const range = sheet.getRange(`${startColChar}:${startColChar}`);
    const found = range.createTextFinder(String(valToSearch)).matchEntireCell(true).findNext();
    if (found) throw new Error(`「${valToSearch}」は既に存在します。`);
    if (listName === 'organization') {
      const idFound = sheet.getRange("F:F").createTextFinder(String(orgId)).matchEntireCell(true).findNext();
      if (idFound) throw new Error(`団体ID「${orgId}」は既に存在します。`);
    }
    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, colIndex).setValue(valToSearch);
    if (listName === 'organization') {
      sheet.getRange(lastRow, 6).setValue(orgId); // F列に団体IDを記録
    }
    clearLargeCache("CONSTANTS_BASE_DATA");
    return `「${valToSearch}」を追加しました。`;
  });
}
function deleteMasterListItem(listName, itemValue) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.CONSTANTS);
    let startColChar;
    if (listName === 'category') startColChar = "A";
    else if (listName === 'organization') startColChar = "F"; 
    else if (listName === 'location') startColChar = "C";
    else throw new Error("無効なリスト名です: " + String(listName));
    const range = sheet.getRange(`${startColChar}:${startColChar}`);
    const found = range.createTextFinder(String(itemValue)).matchEntireCell(true).findNext();
    if (found) {
      const targetRow = found.getRow();
      if (listName === 'organization') {
        sheet.getRange(targetRow, 2).clearContent();
        sheet.getRange(targetRow, 6).clearContent();
      } else {
        found.clearContent();
      }
      clearLargeCache("CONSTANTS_BASE_DATA");
      return `削除しました。`;
    } else {
      throw new Error(`該当データが見つかりませんでした。`);
    }
  });
}
// システム設定
function parseDateSafe(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}
function formatDateForInput(date) {
  if (!date) return "";
  try {
    const d = new Date(date);
    if (isNaN(d.getTime())) return "";
    return Utilities.formatDate(d, "JST", "yyyy-MM-dd");
  } catch (e) { return ""; }
}
function getSystemSettings() {
  const sheet = getSheet(SHEET_NAME.SETTINGS);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = row[0];
    if (!key) continue;
    if (key === 'COMMITTEE_CHAIR' || key === 'GENERAL_AFFAIRS') {
      settings[key] = row[6]; 
    } 
    else if (key === 'FESTIVAL_PERIOD') {
      settings[key] = { 
        start: formatDateForInput(row[2]), 
        end: formatDateForInput(row[3]) 
      };
    } 
    else {
      settings[key] = {
        name: row[1],
        start: formatDateForInput(row[2]),
        end: formatDateForInput(row[3]),
        reStart: formatDateForInput(row[4]),
        reEnd: formatDateForInput(row[5])
      };
    }
  }
  return settings;
}
function saveSystemSettings(settings) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.SETTINGS);
    const data = sheet.getDataRange().getValues();
    const findRow = (key) => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) return i + 1;
      }
      return -1;
    };
    const setRowData = (rowNum, key, name, v1, v2, v3, v4, vG) => {
      if (rowNum > 0) {
        sheet.getRange(rowNum, 3, 1, 5).setValues([[v1||"", v2||"", v3||"", v4||"", vG||""]]);
      } else {
        sheet.appendRow([key, name, v1||"", v2||"", v3||"", v4||"", vG||""]);
      }
    };
    ['PHASE_1', 'PHASE_2', 'PHASE_3'].forEach(key => {
      if (settings[key]) {
        const s = settings[key];
        const r = findRow(key);
        setRowData(r, key, s.name||'期間設定', s.start, s.end, s.reStart, s.reEnd, "");
      }
    });
    if (settings.COMMITTEE_CHAIR) {
      const r = findRow("COMMITTEE_CHAIR");
      setRowData(r, "COMMITTEE_CHAIR", "実行委員長名", "", "", "", "", settings.COMMITTEE_CHAIR);
    }
    if (settings.GENERAL_AFFAIRS) {
      const r = findRow("GENERAL_AFFAIRS");
      setRowData(r, "GENERAL_AFFAIRS", "総務局長名", "", "", "", "", settings.GENERAL_AFFAIRS);
    }
    if (settings.FESTIVAL_PERIOD) {
      const r = findRow("FESTIVAL_PERIOD");
      setRowData(r, "FESTIVAL_PERIOD", "学祭本祭期間", settings.FESTIVAL_PERIOD.start, settings.FESTIVAL_PERIOD.end, "", "", "");
    }
    clearLargeCache("SYSTEM_SETTINGS_DATA");
    return "設定を保存しました。";
  });
}
// ==========================================
// 14. Phase 2/3 ロジック (Optimized)
// ==========================================
function uploadFile(data, fileName, mimeType) {
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      console.warn("共有設定の変更に失敗しました（アップロードは完了しているため続行します）: " + shareErr);
    } 
    return { fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
  } catch (e) {
    throw new Error("アップロード失敗: " + e.message);
  }
}
function savePhase2Data(data) {
  return withLock(() => {
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail); // ★修正
    if (!row) throw new Error("Phase 1 未完了");
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    if (data.proposalFileId) sheet.getRange(row, 16).setValue(data.proposalFileId);
    sheet.getRange(row, 17).setValue(JSON.stringify(data.relatedApps));
    sheet.getRange(row, 18).setValue(JSON.stringify(data.adInfo));
    sheet.getRange(row, 19).setValue(JSON.stringify(data.vehicleInfo));
    updatePhase2Status(row);
    clearLargeCache("ADMIN_PROJECT_LIST");
    return "書類・申請情報を保存しました。";
  });
}
function getEquipPageData() {
  const userId = getUserId();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // ID指定推奨
  const sheet = ss.getSheetByName('企画データ');
  const data = sheet.getDataRange().getValues();
  let planningType = 'INDOOR'; // デフォルトは屋内
  let prevData = {}; // 過去の申請データがあれば取得
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId) {
      const typeValue = data[i][2]; // ←ここを確認！(例: C列が区分の場合)
      if (typeValue && typeValue.includes('模擬店')) {
        planningType = 'STALL';
      }
      break;
    }
  }
  return {
    planningType: planningType, // 'STALL' or 'INDOOR'
    prevData: prevData
  };
}
function savePhase2Equipment(data) {
  return withLock(() => {
    const eqMap = new Map();
    (data.equipments || []).forEach(eq => {
      if (eq.count > 0) {
        if (eqMap.has(eq.name)) {
          const existing = eqMap.get(eq.name);
          existing.count += eq.count;
          if (eq.note) existing.note = existing.note ? existing.note + " / " + eq.note : eq.note;
        } else {
          eqMap.set(eq.name, { ...eq });
        }
      }
    });
    const exMap = new Map();
    (data.extras || []).forEach(ex => {
      if (ex.count > 0) {
        if (eqMap.has(ex.category)) {
          const existing = eqMap.get(ex.category);
          existing.count += ex.count;
          if (ex.name && ex.name !== ex.category) {
            existing.note = existing.note ? existing.note + " / " + ex.name : ex.name;
          }
        } else {
          if (exMap.has(ex.category)) {
            const existing = exMap.get(ex.category);
            existing.count += ex.count;
            if (ex.name && ex.name !== ex.category) {
              existing.name = existing.name === existing.category ? ex.name : existing.name + " / " + ex.name;
            }
          } else {
            exMap.set(ex.category, { ...ex });
          }
        }
      }
    });
    data.equipments = Array.from(eqMap.values());
    data.extras = Array.from(exMap.values());
  　const ss = getSpreadsheet();
    const projectSheet = ss.getSheetByName(SHEET_NAME.PROJECTS);
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail); 
    if (!row) throw new Error("Phase 1 が完了していません。");
    const pData = projectSheet.getRange(row, 1, 1, 5).getValues()[0];
    const projectId = pData[0];
    const projectType = pData[2];
    const orgName = pData[4];
    const userInfo = getUserInfo(); 
    data.status = data.isDraft ? "一時保存" : "提出済";
    projectSheet.getRange(row, 20).setValue(JSON.stringify(data)); // T列(20)にJSON保存
    const reserveSheet = ss.getSheetByName(SHEET_NAME.RESERVATION) || ss.insertSheet(SHEET_NAME.RESERVATION);
    const rData = reserveSheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = 1; i < rData.length; i++) {
      const rEmail = rData[i][8]; // I列: Email
      let rTag = "";
      try { rTag = rData[i][13] || ""; } catch (e) { }
      if (rEmail === userEmail && rTag && rTag.includes("FESTIVAL_PHASE2")) {
        rowsToDelete.push(i + 1);
      }
    }
    rowsToDelete.reverse().forEach(r => reserveSheet.deleteRow(r));
    if (!data.isDraft && data.useEquip) {
      const settings = getSystemSettings();
      const festivalConfig = settings['FESTIVAL_PERIOD'] || {}; 
      if (!festivalConfig.start) {
        throw new Error("システム設定エラー: 'FESTIVAL_PERIOD' (本祭期間) が設定されていません。");
      }
      const festivalStart = new Date(festivalConfig.start); 
      const usageStart = new Date(festivalStart);
      usageStart.setDate(festivalStart.getDate() - 1);
      usageStart.setHours(12, 0, 0, 0);
      const usageEnd = new Date(festivalStart);
      usageEnd.setDate(festivalStart.getDate() + 1);
      usageEnd.setHours(19, 0, 0, 0);
      const usageText = `一括申請（${projectType}）`;
      const systemTag = JSON.stringify({ type: "FESTIVAL_PHASE2", projectId: projectId });
      const newReserves = [];
      const addItem = (name, count, detail = "") => {
        newReserves.push([
          Utilities.getUuid(), orgName, name, count, usageStart, usageEnd,
          "審査中", userInfo.name, userEmail, "", usageText + detail, "", "", systemTag
        ]);
      };
      (data.equipments || []).forEach(eq => { if (eq.count > 0) addItem(eq.name, eq.count, eq.note ? ` (${eq.note})` : ""); });
      (data.extras || []).forEach(ex => { if (ex.name && ex.count > 0) addItem(ex.category, ex.count, ` (詳細: ${ex.name})`); });
      if (data.useRakuraku) addItem("簡易テント", 1, " (管理者決定)");
      if (data.useGenerator) addItem("発電機", 1, " (管理者決定)");
      if (data.useExtinguisher) addItem("消火器", 1, " (管理者決定)");
      if (newReserves.length > 0) {
        reserveSheet.getRange(reserveSheet.getLastRow() + 1, 1, newReserves.length, 14).setValues(newReserves);
      }
    }
    try { updateFireApplicationRequirement(projectId); } catch (e) { console.warn(e); }
    updatePhase2Status(row);
    clearLargeCache("ADMIN_PROJECT_LIST");
    clearLargeCache("RESERVATION_BASE_DATA");
    return data.isDraft ? "一時保存しました。" : "備品内容を保存しました。";
  });
}
function savePhase3Data(data) {
  return withLock(() => {
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail);
    if (!row) throw new Error("データが見つかりません。");
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    const currentJson = sheet.getRange(row, 22).getValue();
    const savedData = safeJsonParse(currentJson);
    const newData = { ...savedData, ...data.content };
    sheet.getRange(row, 22).setValue(JSON.stringify(newData));
    updatePhase3Status(row);
    clearLargeCache("ADMIN_PROJECT_LIST");
    return data.isDraft ? "一時保存しました。" : "保存しました。";
  });
}
function submitPhaseFinal(phaseKey) {
  return withLock(() => {
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail); // ★修正
    if (!row) throw new Error("データが見つかりません。");
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    if (phaseKey === 'PHASE_1') {
      sheet.getRange(row, 2).setValue("提出済");
      return "Phase 1: 企画出展手続き の申請を完了しました。";
    } else if (phaseKey === 'PHASE_2') {
      sheet.getRange(row, 15).setValue("提出済");
      return "Phase 2: 企画書提出手続き の申請を完了しました。";
    }
  });
}

// ==========================================
// 15. その他ヘルパー (Optimized)
// ==========================================
function getApprovedEquipment(projectId) {
  const sheet = getSheet(SHEET_NAME.RESERVATION);
  const data = sheet.getDataRange().getValues();
  const list = [];
  if (!projectId) {
    const userEmail = Session.getActiveUser().getEmail();
    const pSheet = getSheet(SHEET_NAME.PROJECTS);
    const found = pSheet.getRange("F:F").createTextFinder(userEmail).matchEntireCell(true).findNext();
    if (found) projectId = pSheet.getRange(found.getRow(), 1).getValue();
  }
  for (let i = 1; i < data.length; i++) {
    const status = data[i][6]; // G列: ステータス
    if (status !== "承認" && status !== "条件付") continue;
    let note = {};
    try {
      const tagCell = data[i][13];
      if (tagCell) note = JSON.parse(tagCell);
    } catch (e) { continue; }
    if (note.projectId === projectId && note.type === "FESTIVAL_PHASE2") {
      list.push({ name: data[i][2], count: data[i][3] });
    }
  }
  return list;
}
function updateFireApplicationRequirement(projectId) {
  const ss = getSpreadsheet();
  const projectSheet = ss.getSheetByName(SHEET_NAME.PROJECTS); // 企画データ
  const found = projectSheet.getRange("A:A").createTextFinder(projectId).matchEntireCell(true).findNext();
  if (!found) {
    console.error('Project ID not found: ' + projectId);
    return;
  }
  const row = found.getRow();
  const jsonString = projectSheet.getRange(row, 20).getValue();
  const jsonData = safeJsonParse(jsonString);
  let isFireRequired = false;
  if (jsonData.useGenerator === true) {
    isFireRequired = true;
  }
  setApplicationStatus(projectId, 'FIRE_APP', isFireRequired ? 'REQUIRED' : 'NONE');
}
function setApplicationStatus(projectId, appKey, status) {
  const ss = getSpreadsheet();
  const statusSheet = ss.getSheetByName('ステータス管理');
  if (!statusSheet) {
    console.error("「ステータス管理」シートが見つかりません");
    return;
  }
  if (appKey !== 'FIRE_APP') {
    console.warn("現在は FIRE_APP のみに対応しています");
    return;
  }
  const found = statusSheet.getRange("A:A").createTextFinder(projectId).matchEntireCell(true).findNext();
  if (found) {
    statusSheet.getRange(found.getRow(), 2).setValue(status);
  } else {
    statusSheet.appendRow([projectId, status]);
  }
}
function getCookingTimeSlots() {
  return [
    "前日 14:30-19:00",
    "1日目 08:00-10:00", "1日目 10:00-12:00", "1日目 12:00-15:00", "1日目 15:00-18:00",
    "2日目 08:00-10:00", "2日目 10:00-12:00", "2日目 12:00-15:00", "2日目 15:00-18:00"
  ];
}
function checkOrgNameSimilar(inputName) {
  const orgList = getOrganizationList();
  const normalizedInput = inputName.replace(/\s+/g, '').toLowerCase();
  return orgList.filter(org => {
    const normalizedOrg = org.replace(/\s+/g, '').toLowerCase();
    return normalizedOrg === normalizedInput || normalizedOrg.includes(normalizedInput) || normalizedInput.includes(normalizedOrg);
  });
}
function checkUserExists(email) {
  const sheet = getSheet(SHEET_NAME.USER_MGMT);
  const found = sheet.getRange("D:D").createTextFinder(email).matchEntireCell(true).findNext();
  if (found) {
    const row = sheet.getRange(found.getRow(), 1, 1, 8).getValues()[0];
    return { exists: true, name: row[1] + " " + row[2], department: row[6] + " / " + row[7] };
  }
  return { exists: false };
}
function getPhaseStatus(key, settings) { /* 元のコードと同じロジックのため省略可 */
  const now = new Date();
  const config = settings[key];
  if (!config) return { status: 'UNKNOWN', label: '不明' };
  const start = new Date(config.start);
  const end = new Date(config.end);
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);
  if (config.reStart && config.reEnd) {
    const reStart = new Date(config.reStart);
    const reEnd = new Date(config.reEnd);
    reStart.setHours(0, 0, 0, 0);
    reEnd.setHours(23, 59, 59, 999);
    if (now >= reStart && now <= reEnd) return { status: 'RESUBMIT', label: '再提出期間', isOpen: true };
  }
  if (now < start) return { status: 'BEFORE', label: '受付前', isOpen: false };
  else if (now > end) return { status: 'CLOSED', label: '受付終了', isOpen: false };
  else return { status: 'OPEN', label: '受付中', isOpen: true };
}
function getHomeData(isAdminView) {
  const userInfo = getUserInfo();
  let userId = null;
  let settings = {};
  let userProgress = {};
  try {
    const settingsCache = getLargeCache("SYSTEM_SETTINGS_DATA");
    if (settingsCache) {
      settings = JSON.parse(settingsCache);
    } else {
      settings = getSystemSettings();
      setLargeCache("SYSTEM_SETTINGS_DATA", JSON.stringify(settings));
    }
    userId = getUserId();
    userProgress = getUserProgress(userId); // 個人の進捗は常に最新を取得
    const phaseDef = [
      { key: 'PHASE_1', name: '企画出展手続き' },
      { key: 'PHASE_2', name: '企画書申請' },
      { key: 'PHASE_3', name: '最終手続' }
    ];
    const phaseList = phaseDef.map((p, index) => {
      const config = settings[p.key] || {}; 
      const sheetStatus = userProgress[p.key] || 'NONE';
      let statusInfo = determinePhaseStatus(config, sheetStatus, isAdminView);
      if (statusInfo.code !== 'DELETED' && !isAdminView && index > 0 && statusInfo.code !== 'APPROVED') {
        const prevKey = phaseDef[index - 1].key;
        const prevStatus = userProgress[prevKey] || 'NONE';
        if (prevStatus !== 'SUBMITTED' && prevStatus !== 'APPROVED') {
          statusInfo.code = 'LOCKED';
          statusInfo.label = '受付前';
          statusInfo.message = '前の手続きが完了していません';
          statusInfo.isEditable = false;
        }
      }
      return {
        key: p.key,
        name: config.name || p.name,
        code: statusInfo.code,
        label: statusInfo.label,
        dateStr: statusInfo.dateStr,
        message: statusInfo.message,
        isEditable: statusInfo.isEditable
      };
    });
    let files = [];
    const filesCache = getLargeCache("FILES_DATA");
    if (filesCache) {
      files = JSON.parse(filesCache);
    } else {
      const filesSheet = getSheet(SHEET_NAME.FILES);
      if (filesSheet && filesSheet.getLastRow() > 1) {
        const fileData = filesSheet.getRange(2, 1, filesSheet.getLastRow() - 1, 6).getValues();
        files = fileData.map(r => ({
          title: r[1], target: r[2], fileId: r[3],
          date: formatDateStr(r[4])
        }));
        setLargeCache("FILES_DATA", JSON.stringify(files));
      }
    }
    const filteredFiles = files.filter(f => {
      if (userInfo.role === 'オーナー' || userInfo.role === '管理者') return true;
      return f.target === "全体" || f.target === userInfo.role;
    });
    return {
      userName: userInfo.name,
      phaseList: phaseList,
      files: filteredFiles
    };
  } catch (e) {
    console.error("getHomeData Fatal Error: " + e.stack);
    return { 
      userName: userInfo.name || "Error", 
      phaseList: [], 
      files: [], 
      error: "データ読み込みエラー: " + e.message 
    };
  }
}
function determinePhaseStatus(config, sheetStatus, isAdminView) {
  if (sheetStatus === 'DELETED') {
    return { code: 'DELETED', label: '削除済', dateStr: '-', isEditable: false, message: 'この企画は運営により取り消されました。' };
  }
  if (sheetStatus === 'APPROVED') {
    return { code: 'APPROVED', label: '承認済', dateStr: isAdminView ? '期間無制限' : null, isEditable: false };
  }
  if (isAdminView) {
    return { code: 'OPEN', label: '内部者モード', dateStr: '期間無制限', isEditable: true };
  }
  if (!config || !config.start) {
    return { code: 'LOCKED', label: '期間設定なし', dateStr: '-', isEditable: false };
  }
  const now = new Date();
  const start = new Date(config.start);
  const end = new Date(config.end);
  end.setHours(23, 59, 59, 999);
  let reStart = null, reEnd = null;
  let inResubmitPeriod = false;
  if (config.reStart && config.reEnd) {
    reStart = new Date(config.reStart);
    reEnd = new Date(config.reEnd);
    reEnd.setHours(23, 59, 59, 999);
    if (now >= reStart && now <= reEnd) inResubmitPeriod = true;
  }
  let dateStr = `${formatDateStr(start)} 〜 ${formatDateStr(end)}`;
  if (inResubmitPeriod) dateStr = `${formatDateStr(reStart)} 〜 ${formatDateStr(reEnd)} (再提出期間)`;
  if (inResubmitPeriod) {
    if (sheetStatus === 'RESUBMIT' || sheetStatus === '再提出') {
      return { code: 'RESUBMIT', label: '差戻 (要修正)', dateStr, isEditable: true, message: '修正が必要です' };
    } else if (sheetStatus === 'SUBMITTED') {
      return { code: 'RESUBMITTED', label: '再提出済', dateStr, isEditable: true };
    }
    return { code: 'OPEN', label: '受付中 (再提出期間)', dateStr, isEditable: true };
  }
  if (now < start) {
    return { code: 'LOCKED', label: '受付前', dateStr, isEditable: false };
  }
  if (now > end) {
    return { code: 'CLOSED', label: '受付終了', dateStr, isEditable: false };
  }
  if (sheetStatus === 'SUBMITTED') {
    return { code: 'SUBMITTED', label: '提出済', dateStr, isEditable: true };
  } else if (sheetStatus === 'TEMP') {
    return { code: 'TEMP', label: '一時保存', dateStr, isEditable: true };
  } else {
    return { code: 'OPEN', label: '受付中', dateStr, isEditable: true };
  }
}
function formatDateStr(d) {
  if (!d) return "-";
  try {
    const dateObj = new Date(d);
    if (isNaN(dateObj.getTime())) return "-";
    return Utilities.formatDate(dateObj, "JST", "MM/dd");
  } catch (e) {
    return "-";
  }
}
function getFestivalEquipList() {
  const cacheKey = "FESTIVAL_EQUIP_LIST_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.FES_EQUIP);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // 備品名がある行のみ
      list.push({
        name: data[i][0],      // A列: 備品名
        targetType: data[i][1],// B列: 対象企画
        limit: data[i][2]      // C列: 上限数
      });
    }
  }
  setLargeCache(cacheKey, JSON.stringify(list));
  return list;
}
function getEquipCategories() { return ["テント", "机", "椅子", "電源", "調理器具", "清掃用具", "その他"]; }
function getRakurakuTentLimit(t) { return (t === "模擬店") ? 1 : 0; }

// ==========================================
// 16. リスト取得・ユーザー管理 (Optimized)
// ==========================================
function getAllItemsList() {
  const cacheKey = "ITEMS_LIST_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  try {
    const infoSheet = getSheet(SHEET_NAME.INFO);
    const data = infoSheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = {
      id: headers.indexOf("備品ID"),
      name: headers.indexOf("備品名"),
      category: headers.indexOf("カテゴリ"),
      status: headers.indexOf("状態"),
      location: headers.indexOf("現在地")
    };
    if (idx.id === -1 || idx.status === -1) throw new Error("ヘッダーエラー");

    const result = data
      .filter(row => row[idx.status] !== "")
      .map(row => ({
        id: row[idx.id],
        name: row[idx.name],
        category: row[idx.category],
        status: row[idx.status],
        location: row[idx.location]
      }));
    setLargeCache(cacheKey, JSON.stringify(result)); // キャッシュに保存
    return result;
  } catch (e) {
    console.error("getAllItemsListエラー: " + e);
    return [];
  }
}
function getOperationHistory() {
  const cacheKey = "OP_HISTORY_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  try {
    const historySheet = getSheet(SHEET_NAME.HISTORY);
    if (!historySheet) return [];
    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) { return [];}
    const result = data.map((row) => {
      const timeValue = row[3]; 
      return {
        logId: row[0] || "---", 
        id: row[1] || "不明ID",
        type: row[2] || "不明操作",
        timestamp: (timeValue instanceof Date) ? timeValue.toISOString() : String(timeValue || ""),
        user: row[6] || "システム"
      };
    });
    setLargeCache(cacheKey, JSON.stringify(result));
    return result;
  } catch (e) {return [];}
}
function getUserList() {
  const cacheKey = "USER_LIST_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  try {
    const userSheet = getSheet(SHEET_NAME.USER_MGMT);
    const data = userSheet.getDataRange().getValues();
    const headers = data.shift();
    const idx = {
      last: headers.indexOf("姓"),
      first: headers.indexOf("名"),
      email: headers.indexOf("メールアドレス"),
      role: headers.indexOf("ロール"),
      dept: headers.indexOf("所属区分"),
      deptName: headers.indexOf("所属名")
    };
    const result = data
      .filter(row => row[idx.email]) 
      .filter(row => row[idx.role] !== "ゲスト") // 修正：定義したidxを利用
      .map(row => ({
        lastName: row[idx.last],
        firstName: row[idx.first],
        email: row[idx.email],
        role: row[idx.role],
        department: row[idx.dept],
        deptName: (idx.deptName !== -1) ? row[idx.deptName] : ""
      }));
    setLargeCache(cacheKey, JSON.stringify(result));
    return result;
  } catch (e) {
    console.error("getUserListエラー: " + e);
    return [];
  }
}
function updateUserRole(email, newRole, password) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.USER_MGMT);
    const found = sheet.getRange("D:D").createTextFinder(email).matchEntireCell(true).findNext();
    if (!found) throw new Error("指定されたユーザーが見つかりません");
    const row = found.getRow();
    sheet.getRange(row, 6).setValue(newRole);
    clearLargeCache("USER_LIST_DATA");
    return `権限を変更しました。`;
  });
}
function checkAdminPassword(submittedPassword) {
  return submittedPassword === MASTER_PASSWORD;
}

// ==========================================
// 17. その他自動化 (Optimized)
// ==========================================
function updateFireApplicationStatus(userId) {
  const ss = getSpreadsheet();
  const planSheet = ss.getSheetByName(SHEET_NAME.PROJECTS);
  const found = planSheet.getRange("A:A").createTextFinder(userId).matchEntireCell(true).findNext();
  if (found) {
    const row = found.getRow();
    const useFire = planSheet.getRange(row, 8).getValue();
    const status = (useFire === 'true') ? 'REQUIRED' : 'NOT_REQUIRED';
    setApplicationStatus(userId, 'FIRE_APP', status);
  } else {
    console.warn(`updateFireApplicationStatus: User ${userId} not found.`);
  }
}
function calcPhaseStatus(config) {
  if (!config) return { code: 'UNKNOWN', label: '設定なし', isOpen: false };
  const now = new Date();
  if (!config.start) return { code: 'TBD', label: '未定', isOpen: false };
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const start = new Date(config.start.getFullYear(), config.start.getMonth(), config.start.getDate());
  const end = config.end ? new Date(config.end.getFullYear(), config.end.getMonth(), config.end.getDate()) : null;
  if (config.reStart && config.reEnd) {
    const reStart = new Date(config.reStart.getFullYear(), config.reStart.getMonth(), config.reStart.getDate());
    const reEnd = new Date(config.reEnd.getFullYear(), config.reEnd.getMonth(), config.reEnd.getDate());
    if (today >= reStart && today <= reEnd) {
      return { code: 'RESUBMIT', label: '再提出期間', isOpen: true, isResubmit: true };
    }
  }
  if (today < start) {
    return { code: 'BEFORE', label: '受付前', isOpen: false };
  }
  if (end && today > end) {
    return { code: 'CLOSED', label: '受付終了', isOpen: false };
  }
  return { code: 'OPEN', label: '受付中', isOpen: true };
}
function checkFireRequirement(userId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('企画データ');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userId) {
      const jsonString = data[i][9]; // J列(index 9)
      try {
        const json = JSON.parse(jsonString);
        return json.useGenerator === true || json.useGenerator === "true";
      } catch (e) {
        return false;
      }
    }
  }
  return false;
}
function getFestivalSettings() {
  const sheet = getSheet(SHEET_NAME.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const settings = {
    chairName: "（未設定）",
    affairsName: "（未設定）",
    startDate: new Date(),
    endDate: new Date()
  };
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    if (key === "COMMITTEE_CHAIR") settings.chairName = data[i][6]; // G列
    if (key === "GENERAL_AFFAIRS") settings.affairsName = data[i][6]; // G列
    if (key === "FESTIVAL_PERIOD") {
      settings.startDate = new Date(data[i][2]); // C列
      settings.endDate = new Date(data[i][3]);   // D列
    }
  }
  return settings;
}
function getUserProgress(userId) {
  if (!userId) return { PHASE_1: 'NONE', PHASE_2: 'NONE', PHASE_3: 'NONE', FIRE_APP: 'NONE' };
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  const row = findProjectRowByEmail(userId);
  if (!row) {
    return { PHASE_1: 'NONE', PHASE_2: 'NONE', PHASE_3: 'NONE', FIRE_APP: 'NONE' };
  }
  const rowData = sheet.getRange(row, 1, 1, 21).getValues()[0];
  const normalize = (status) => {
    if (!status) return 'NONE';
    const s = status.toString();
    if (s === '削除') return 'DELETED';
    if (s.includes('差戻') || s.includes('再提出')) return 'RESUBMIT';
    if (s.includes('承認') || s.includes('条件付')) return 'APPROVED';
    if (s.includes('一時保存')) return 'TEMP';
    if (s.includes('提出') || s.includes('申請中')) return 'SUBMITTED';
    return 'NONE';
  };
  const p1 = rowData[1];
  const p2 = rowData[14];
  const p3 = rowData[20];
  let fireStatus = 'NONE';
  const statusSheet = getSheet("ステータス管理");
  if (statusSheet && rowData[0]) {
    try {
      const fFound = statusSheet.getRange("A:A").createTextFinder(rowData[0].toString()).matchEntireCell(true).findNext();
      if (fFound) {
        fireStatus = (p2 && p2.includes('提出')) ? 'SUBMITTED' : 'NONE';
      }
    } catch (e) {}
  }
  return {
    PHASE_1: normalize(p1),
    PHASE_2: normalize(p2),
    PHASE_3: normalize(p3),
    FIRE_APP: fireStatus
  };
}
/**
 * Phase 1 のステータスを自動判定して更新 (B列)
 */
function updatePhase1Status(row) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  const data = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const basicJson = safeJsonParse(data[9]);  // J列 Basic
  const locJson = safeJsonParse(data[12]); // M列 Location
  const roomJson = safeJsonParse(data[13]); // N列 Room
  const isBasicDone = !!(basicJson.projectName && basicJson.memberCount);
  const isLocDone = !!(locJson.loc1); // 第1希望があればOKとみなす
  const isRoomDone = (roomJson.useRoom !== undefined); // 使用有無が選択されていればOK
  let status = "未着手";
  if (isBasicDone && isLocDone && isRoomDone) {
    status = "提出済";
  } else if (isBasicDone || isLocDone || isRoomDone) {
    status = "一時保存";
  }
  sheet.getRange(row, 2).setValue(status);
}
/**
 * Phase 2 のステータスを自動判定して更新 (O列)
 */
function updatePhase2Status(row) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  const data = sheet.getRange(row, 1, 1, 20).getValues()[0];
  const fileId = data[15]; // P列
  const equipJson = safeJsonParse(data[19]); // T列
  const isDocDone = !!fileId; // 企画書ファイルがあればOK
  const isEquipDone = (equipJson.status === "提出済");
  const isEquipTouched = (equipJson.status === "一時保存" || isEquipDone);
  let status = "未着手";
  if (isDocDone && isEquipDone) {
    status = "提出済";
  } else if (isDocDone || isEquipTouched) {
    status = "一時保存";
  }
  sheet.getRange(row, 15).setValue(status);
}
/**
 * Phase 3 のステータスを自動判定して更新 (U列)
 */
function updatePhase3Status(row) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  const data = sheet.getRange(row, 1, 1, 22).getValues()[0];
  const projectType = data[2]; // C列
  const p2Equip = safeJsonParse(data[19]); // T列
  const content = safeJsonParse(data[21]); // V列
  const isPledgeDone = !!content.pledgeFileId;
  const isStaffDone = !!(content.staff && content.staff.prep && content.staff.prep.length > 0);
  const isEquipRequired = (p2Equip && p2Equip.useEquip);
  const isContractDone = isEquipRequired ? !!content.contractFileId : true;
  let isCookSurveyDone = true;
  let isCookDocDone = true;
  if (projectType === "模擬店") {
    isCookSurveyDone = !!(content.cookingSurvey);
    if (content.cookingSurvey && content.cookingSurvey.use) {
      isCookDocDone = !!content.cookingDocFileId;
    }
  }
  let status = "一時保存"; // デフォルト
  if (isPledgeDone && isStaffDone && isContractDone && isCookSurveyDone && isCookDocDone) {
    status = "提出済";
  }
  sheet.getRange(row, 21).setValue(status);
}
// 3. 備品借用書 (Excel/Spreadsheet) 生成機能
function createLoanForm(projectId) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  let targetId = projectId;
  if (!targetId) {
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail);
    if (row) targetId = sheet.getRange(row, 1).getValue();
  }
  if (!targetId) throw new Error("企画IDを特定できませんでした。");
  const found = sheet.getRange("A:A").createTextFinder(targetId).matchEntireCell(true).findNext();
  if (!found) throw new Error("企画データが見つかりません");
  const rowData = sheet.getRange(found.getRow(), 1, 1, 24).getValues()[0];
  const orgName = rowData[4]; 
  const repJson = safeJsonParse(rowData[7]);
  const repName = repJson.name || "";
  const repId = repJson.studentId || "";
  const currentEquips = getApprovedEquipment(targetId);
  const content = safeJsonParse(rowData[21]); // V列
  const contract = content.contractInfo;
  let items = currentEquips;
  let isSigned = false;
  let signName = repName;
  let signText = "（未同意：内容確認用）";
  if (contract) {
    const currentStr = JSON.stringify(currentEquips.sort((a,b) => a.name > b.name ? 1 : -1));
    const snapshotStr = contract.snapshot ? JSON.stringify(contract.snapshot.sort((a,b) => a.name > b.name ? 1 : -1)) : currentStr;
    if (currentStr === snapshotStr) {
      items = contract.snapshot || currentEquips;
      isSigned = true;
      if (contract.isProxy) {
        signName = contract.signatureName;
        signText = `電子同意済 (代筆：${contract.proxyName} / 事由：${contract.proxyReason || "記載なし"})`;
      } else {
        signName = contract.signatureName || repName;
        signText = `電子同意済`;
      }
    }
  }
  if (items.length === 0) throw new Error("承認された備品がありません。");
  const settings = getFestivalSettings();
  const currentYear = new Date().getFullYear();
  const loanStart = new Date(settings.startDate);
  loanStart.setDate(loanStart.getDate() - 1);
  const loanStartStr = Utilities.formatDate(loanStart, "JST", "MM月dd日 (E) 12:00");
  const loanEndStr = Utilities.formatDate(settings.endDate, "JST", "MM月dd日 (E) 18:00");
  const loanPeriodStr = `${loanStartStr} ～ ${loanEndStr}`;
  const tempSs = SpreadsheetApp.create(`Temp_Loan_${targetId}`);
  const s = tempSs.getActiveSheet();
  s.setHiddenGridlines(true); 
  const fontName = "游明朝"; 
  s.getDataRange().setFontFamily(fontName);
  s.setColumnWidth(1, 15);  
  s.setColumnWidth(2, 20);  
  s.setColumnWidth(3, 100); 
  s.setColumnWidth(4, 340); 
  s.setColumnWidth(5, 70);  
  s.setColumnWidth(6, 40);  
  const COLOR_BG_HEADER = "#E2EFDA"; 
  const COLOR_BORDER = "#385723";    
  s.getRange("B2:E2").merge()
   .setValue("＜ 備品借用書（控え）＞")
   .setHorizontalAlignment("center").setFontSize(20).setFontWeight("bold");
  s.setRowHeight(2, 45);
  s.getRange("C4").setValue("山梨大学医学部学祭実行委員会");
  s.getRange("C5").setValue(`委員長　　${settings.chairName}　殿`);
  s.getRange("C6").setValue(`総務局長　${settings.affairsName}　殿`);
  s.setRowHeights(4, 3, 22);
  const today = new Date();
  s.getRange("C8:F8").merge()
   .setValue(`令和 ${today.getFullYear()-2018} 年 ${today.getMonth()+1} 月 ${today.getDate()} 日`)
   .setHorizontalAlignment("right");
  s.setRowHeight(8, 25);
  s.getRange("C10").setValue("団体名");
  s.getRange("D10:F10").merge().setValue("　" + orgName);
  s.getRange("C11").setValue(isSigned ? "同意者氏名" : "代表者氏名");
  s.getRange("D11:F11").merge().setValue(`　${signName}　　${signText}`).setFontSize(11);
  s.getRange("C12").setValue("学籍番号");
  s.getRange("D12:F12").merge().setValue("　" + repId).setFontSize(12);
  if (isSigned && contract.signedAt) {
    const date = new Date(contract.signedAt);
    const timeStr = Utilities.formatDate(date, "JST", "yyyy年MM月dd日 HH:mm:ss");
    s.getRange("C13").setValue("同意日時");
    s.getRange("D13:F13").merge().setValue("　" + timeStr).setFontSize(11);
  }
  s.setRowHeights(10, 4, 26);
  s.getRange("B15:F15").merge()
   .setValue(`梨医祭 ${currentYear} において、以下の物品を以下の条件により借用したく申請いたします。`)
   .setVerticalAlignment("bottom");
  s.setRowHeight(15, 40);
  const startRow = 17;
  const setBorder = (r) => r.setBorder(true, true, true, true, null, null, COLOR_BORDER, SpreadsheetApp.BorderStyle.THIN);
  const h1 = s.getRange(startRow, 2, 1, 3).merge().setValue("物品名");
  const h2 = s.getRange(startRow, 5, 1, 2).merge().setValue("個数");
  [h1, h2].forEach(r => { 
    setBorder(r); 
    r.setBackground(COLOR_BG_HEADER).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
  });
  s.setRowHeight(startRow, 25);
  let cur = startRow + 1;
  items.forEach(item => {
    const cleanName = item.name.replace(/[□■]/g, '').trim(); 
    const r1 = s.getRange(cur, 2, 1, 3).merge().setValue("　" + cleanName);
    const r2 = s.getRange(cur, 5, 1, 2).merge().setValue(item.count).setHorizontalAlignment("center");
    setBorder(r1); setBorder(r2);
    s.setRowHeight(cur, 30);
    cur++;
  });
  for(let i=0; i<5; i++) {
    setBorder(s.getRange(cur, 2, 1, 3).merge());
    setBorder(s.getRange(cur, 5, 1, 2).merge());
    s.setRowHeight(cur, 30);
    cur++;
  }
  const nRow = cur + 1;
  const noteFontSize = 9;
  s.getRange(nRow, 2).setValue("1.");
  s.getRange(nRow, 3, 1, 3).merge().setValue("借用期間");
  s.getRange(nRow+1, 3, 1, 3).merge().setValue("　" + loanPeriodStr).setFontWeight("bold");
  s.getRange(nRow+3, 2).setValue("2.");
  s.getRange(nRow+3, 3, 1, 3).merge().setValue("使用・管理について");
  const longTextConfig = (range, text) => {
    range.merge().setValue(text)
         .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
         .setVerticalAlignment("top")
         .setFontSize(noteFontSize);
  };
  s.getRange(nRow+4, 2).setValue("（１）").setFontSize(noteFontSize).setVerticalAlignment("top");
  longTextConfig(s.getRange(nRow+4, 3, 1, 4), "使用した物品は、注意をもって使用および管理すること。万一、破損させてしまった場合には、損害賠償の責任を負うものとする。");
  s.setRowHeight(nRow+4, 28);
  s.getRange(nRow+5, 2).setValue("（２）").setFontSize(noteFontSize).setVerticalAlignment("top");
  longTextConfig(s.getRange(nRow+5, 3, 1, 4), "紛失・盗難等によって返却できない事態が生じた場合、団体が責任をもって、借用した物品と同様の代替品を弁償または相当額を賠償すること。");
  s.setRowHeight(nRow+5, 28);
  s.getRange(nRow+6, 2).setValue("（３）").setFontSize(noteFontSize).setVerticalAlignment("top");
  longTextConfig(s.getRange(nRow+6, 3, 1, 4), "借用した物品は目的に沿って使用し、転貸等は行わないこと。");
  s.setRowHeight(nRow+6, 20);
  s.getRange(nRow+8, 3, 1, 4).merge()
   .setValue("作成: 山梨大学医学部学祭実行委員会 総務局")
   .setHorizontalAlignment("right").setFontSize(10);
  SpreadsheetApp.flush();
  const ssid = tempSs.getId();
  const url = `https://docs.google.com/spreadsheets/d/${ssid}/export?format=pdf`
            + `&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&gridlines=false&fzr=false`
            + `&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5`;
  const token = ScriptApp.getOAuthToken();
  const options = { headers: { 'Authorization': 'Bearer ' + token } };
  const response = UrlFetchApp.fetch(url, options);
  const pdfBlob = response.getBlob().setName(`備品借用書_${orgName}.pdf`);
  const base64 = Utilities.base64Encode(pdfBlob.getBytes());
  DriveApp.getFileById(ssid).setTrashed(true);
  return { 
    base64: base64, 
    filename: `備品借用書${currentYear}_${orgName}.pdf`,
    contentType: 'application/pdf'
  };
}
function submitPhase3Contract(mainName, proxyName, role, proxyReason) {
  return withLock(() => {
    if (!mainName || mainName === "未登録") throw new Error("正責任者名が無効です。");
    if (role === 'SUB' && (!proxyName || proxyName === "未登録")) throw new Error("代筆者名が無効です。");
    if (role === 'SUB' && !proxyReason) throw new Error("代筆理由を選択してください。");
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail);
    if (!row) throw new Error("企画が見つかりません");
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    const data = sheet.getRange(row, 1, 1, 24).getValues()[0];
    const projectId = data[0];
    const currentEquipList = getApprovedEquipment(projectId);
    const contractInfo = {
      signedAt: new Date().getTime(),
      signatureName: mainName,
      proxyName: proxyName,
      proxyReason: proxyReason, // ★理由を保存
      signerEmail: userEmail,
      role: role,
      isProxy: (role === 'SUB'),
      snapshot: currentEquipList
    };
    let content = safeJsonParse(data[21]);
    content.contractInfo = contractInfo;
    content.contractFileId = "SIGNED_ELECTRONICALLY";
    sheet.getRange(row, 22).setValue(JSON.stringify(content));
    clearLargeCache("ADMIN_PROJECT_LIST"); 
    return role === 'SUB' ? `代筆（理由：${proxyReason}）としてサインを完了しました。` : "サインを完了しました。";
  });
}
// ==========================================
// 18. マイページ・チーム管理機能
// ==========================================
function updateProjectTeam(action, data) {
  return withLock(() => {
    const userEmail = Session.getActiveUser().getEmail();
    const row = findProjectRowByEmail(userEmail);
    if (!row) throw new Error("企画データが見つかりません。");
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    const rowData = sheet.getRange(row, 1, 1, 10).getValues()[0]; // A列(1)からJ列(10)まで取得
    // --- 1. 変更前の情報を保持 ---
    const orgName = rowData[4];    // E列
    const projectName = rowData[3]; // D列
    const oldMainEmail = rowData[5]; // F列
    const oldSubEmail = rowData[6];  // G列
    const oldMainInfo = safeJsonParse(rowData[7]); // H列
    const oldSubInfo = safeJsonParse(rowData[8]);  // I列
    let resultMessage = "";
    let newMainEmail = oldMainEmail;
    let newSubEmail = oldSubEmail;
    let newMainInfo = oldMainInfo;
    let newSubInfo = oldSubInfo;
    // --- 2. 変更処理の実行 ---
    if (action === 'change_sub') {
      const newEmail = data.email;
      const newInfo = data.info;
      if (!newEmail || !newEmail.includes('@')) throw new Error("有効なメールアドレスを入力してください。");
      if (newEmail.trim().toLowerCase() === userEmail.toLowerCase()) {
        throw new Error("正責任者と副責任者に同じメールアドレスを設定することはできません。");
      }
      const allData = sheet.getDataRange().getValues();
      const currentProjectId = rowData[0];
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(currentProjectId)) continue;
        if (allData[i][5] === newEmail || allData[i][6] === newEmail) {
          throw new Error(`指定された副責任者 (${newEmail}) は、既に他の企画に登録されています。`);
        }
      }
      const userCheck = checkUserExists(newEmail);
      if (!userCheck.exists) throw new Error("指定されたメールアドレスはユーザー登録されていません。");
      newSubEmail = newEmail;
      newSubInfo = {
        name: newInfo.name || userCheck.name,
        studentId: newInfo.studentId || "",
        dept: newInfo.dept || "",
        grade: newInfo.grade || "",
        tel: newInfo.tel || ""
      };
      sheet.getRange(row, 7).setValue(newSubEmail); // G列
      sheet.getRange(row, 9).setValue(JSON.stringify(newSubInfo)); // I列
      resultMessage = `副責任者を ${newSubInfo.name} さんに変更しました。`;
    } else if (action === 'swap') {
      if (!oldSubEmail) throw new Error("副責任者が設定されていないため交代できません。");
      newMainEmail = oldSubEmail;
      newSubEmail = oldMainEmail;
      newMainInfo = oldSubInfo;
      newSubInfo = oldMainInfo;
      sheet.getRange(row, 6, 1, 2).setValues([[newMainEmail, newSubEmail]]); // F, G列
      sheet.getRange(row, 8, 1, 2).setValues([[JSON.stringify(newMainInfo), JSON.stringify(newSubInfo)]]); // H, I列
      resultMessage = "正責任者と副責任者を交代しました。ページを再読み込みします。";
    } else {
      throw new Error("不明な操作です。");
    }
    // --- 3. 管理者への自動問い合わせ通知処理 ---
    try {
      const inqSheet = getSheet(SHEET_NAME.INQUIRIES);
      if (inqSheet) {
        const userInfo = getUserInfo(); // 操作者の情報取得
        const timestamp = new Date();
        const content = `【システム自動通知：責任者変更】
        団体名：${orgName}
        企画名：${projectName}

        ■変更前
        正責任者：${oldMainInfo.name || "未設定"} (${oldMainEmail})
        副責任者：${oldSubInfo.name || "未設定"} (${oldSubEmail || "未設定"})

        ■変更後
        正責任者：${newMainInfo.name || "未設定"} (${newMainEmail})
        副責任者：${newSubInfo.name || "未設定"} (${newSubEmail})`;
        inqSheet.appendRow([
          Utilities.getUuid(),
          userEmail,          // 送信者Email (操作した人)
          userInfo.name,      // 送信者氏名
          userInfo.deptName || "所属不明", // 送信者所属
          "通常",             // 種別
          "責任者変更",       // カテゴリ
          content,            // 内容
          "未読",             // ステータス
          timestamp,          // 日時
          ""                  // 返信欄
        ]);
      }
    } catch (logError) {
      console.error("Auto-Inquiry log failed: " + logError);
    }
    return resultMessage;
  });
}
function findProjectRowByEmail(userEmail) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  if (!sheet) return null;
  let found = sheet.getRange("F:F").createTextFinder(userEmail).matchEntireCell(true).findNext();
  if (found) return found.getRow();
  found = sheet.getRange("G:G").createTextFinder(userEmail).matchEntireCell(true).findNext();
  if (found) return found.getRow();
  return null;
}
// ==========================================
// 19. 管理画面：企画管理・設定 (Admin Extensions)
// ==========================================
function getAdminProjectList() {
  const cacheKey = "ADMIN_PROJECT_LIST";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift(); // ヘッダー削除
  const result = data.map(row => {
    let ts = row[10];
    let tsStr = "";
    try {
      if (ts) tsStr = new Date(ts).toISOString();
    } catch(e) {
      tsStr = ""; 
    }
    return {
      id: row[0],
      p1_status: row[1],
      type: row[2],
      name: row[3],
      org: row[4],
      rep_email: row[5],
      sub_email: row[6],
      timestamp: tsStr,
      p2_status: row[14],
      proposal_file: row[15],
      p3_status: row[20],
      details: {
        mainRep: safeJsonParse(row[7]),
        subRep: safeJsonParse(row[8]),
        basicInfo: safeJsonParse(row[9]),
        location: safeJsonParse(row[12]),
        room: safeJsonParse(row[13]),
        relatedApps: safeJsonParse(row[16]),
        adInfo: safeJsonParse(row[17]),
        vehicleInfo: safeJsonParse(row[18]),
        p2_equip: safeJsonParse(row[19]),    
        p3_content: safeJsonParse(row[21])   
      }
    };
  }).reverse();
  setLargeCache(cacheKey, JSON.stringify(result));
  return result;
}
function updateProjectStatus(id, phase, statusKey, reason, fileObj) {
  return withLock(() => {
    let finalReason = reason || "";
    if (fileObj && fileObj.data) {
      try {
        const folder = DriveApp.getFolderById(DIST_FOLDER_ID); 
        const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, fileObj.name);
        const newFile = folder.createFile(blob);
        try {
          newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch(e) {
          console.warn("ファイル共有設定の失敗: " + e.message);
        }
        const fileUrl = newFile.getUrl();
        finalReason += `\n\n【添付資料】\n${fileUrl}`;
      } catch (e) {
        throw new Error("ファイルのアップロードに失敗しました: " + e.message);
      }
    }
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    const found = sheet.getRange("A:A").createTextFinder(id).matchEntireCell(true).findNext();
    if (!found) throw new Error("企画が見つかりません");
    const row = found.getRow();
    let statusCol = 0;
    let commentCol = 0;
    if (phase === 'PHASE_1') {
      statusCol = 2;  // B列 (ステータス)
      commentCol = 12; // L列 (コメント)
    } else if (phase === 'PHASE_2') {
      statusCol = 15; // O列 (ステータス)
      commentCol = 23; // W列 (コメント)
    } else if (phase === 'PHASE_3') {
      statusCol = 21; // U列 (ステータス)
      commentCol = 24; // X列 (コメント)
    }
    let statusLabel = statusKey;
    if (statusKey === 'APPROVED') statusLabel = '承認';
    else if (statusKey === 'RESUBMIT') statusLabel = '差戻';
    else if (statusKey === 'SUBMITTED') statusLabel = '提出済';
    const previousStatus = sheet.getRange(row, statusCol).getValue();
    sheet.getRange(row, statusCol).setValue(statusLabel);
    if (statusKey === 'RESUBMIT' || finalReason !== "") {
      sheet.getRange(row, commentCol).setValue(finalReason);
    }
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    const historySheet = getSheet(SHEET_NAME.HISTORY);
    historySheet.appendRow([
      generateHistoryId(), // A
      id,                       // B
      "企画審査",               // C
      timestampString,          // D
      `${phase}: ${previousStatus}`, // E: 変更前のフェーズとステータス
      `${phase}: ${statusLabel}`,    // F: 変更後のフェーズとステータス
      Session.getActiveUser().getEmail(), // G
      finalReason               // H: 理由を特記事項に
    ]);
    clearLargeCache("ADMIN_PROJECT_LIST");
    return `ステータスを「${statusLabel}」に更新しました。`;
  });
}
function manageFiles(action, data) {
  if (action === 'get') {
    const cachedData = getLargeCache("FILES_DATA");
    if (cachedData) {
      return JSON.parse(cachedData).reverse();
    }
    const sheet = getSheet(SHEET_NAME.FILES);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    const files = rows.map(r => {
      let dateStr = "";
      try { dateStr = new Date(r[4]).toISOString(); } catch(e) {}
      return {
        id: r[0], title: r[1], target: r[2], fileId: r[3], date: dateStr
      };
    });
    setLargeCache("FILES_DATA", JSON.stringify(files));
    return files.reverse();
  }
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.FILES); // 追加/削除用に取得
    if (action === 'add') {
      let fileId = "";
      if (data.fileData && data.fileName) {
        try {
          const folder = DriveApp.getFolderById(DIST_FOLDER_ID);
          const blob = Utilities.newBlob(Utilities.base64Decode(data.fileData), data.mimeType, data.fileName);
          const file = folder.createFile(blob);
          try {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch(e) {
            console.warn("ファイルの共有設定に失敗しました（続行します）");
          }
          fileId = file.getId();
        } catch (e) {
          throw new Error("アップロード失敗: " + e.message);
        }
      }
      const pubDate = data.date ? new Date(data.date) : new Date();
      const newId = new Date().getTime().toString();
      sheet.appendRow([newId, data.title, data.target, fileId, pubDate, ""]);
      clearLargeCache("FILES_DATA");
      return "資料を追加しました。";
    }
    if (action === 'delete') {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(data.id)) {
          sheet.deleteRow(i + 1);
          clearLargeCache("FILES_DATA");
          return "資料を削除しました。";
        }
      }
      throw new Error("削除対象が見つかりません。");
    }
  });
}
// ==========================================
// 21. 帳票出力機能 (PDF/Excel 直接ダウンロード版)
// ==========================================
function exportData(year, docType, format) {
  return withLock(() => {
    const list = getAdminProjectList();
    const targets = list.filter(p => {
      const d = new Date(p.timestamp);
      const pYear = isNaN(d.getTime()) ? new Date().getFullYear() : d.getFullYear();
      return String(pYear) === String(year) && (p.p1_status.includes('APPROVED') || p.p1_status.includes('承認'));
    });
    if (targets.length === 0) throw new Error("対象年度の承認済みデータがありません。");
    const ssName = `FesTrack_${docType === 'staff' ? '運営スタッフ名簿' : '備品貸出表'}_${year}年度`;
    const ss = SpreadsheetApp.create(ssName);
    const defaultSheet = ss.getSheets()[0];
    if (docType === 'staff') {
      createStaffSheets(ss, targets, year);
    } else {
      createEquipSheet(ss, targets, year);
    }
    ss.deleteSheet(defaultSheet);
    SpreadsheetApp.flush();
    const ssId = ss.getId();
    let blob;
    const token = ScriptApp.getOAuthToken();
    if (format === 'pdf') {
      const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&fzr=true&ir=false&ic=false`;
      const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
      blob = response.getBlob().setName(`${ssName}.pdf`);
    } else {
      const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx`;
      const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
      blob = response.getBlob().setName(`${ssName}.xlsx`);
    }
    DriveApp.getFileById(ssId).setTrashed(true);
    return {
      filename: blob.getName(),
      mimeType: blob.getContentType(),
      data: Utilities.base64Encode(blob.getBytes())
    };
  });
}
function createStaffSheets(ss, targets, year) {
  const categories = [
    { key: 'prep', name: '前日準備' },
    { key: 'clean', name: '最終片付け' },
    { key: 'cook', name: '調理室清掃' },
  ];
  categories.forEach(cat => {
    const sheet = ss.insertSheet(cat.name);
    const header = ["No", "企画名", "団体名", "氏名 (1)", "学籍番号 (1)", "氏名 (2)", "学籍番号 (2)", "備考"];
    const rows = [];
    let no = 1;
    targets.forEach(p => {
      const st = p.details.p3_content?.staff || {};
      const list = st[cat.key] || [];
      if (list.length > 0) {
        for (let i = 0; i < list.length; i += 2) {
          const m1 = list[i];
          const m2 = list[i+1] || { name: "", studentId: "" };
          rows.push([no++, p.name, p.org, m1.name, m1.id || m1.studentId || "", m2.name, m2.id || m2.studentId || "", ""]);
        }
      }
    });
    if (rows.length > 0) {
      writeSheetWithTitle(sheet, `【${year}年度】運営スタッフ名簿（${cat.name}）`, header, rows);
      sheet.setColumnWidth(1, 40);  sheet.setColumnWidth(2, 180); sheet.setColumnWidth(3, 180);
      sheet.setColumnWidth(4, 110); sheet.setColumnWidth(5, 90);  sheet.setColumnWidth(6, 110);
      sheet.setColumnWidth(7, 90);  sheet.setColumnWidth(8, 200);
    } else {
      sheet.getRange(1, 1).setValue("該当データなし");
    }
  });
}
function createEquipSheet(ss, targets, year) {
  const sheet = ss.insertSheet("備品貸出一覧");
  const fixedItems = ["鉄骨テント", "簡易テント", "パイプ椅子", "長机", "発電機", "消火器", "大机", "小机", "椅子", "ホワイトボード"];
  const equipMap = new Map();
  const dataMap = targets.map(p => {
    const eq = p.details.p2_equip || {};
    const rowObj = { name: p.name, org: p.org };
    fixedItems.forEach(k => rowObj[k] = 0);
    if (eq.useRakuraku) rowObj["簡易テント"] = 1;
    if (eq.useGenerator) rowObj["発電機"] = 1; 
    if (eq.useExtinguisher) rowObj["消火器"] = 1;
    (eq.equipments || []).forEach(e => {
      if (fixedItems.includes(e.name)) {
        rowObj[e.name] = (rowObj[e.name] || 0) + Number(e.count);
      } else {
        equipMap.set(e.name, true);
        rowObj[e.name] = Number(e.count);
      }
    });
    return rowObj;
  });
  const dynamicItems = Array.from(equipMap.keys());
  const allItems = [...fixedItems, ...dynamicItems];
  const header = ["No", "企画名", "団体名", ...allItems];
  const rows = [];
  const totals = new Array(allItems.length).fill(0);
  dataMap.forEach((d, idx) => {
    const row = [idx + 1, d.name, d.org];
    allItems.forEach((key, i) => {
      const val = d[key] || 0;
      row.push(val);
      totals[i] += val;
    });
    rows.push(row);
  });
  rows.push(["", "総計", "-", ...totals]);
  writeSheetWithTitle(sheet, `【${year}年度】備品貸出一覧表`, header, rows);
  const lastRow = 4 + rows.length;
  sheet.getRange(lastRow, 1, 1, header.length).setBackground("#E2EFDA").setFontWeight("bold");
  sheet.setColumnWidth(1, 40); sheet.setColumnWidth(2, 200); sheet.setColumnWidth(3, 200);
}
function writeSheetWithTitle(sheet, title, header, rows) {
  const lastCol = header.length;
  const dataStartRow = 4;
  const lastRow = dataStartRow + rows.length;
  sheet.getRange(1, 1, 1, lastCol).merge()
       .setValue(title)
       .setFontSize(18)
       .setFontWeight("bold")
       .setFontFamily("Noto Sans JP")
       .setVerticalAlignment("middle")
       .setFontColor("#203764");
  sheet.setRowHeight(1, 40);
  const now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm");
  sheet.getRange(2, 1, 1, lastCol).merge()
       .setValue(`出力日時: ${now}`)
       .setFontSize(9)
       .setFontFamily("Noto Sans JP")
       .setFontColor("#666666")
       .setHorizontalAlignment("right");
  sheet.getRange(dataStartRow, 1, 1, lastCol).setValues([header]);
  sheet.getRange(dataStartRow + 1, 1, rows.length, lastCol).setValues(rows);
  const tableRange = sheet.getRange(dataStartRow, 1, rows.length + 1, lastCol);
  tableRange.setFontFamily("Noto Sans JP")
            .setVerticalAlignment("middle")
            .setFontSize(10)
            .setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);
  const headerRange = sheet.getRange(dataStartRow, 1, 1, lastCol);
  headerRange.setBackground("#4472C4")
             .setFontColor("white")
             .setFontWeight("bold")
             .setHorizontalAlignment("center");
  sheet.setRowHeight(dataStartRow, 25);
  for (let i = dataStartRow + 1; i <= lastRow; i++) {
    if ((i - dataStartRow) % 2 === 0) {
      sheet.getRange(i, 1, 1, lastCol).setBackground("#F2F2F2");
    }
  }
}
// ==========================================
// 21. 問い合わせ・通知機能
// ==========================================
function submitInquiry(data) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.INQUIRIES); // "問い合わせ"
    if (!sheet) throw new Error("問い合わせシートが見つかりません");
    const id = Utilities.getUuid();
    const userEmail = Session.getActiveUser().getEmail();
    const timestamp = new Date();
    let senderName = data.name;
    let senderDept = data.dept;
    try {
      const userSheet = getSheet(SHEET_NAME.USER_MGMT);
      const found = userSheet.getRange("D:D").createTextFinder(userEmail).matchEntireCell(true).findNext();
      if (found) {
        const row = found.getRow();
        const lastName = userSheet.getRange(row, 2).getValue();
        const firstName = userSheet.getRange(row, 3).getValue();
        senderName = `${lastName} ${firstName}`;
        senderDept = userSheet.getRange(row, 8).getValue();
      }
    } catch(e) {
      console.warn("User info lookup failed, using client data.", e);
    }
    sheet.appendRow([
      id,
      userEmail,
      senderName,
      senderDept,
      data.type,
      data.category,
      data.content,
      "未読",
      timestamp,
      ""
    ]);
    if (data.type === "緊急") {
      const adminEmails = getAdminEmails();
      if (adminEmails.length > 0) {
        MailApp.sendEmail({
          to: adminEmails.join(","),
          subject: "【FesTrack/緊急】新しい問い合わせがあります",
          body: `緊急の問い合わせを受信しました。\n\n送信者: ${data.name} (${data.dept})\nカテゴリ: ${data.category}\n\n内容:\n${data.content}\n\n管理画面を確認してください。`
        });
      }
    }
    clearLargeCache("INQUIRIES_BASE_DATA");
    return "送信しました。";
  });
}
function getUserMessages() {
  const userEmail = Session.getActiveUser().getEmail();
  let userCreatedAt = new Date(0);
  const userList = getUserList(); // キャッシュ化されたユーザーリストを取得
  const u = userList.find(u => u.email.toLowerCase() === userEmail.toLowerCase());
  if (u && u.createdAt) { // createdAtがUSER_MGMTにある前提（必要ならgetUserListを要改修）
     userCreatedAt = new Date(u.createdAt);
  }
  const iData = getBaseInquiryData();
  const myInquiries = [];
  if (iData.length > 1) {
    for (let i = 1; i < iData.length; i++) {
      if (iData[i][1] === userEmail) {
        myInquiries.push({
          id: iData[i][0], type: 'inquiry', subject: `問い合わせ: ${iData[i][5]}`, 
          content: iData[i][6], status: iData[i][7], date: iData[i][8], reply: iData[i][9] 
        });
      }
    }
  }
  const nData = getBaseNoticeData();
  const myNotifications = [];
  if (nData.length > 1) {
    for (let i = 1; i < nData.length; i++) {
      const target = nData[i][1];
      const noticeDate = new Date(nData[i][4]);
      if (target === userEmail || (target === 'ALL' && noticeDate >= userCreatedAt)) {
        myNotifications.push({
          id: nData[i][0], type: 'notice', subject: nData[i][2],
          content: nData[i][3], date: noticeDate
        });
      }
    }
  }
  const result = [...myInquiries, ...myNotifications].sort((a,b) => new Date(b.date) - new Date(a.date));
  return result.map(r => {
    r.date = new Date(r.date).toISOString();
    return r;
  });
}
function getAdminEmails() {
  const sheet = getSheet(SHEET_NAME.USER_MGMT);
  const data = sheet.getDataRange().getValues();
  const emails = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][5] === 'オーナー' || data[i][5] === '管理者') && data[i][3]) {
      emails.push(data[i][3]);
    }
  }
  return emails;
}
function getBaseInquiryData() {
  const cacheKey = "INQUIRIES_BASE_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.INQUIRIES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  setLargeCache(cacheKey, JSON.stringify(data));
  return data;
}
function getBaseNoticeData() {
  const cacheKey = "NOTICES_BASE_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.NOTIFICATIONS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  setLargeCache(cacheKey, JSON.stringify(data));
  return data;
}
function getAdminInquiries() {
  return withLock(() => {
    const data = getBaseInquiryData();
    if (data.length === 0) return [];
    const result = [];
    const now = new Date();
    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[7]; 
      const dateVal = row[8]; 
      const date = new Date(dateVal);
      const diffDays = (now - date) / (1000 * 60 * 60 * 24);
      if ((status === '完了' || status === '返信済' || status === '既読') && diffDays > 30) {
        rowsToDelete.push(i + 1);
        continue;
      }
      result.push({
        id: row[0], email: row[1], name: row[2], dept: row[3],
        type: row[4], category: row[5], content: row[6],
        status: status, date: dateVal, reply: row[9]
      });
    }
    if (rowsToDelete.length > 0) {
      const sheet = getSheet(SHEET_NAME.INQUIRIES);
      rowsToDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
      clearLargeCache("INQUIRIES_BASE_DATA"); // 行削除が起きたらキャッシュクリア
    }
    return result.reverse().map(r => {
      try { r.date = new Date(r.date).toISOString(); } catch(e){}
      return r;
    });
  });
}
function getAdminMessages() {
  const iData = getBaseInquiryData();
  const inquiries = [];
  if (iData.length > 1) {
    for (let i = 1; i < iData.length; i++) {
      inquiries.push({
        id: iData[i][0], type: 'inquiry', status: iData[i][7],
        category: iData[i][5], content: iData[i][6], date: iData[i][8],
        userName: iData[i][2], userEmail: iData[i][1], userDept: iData[i][3], reply: iData[i][9]
      });
    }
  }
  const nData = getBaseNoticeData();
  const notifications = [];
  if (nData.length > 1) {
    for (let i = 1; i < nData.length; i++) {
      notifications.push({
        id: nData[i][0], type: 'notice', target: nData[i][1],
        subject: nData[i][2], content: nData[i][3], date: nData[i][4], status: '配信済' 
      });
    }
  }
  const result = [...inquiries, ...notifications].sort((a, b) => new Date(b.date) - new Date(a.date));
  return result.map(r => {
    r.date = new Date(r.date).toISOString();
    return r;
  });
}
function getSimpleUserMessages() {
  const userEmail = Session.getActiveUser().getEmail();
  const now = new Date();
  let lastAccessTime = new Date(0);
  let userCreatedAt = new Date(0); 
  const sheet = getSheet(SHEET_NAME.USER_MGMT);
  if (sheet) {
    const finder = sheet.getRange("D:D").createTextFinder(userEmail).matchEntireCell(true).findNext();
    if (finder) {
      const row = finder.getRow();
      const createdAtVal = sheet.getRange(row, 4).getValue();
      if (createdAtVal && createdAtVal instanceof Date) userCreatedAt = createdAtVal;
      
      const lastAccessCell = sheet.getRange(row, 9); // I列
      const val = lastAccessCell.getValue();
      if (val && val instanceof Date) lastAccessTime = val;
      
      lastAccessCell.setValue(now);
    }
  }
  const myMessages = [];
  const iData = getBaseInquiryData();
  if (iData.length > 1) {
    for (let i = 1; i < iData.length; i++) {
      if (iData[i][1] === userEmail && iData[i][9]) { 
        const replyDate = iData[i][8];
        myMessages.push({
          id: iData[i][0], type: 'reply', subject: `Re: ${iData[i][5]}`, 
          content: iData[i][9], date: replyDate,
          isUnread: new Date(replyDate) > lastAccessTime
        });
      }
    }
  }
  const nData = getBaseNoticeData();
  if (nData.length > 1) {
    for (let i = 1; i < nData.length; i++) {
      const target = nData[i][1];
      const noticeDate = new Date(nData[i][4]);
      if (target === userEmail || (target === 'ALL' && noticeDate >= userCreatedAt)) {
        myMessages.push({
          id: nData[i][0], type: 'notice', subject: nData[i][2],
          content: nData[i][3], date: noticeDate,
          isUnread: noticeDate > lastAccessTime
        });
      }
    }
  }
  const result = myMessages.sort((a, b) => new Date(b.date) - new Date(a.date));
  return result.map(r => {
    r.date = new Date(r.date).toISOString();
    return r;
  });
}
function replyToInquiry(id, replyContent) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.INQUIRIES);
    const found = sheet.getRange("A:A").createTextFinder(id).matchEntireCell(true).findNext();
    if (!found) throw new Error("データが見つかりません");
    const row = found.getRow();
    sheet.getRange(row, 8).setValue("返信済"); // H列 Status
    sheet.getRange(row, 10).setValue(replyContent); // J列 Reply
    clearLargeCache("INQUIRIES_BASE_DATA");
    return "返信しました。";
  });
}
function markInquiryAsRead(id) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.INQUIRIES);
    const found = sheet.getRange("A:A").createTextFinder(id).matchEntireCell(true).findNext();
    if (!found) throw new Error("データが見つかりません");
    const row = found.getRow();
    const currentStatus = sheet.getRange(row, 8).getValue();
    if (currentStatus !== '返信済') {
      sheet.getRange(row, 8).setValue("既読");
      clearLargeCache("INQUIRIES_BASE_DATA");
    }
    return "既読にしました。";
  });
}
function sendBroadcast(target, subject, message) {
  const sheet = getSheet(SHEET_NAME.NOTIFICATIONS);
  const id = Utilities.getUuid();
  sheet.appendRow([id, target, subject, message, new Date()]);
  clearLargeCache("NOTICES_BASE_DATA");
  return "送信しました。";
}
function updateNoticeRecord(id, subject, content) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.NOTIFICATIONS); 
    if (!sheet) throw new Error("お知らせシートが見つかりません。");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) { // A列がIDと仮定
        sheet.getRange(i + 1, 3).setValue(subject); // C列: 件名
        sheet.getRange(i + 1, 4).setValue(content); // D列: 本文
        sheet.getRange(i + 1, 5).setValue(new Date()); // E列: 日付を更新
        clearLargeCache("NOTICES_BASE_DATA");
        return "お知らせの内容を更新しました。";
      }
    }
    throw new Error("指定されたお知らせデータが見つかりません。");
  });
}
function deleteNoticeRecord(id) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.NOTIFICATIONS);
    if (!sheet) throw new Error("お知らせシートが見つかりません。");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        clearLargeCache("NOTICES_BASE_DATA");
        return "お知らせを削除しました。";
      }
    }
    throw new Error("指定されたお知らせデータが見つかりません。");
  });
}
function sendWeeklyDigest() {
  const sheet = getSheet(SHEET_NAME.INQUIRIES);
  const data = sheet.getDataRange().getValues();
  let unreadCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === '未読' && data[i][4] === '通常') {
      unreadCount++;
    }
  }
  if (unreadCount > 0) {
    const adminEmails = getAdminEmails();
    if (adminEmails.length > 0) {
      MailApp.sendEmail({
        to: adminEmails.join(","),
        subject: "【FesTrack】週次問い合わせレポート",
        body: `現在、${unreadCount} 件の未読問い合わせ（通常）があります。\n管理画面を確認してください。`
      });
    }
  }
}
function triggerAnnualUpdate() {
  if (new Date().getMonth() !== 3) return;
  const now = new Date();
  const currentYear = now.getFullYear();
  const subject = `期間再設定のお願い【FesTrack】`;
  const message = `
昨年度は、本サービスをご利用いただき誠にありがとうございました。
新歓オリエンテーションを目前に控える中ですが、各種設定内の次の期間について日付の変更をお願いいたします。
・本祭期間：本祭開始日、本祭終了日
・受付設定期間：各フェーズ通常期間・再提出期間それぞれの 開始日、終了日
※年度は、システムにより自動で本年度（${currentYear}年）に更新しています。
本年度も、よろしくお願い申し上げます。
`.trim();
  sendBroadcast("管理者", subject, message);
  updateSettingsYearTo(currentYear);
}
function updateSettingsYearTo(newYear) {
  const sheet = getSheet(SHEET_NAME.SETTINGS);
  if (!sheet) return;
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const oldYear = newYear - 1;
  const regex = new RegExp(`(\\D|^)${oldYear}([-/年])`, "g");
  const newData = data.map(row => {
    return row.map(cell => {
      if (typeof cell === 'string') {
        return cell.replace(regex, `$1${newYear}$2`);
      }
      return cell;
    });
  });
  dataRange.setValues(newData);
}
function triggerNewYearGreeting() {
  if (new Date().getMonth() !== 0) return;
  const now = new Date();
  const newYear = now.getFullYear();
  const oldYear = newYear - 1;
  const subject = `新年のご挨拶【FesTrack】`;
  const message = `
新年、明けましておめでとうございます。
旧年中は、本サービスをご利用いただき、また梨医祭${oldYear}の円滑な運営にご協力いただき、誠にありがとうございました。
本年も、更なるサービスの向上に努めてまいりますので、よろしくお願い申しあげます。
${newYear}年 元日
  `.trim();
  sendBroadcast("ALL", subject, message);
}
function cleanupOldMessages() {
  const sheet = getSheet(SHEET_NAME.NOTIFICATIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const now = new Date();
  const sixMonthsAgo = new Date(now.setMonth(now.getMonth() - 6));
  for (let i = data.length - 1; i >= 1; i--) {
    const dateCell = data[i][4]; // E列
    if (!dateCell) continue;
    const rowDate = new Date(dateCell);
    if (!isNaN(rowDate.getTime()) && rowDate < sixMonthsAgo) {
      sheet.deleteRow(i + 1);
    }
  }
}
function getNoticeTemplates() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('NOTICE_TEMPLATES');
  const customTemplates = saved ? JSON.parse(saved) : [];
  const defaultTemplates = [
    { name: "不具合・エラー", subject: "【不具合情報】現在確認されている不具合について", content: "いつもFesTrackをご利用いただきありがとうございます。\n現在、以下の不具合が確認されております。\n\n【発生日時】〇月〇日ごろ～\n【不具合内容】\n【現在の状況・対応】\n\nご迷惑をおかけし申し訳ございません。復旧まで今しばらくお待ちください。" },
    { name: "緊急メンテ", subject: "【重要】緊急メンテナンス実施のお知らせ（月／日）", content: "いつもFesTrackをご利用いただきありがとうございます。\nシステムの安定稼働のため、以下の日程で緊急メンテナンスを実施いたします。\n\n【メンテナンス日時】\n〇月〇日 〇:〇 〜 〇:〇（予定）\n\n※メンテナンス中はシステムをご利用いただけません。\n※作業状況によってはメンテナンス終了時間を延長する場合があります。\nご協力のほどよろしくお願いいたします。" },
    { name: "システム改修", subject: "新機能or仕様変更or不具合解消orシステムアップデートのお知らせ", content: "いつもFesTrackをご利用いただきありがとうございます。\nシステムをアップデートし、以下の機能追加・修正を行いました。\n\n【変更後のバージョン】version x.x.x\n【主な変更点】\n・\n・\n\n引き続きよろしくお願いいたします。" }
  ];
  return defaultTemplates.concat(customTemplates);
}
function saveCustomNoticeTemplate(name, subject, content) {
  return withLock(() => {
    const props = PropertiesService.getScriptProperties();
    const saved = props.getProperty('NOTICE_TEMPLATES');
    const customTemplates = saved ? JSON.parse(saved) : [];
    customTemplates.push({ name: name, subject: subject, content: content });
    props.setProperty('NOTICE_TEMPLATES', JSON.stringify(customTemplates));
    return "テンプレートを保存しました。";
  });
}
// ==========================================
// 22. 販売業務関連
// ==========================================
function getAvailableShops() {
  const cacheKey = "AVAILABLE_SHOPS_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.SALES_CONFIG);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues().slice(1);
  const uniqueShops = {};
  data.forEach(row => {
    const parentCode = row[1]; // B列
    const name = row[2];       // C列
    uniqueShops[parentCode] = { code: parentCode, name: name, type: row[4] };
  });
  const results = Object.values(uniqueShops);
  setLargeCache(cacheKey, JSON.stringify(results));
  return results;
}
function getOrInitShopSession(parentCode) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.SALES_CONFIG);
    const data = sheet.getDataRange().getValues();
    const thisYear = new Date().getFullYear();
    // A. 今年のデータが既にあるか探す
    let currentRow = data.find(row => row[1] === parentCode && row[3] == thisYear);
    if (currentRow) {
      const result = {
        shopUid: currentRow[0],
        name: currentRow[2],
        type: currentRow[4],
        config: JSON.parse(currentRow[5] || '{}'),
        isNew: false
      };
      return result;
    }
    // B. 今年のデータがない -> 新規作成（去年の設定を引き継ぐ）
    const lastYearRow = data
      .filter(row => row[1] === parentCode)
      .sort((a, b) => b[3] - a[3])[0];
    let baseConfig = lastYearRow ? JSON.parse(lastYearRow[5] || '{}') : {};
    let baseName = lastYearRow ? lastYearRow[2] : "新規店舗";
    let baseType = lastYearRow ? lastYearRow[4] : "FreeInput";
    const newUid = `${parentCode}_${thisYear}`;
    const newRow = [
      newUid,      // A: UID
      parentCode,  // B: ParentCode
      baseName,    // C: Name
      thisYear,    // D: Year
      baseType,    // E: Type
      JSON.stringify(baseConfig) // F: Config (引継ぎ済)
    ];
    sheet.appendRow(newRow);
    const newResult = {
      shopUid: newUid,
      name: baseName,
      type: baseType,
      config: baseConfig,
      isNew: true
    };
    return newResult;
  });
}
function updateShopConfig(shopUid, newConfig) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.SALES_CONFIG);
    const range = sheet.getRange("A:A");
    const found = range.createTextFinder(shopUid).matchEntireCell(true).findNext();
    if (!found) throw new Error("店舗データが見つかりません");
    sheet.getRange(found.getRow(), 6).setValue(JSON.stringify(newConfig));
    clearLargeCache("AVAILABLE_SHOPS_DATA");
    return "メニューを更新しました";
  });
}
// 2. 売上登録
function registerSaleTransaction(data) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.SALES_TRANSACTIONS);
    const id = "TX-" + new Date().getTime().toString(24);
    const timestamp = new Date();
    const row = [
      id,
      data.shopId,
      timestamp,
      data.staffName,
      JSON.stringify(data.items),
      data.totalAmount,
      data.paymentMethod,
      data.status || 'Completed'
    ];
    sheet.appendRow(row);
    return { success: true, txId: id, timestamp: timestamp.toString() };
  });
}
// 3. 分析データの取得
function getSalesAnalysisData(shopId) {
  // 売上データ
  const txSheet = getSheet(SHEET_NAME.SALES_TRANSACTIONS);
  const txData = txSheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] === shopId);
  // 経費データ
  const exSheet = getSheet(SHEET_NAME.SALES_EXPENSES);
  const exData = exSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] === shopId);
  // 過去データ（同じShopNameの過去年度）も取得するロジックをここに追加可能
  return {
    transactions: txData.map(r => ({
      date: r[2], total: r[5], items: JSON.parse(r[4])
    })),
    expenses: exData.map(r => ({ item: r[3], amount: r[4] }))
  };
}
// 店舗追加用関数
function addNewShop(data) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.SALES_CONFIG);
    const parentCode = `SHOP_P_${Math.floor(Math.random() * 100000)}`;
    const uid = `${parentCode}_${data.year}`;
    let initialConfig = {};
    if(data.type === 'Preset') {
      initialConfig = { products: [] };
    } else {
      initialConfig = { hasQty: true };
    }
    const row = [
      uid,           // A列: 今年のユニークID
      parentCode,    // B列: ★追加 (親コード)
      data.name,     // C列: 店舗名
      data.year,     // D列: 年度
      data.type,     // E列: タイプ
      JSON.stringify(initialConfig) // F列: 設定JSON
    ];
    sheet.appendRow(row);
    clearLargeCache("AVAILABLE_SHOPS_DATA");
    return "OK";
  });
}
function updateTransactionStatus(txId, newStatus) {
  const sheet = getSheet(SHEET_NAME.SALES_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === txId) {
      sheet.getRange(i + 1, 8).setValue(newStatus); 
      return "Updated";
    }
  }
}
function getRecentOrders(shopUid) {
  const sheet = getSheet(SHEET_NAME.SALES_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const results = [];
  const targetId = String(shopUid || "").trim();
  for (let i = data.length - 1; i > 0; i--) {
    const row = data[i];
    if (!row[0]) continue; 
    const rowShopId = String(row[1] || "").trim();
    if (targetId && rowShopId && rowShopId !== targetId) {
      continue;
    }
    let itemsArr = [];
    try { itemsArr = JSON.parse(row[4] || '[]'); } catch(e){}
    let safeDate = "";
    if (row[2] instanceof Date) {
      safeDate = row[2].toISOString();
    } else {
      safeDate = String(row[2]);
    }
    results.push({
      id: row[0],
      timestamp: safeDate,
      items: itemsArr,
      status: row[7] || 'Making'
    });
    if (results.length >= 50) break;
  }
  return JSON.stringify(results);
}
function getSalesAnalysisData(shopUid) {
  const sheet = getSheet(SHEET_NAME.SALES_TRANSACTIONS);
  if (!sheet) {
    return null;
  }
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  const today = new Date();
  const currentYear = today.getFullYear();
  const lastYear = currentYear - 1;
  const currentYearData = [];
  const lastYearData = [];
  rows.forEach(row => {
    if (String(row[1]) !== String(shopUid)) return;
    const date = new Date(row[2]);
    const year = date.getFullYear();
    if (year === currentYear) {
      currentYearData.push(parseRowData(row));
    } else if (year === lastYear) {
      lastYearData.push(parseRowData(row));
    }
  });
  return {
    current: aggregateSalesData(currentYearData),
    last: aggregateSalesData(lastYearData),
    shopName: ""
  };
}
function getShopName(shopUid, data) {
  return ""; 
}
function parseRowData(row) {
  let items = [];
  try { items = JSON.parse(row[4] || '[]'); } catch(e){}
  let amountStr = String(row[5]).replace(/[^0-9.-]/g, '');
  let amount = Number(amountStr) || 0;
  return {
    timestamp: new Date(row[2]),
    items: items,
    amount: amount,
    method: row[6],
    status: row[7]
  };
}
function aggregateSalesData(rows) {
  const summary = {
    totalSales: 0,
    totalCount: 0,
    customerUnit: 0,
    totalRecordedQty: 0, 
    methods: { Cash: 0, PayPay: 0, CashCount: 0, PayPayCount: 0 },
    items: {}, 
    hourly: {}, 
    donation: 0,
    discounts: {
      ticketCount: 0, ticketAmount: 0,
      manualCount: 0, manualAmount: 0
    }
  };
  rows.forEach(r => {
    if (r.status !== 'Provided' && r.status !== 'Completed') return;
    summary.totalSales += r.amount;
    summary.totalCount++;
    if (r.method === '現金' || r.method === 'Cash') {
      summary.methods.Cash += r.amount;
      summary.methods.CashCount++;
    } else {
      summary.methods.PayPay += r.amount;
      summary.methods.PayPayCount++;
    }
    const hour = r.timestamp.getHours();
    if (!summary.hourly[hour]) summary.hourly[hour] = { sales: 0, count: 0 };
    summary.hourly[hour].sales += r.amount;
    summary.hourly[hour].count++;
    let itemSubtotal = 0;
    r.items.forEach(item => {
      itemSubtotal += (item.price * item.qty);
      if (item.name.includes('支援') || item.name.includes('寄付')) {
        summary.donation += (item.price * item.qty);
        return;
      }
      if (item.name.includes('割引券')) {
        summary.discounts.ticketCount++;
        summary.discounts.ticketAmount += Math.abs(item.price * item.qty);
        return;
      }
      if (item.name.includes('値引') || item.name.includes('Discount')) {
        summary.discounts.manualCount++;
        summary.discounts.manualAmount += Math.abs(item.price * item.qty);
        return;
      }
      if (!summary.items[item.name]) {
        summary.items[item.name] = { sales: 0, qty: 0, recordedQty: 0 };
      }
      summary.items[item.name].sales += (item.price * item.qty);
      summary.items[item.name].qty += item.qty;
      const rQty = (item.recordedQty !== undefined && item.recordedQty !== null) ? item.recordedQty : item.qty;
      summary.items[item.name].recordedQty += rQty;
      summary.totalRecordedQty += rQty;
    });
    if (r.amount < itemSubtotal && r.amount >= 0) {
      const diff = itemSubtotal - r.amount;
      summary.discounts.manualAmount += diff;
      summary.discounts.manualCount++;
    }
  });
  summary.customerUnit = summary.totalCount > 0 ? Math.round(summary.totalSales / summary.totalCount) : 0;
  return summary;
}
function getCsvDownloadData(shopUid) {
  const sheetName = (typeof SHEET_NAME !== 'undefined' && SHEET_NAME.SALES_TRANSACTIONS) ? SHEET_NAME.SALES_TRANSACTIONS : 'SALES_TRANSACTIONS';
  const sheet = getSheet(sheetName);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear();
  let totalSales = 0;
  const targetRows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[1]) !== String(shopUid)) continue;
    const date = new Date(row[2]);
    if (date.getFullYear() !== currentYear) continue;
    const amountStr = String(row[5]).replace(/[^0-9.-]/g, '');
    const amount = Number(amountStr) || 0;
    totalSales += amount;
    let itemsStr = "";
    try {
      const items = JSON.parse(row[4] || '[]');
      const isFlea = items.some(it => it.name.includes('商品') && !it.name.includes('タピオカ')); 
      if (isFlea) {
        const mainItem = items.find(it => it.price > 0 && !it.name.includes('支援'));
        itemsStr = (mainItem && mainItem.recordedQty) ? `個数: ${mainItem.recordedQty}` : "";
      } else {
        itemsStr = items.map(it => {
          if(it.price < 0) return `${it.name}(${it.price})`; 
          return `${it.name} x${it.qty}`;
        }).join(" / ");
      }
    } catch(e) { itemsStr = "データ破損"; }
    targetRows.push({ time: date, content: itemsStr, amount: amount });
  }
  targetRows.sort((a, b) => a.time - b.time);
  const csvHeader1 = `店舗ID:${shopUid} (${currentYear}年),,総売上: ¥${totalSales}`;
  const csvHeader2 = "会計時間,注文内容,会計金額";
  let csvBody = targetRows.map(r => {
    const timeStr = Utilities.formatDate(r.time, Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    const contentEscaped = `"${r.content.replace(/"/g, '""')}"`;
    return `${timeStr},${contentEscaped},${r.amount}`;
  }).join("\n");
  return csvHeader1 + "\n" + csvHeader2 + "\n" + csvBody;
}
function cancelTransaction(txId) {
  const sheet = getSheet(SHEET_NAME.SALES_TRANSACTIONS);
  if (!sheet) throw new Error("売上シートが見つかりません");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === txId) {
      sheet.getRange(i + 1, 8).setValue('Cancelled');
      return true;
    }
  }
  throw new Error("指定された取引IDが見つかりません");
}
// ==========================================
// 23. FAQ 取得処理
// ==========================================
function getFAQList() {
  const cacheKey = "FAQ_LIST_DATA";
  const cachedData = getLargeCache(cacheKey);
  if (cachedData) return JSON.parse(cachedData);
  const sheet = getSheet(SHEET_NAME.FAQ);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  const results = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[4] === '公開') {
      results.push({
        id: row[0],
        category: row[1],
        question: row[2],
        answer: row[3]
      });
    }
  }
  setLargeCache(cacheKey, JSON.stringify(results));
  return results;
}
function archiveAttendanceToSheet(eventId, eventData) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.ATTENDANCE);
    if (!sheet) {
      throw new Error(`シート「${SHEET_NAME.ATTENDANCE}」が見つかりません。事前に作成してください。`);
    }
    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === eventId) {
        targetRow = i + 1;
        break;
      }
    }
    const now = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");
    const rowData = [
      eventId,
      eventData.title,
      eventData.createdAt,
      JSON.stringify(eventData.labels || []),
      JSON.stringify(eventData.records || {}),
      now
    ];
    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, 6).setValues([rowData]);
      return `「${eventData.title}」の記録を更新しました。`;
    } else {
      sheet.appendRow(rowData);
      return `「${eventData.title}」を新たにシートへ記録しました。`;
    }
  });
}
// ==========================================
// 24. 企画管理センター
// ==========================================
function logicalDeleteProject(projectId, reasonType, details) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  if (!sheet) throw new Error("「企画データ」シートが見つかりません。");
  const data = sheet.getDataRange().getValues();
  const remarks = `${reasonType} / ${details}`;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      const row = i + 1;
      const currentStatus = data[i][1];
      sheet.getRange(row, 2).setValue("削除");
      sheet.getRange(row, 15).setValue("削除");
      sheet.getRange(row, 21).setValue("削除");
      sheet.getRange(row, 12).setValue(remarks);
      sheet.getRange(row, 23).setValue(remarks);
      sheet.getRange(row, 24).setValue(remarks);
      recordOperationHistory({
        targetId: projectId,
        type: "企画削除",
        from: currentStatus,
        to: "deleted",
        remarks: remarks
      });
      clearLargeCache("ADMIN_PROJECT_LIST");
      return true;
    }
  }
  throw new Error("対象の企画が見つかりませんでした。");
}
function releaseAccountBinding(projectId) {
  const sheet = getSheet(SHEET_NAME.PROJECTS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      const status = data[i][1];
      if (status !== '削除') {
        throw new Error("この企画は削除ステータスではないため、解除できません。");
      }
      const row = i + 1;
      sheet.getRange(row, 2).setValue("連携解除");
      sheet.getRange(row, 15).setValue("連携解除");
      sheet.getRange(row, 21).setValue("連携解除");
      sheet.getRange(row, 6).setValue("アカウント連携解除済み: " + new Date());
      sheet.getRange(row, 7).setValue("アカウント連携解除済み: " + new Date());
      recordOperationHistory({
        targetId: projectId,
        type: "連携解除",
        from: "deleted",
        to: "emilinated",
        remarks: "再申請許可のための紐づけ解除"
      });
      clearLargeCache("ADMIN_PROJECT_LIST");
      return true;
    }
  }
  throw new Error("対象の企画IDが見つかりませんでした。");
}
function saveCommitteeProjectData(data) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    if (!sheet) throw new Error("企画データシートが見つかりません。");
    const db = sheet.getDataRange().getValues();
    const now = new Date();
    const timestampString = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
    const basicInfoJson = JSON.stringify({
      memberCount: data.count,
      summary: data.summary
    });
    const locInfoJson = JSON.stringify({
      loc1: data.loc,
      timeRequest: data.time
    });
    let rowIndex = -1;
    for(let i = 1; i < db.length; i++) {
      if(db[i][0] === data.id) {
        rowIndex = i + 1;
        break;
      }
    }
    if(rowIndex !== -1) {
      sheet.getRange(rowIndex, 4).setValue(data.name);       // D列: 企画名
      sheet.getRange(rowIndex, 5).setValue(data.dept);       // E列: 所管局
      sheet.getRange(rowIndex, 10).setValue(basicInfoJson);  // J列: 基本情報JSON
      sheet.getRange(rowIndex, 13).setValue(locInfoJson);    // M列: 場所情報JSON
    } else {
      const colCount = sheet.getLastColumn() > 15 ? sheet.getLastColumn() : 25;
      const newRow = new Array(colCount).fill("");
      newRow[0] = data.id;               // A列: 企画ID
      newRow[1] = "主催";                // B列: "主催"（ステータス兼用）
      newRow[2] = "委員会主催";          // C列: "委員会主催"（形態）
      newRow[3] = data.name;             // D列: 企画名
      newRow[4] = data.dept;             // E列: 所管局
      newRow[9] = basicInfoJson;         // J列: 基本情報JSON
      newRow[10] = timestampString;      // K列: タイムスタンプ
      newRow[12] = locInfoJson;          // M列: 場所情報JSON
      sheet.appendRow(newRow);
    }
    clearLargeCache("ADMIN_PROJECT_LIST");
    return true;
  });
}
function deleteCommitteeProjectData(id) {
  return withLock(() => {
    const sheet = getSheet(SHEET_NAME.PROJECTS);
    if (!sheet) throw new Error("企画データシートが見つかりません。");
    const db = sheet.getDataRange().getValues();
    for(let i = 1; i < db.length; i++) {
      if(db[i][0] === id) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue("削除");
        clearLargeCache("ADMIN_PROJECT_LIST");
        return true;
      }
    }
    throw new Error("対象の企画が見つかりませんでした。");
  });
}