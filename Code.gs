/**
 * RFID タオル在庫管理システム（Phase 2.1 — 独立フロントエンド対応）
 *
 * コア機能:
 *   - 入荷・出庫・入庫 の3スキャン操作
 *   - タグ単位の在庫状態管理
 *   - 現在庫一覧（品目別集計）
 *   - Googleスプレッドシート出力（物品順 / スキャン順）
 *   - REST API（外部フロントエンドからのfetchアクセス対応）
 *   - DataWedge / RFIDリーダー連携対応
 *
 * シート構成:
 *   タグマスタ / 品目マスタ / 物件マスタ / 在庫状態 / スキャン履歴 / 物品順 / スキャン順
 */

// ==================== 設定 ====================
const CONFIG = {
  SPREADSHEET_ID: '1ws95DVPF8EK3CCHnWjSMGVQVQZGI9jVn_E7Chv6zV6g',
  SHEETS: {
    TAG_MASTER:    'タグマスタ',
    ITEM_MASTER:   '品目マスタ',
    PROP_MASTER:   '物件マスタ',
    INVENTORY:     '在庫状態',
    SCAN_HISTORY:  'スキャン履歴',
    OUT_ITEM:      '物品順',
    OUT_SCAN:      'スキャン順',
    BARCODE_MASTER:'バーコードマスタ',
    STOCKTAKE_LOG: '棚卸し履歴',
  },
};

// ==================== REST API ====================

/**
 * GET リクエストハンドラ
 * ?action=inventory|history|tags|items|properties|ping
 * &limit=50 (history用)
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'ping';
  var result;

  try {
    switch (action) {
      case 'inventory':
        result = getInventorySummary();
        break;
      case 'history':
        var limit = parseInt(e.parameter.limit) || 50;
        result = getScanHistory(limit);
        break;
      case 'tags':
        result = getTestTagList();
        break;
      case 'items':
        result = getItemMaster();
        break;
      case 'properties':
        result = getPropertyMaster();
        break;
      case 'barcodes':
        result = getBarcodeMaster();
        break;
      case 'ping':
        result = { success: true, message: 'RFID在庫管理API稼働中', timestamp: new Date().toISOString() };
        break;
      default:
        result = { success: false, error: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST リクエストハンドラ
 * body: { action: "scan"|"init", ... }
 */
function doPost(e) {
  var result;

  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    switch (action) {
      case 'scan':
        result = executeScan(body.operationType, body.tagIds, body.propertyId || '', body.operator || '');
        break;
      case 'stainRemoval':
        result = executeStainRemoval(body.tagIds, body.operator || '');
        break;
      case 'init':
        result = initializeSystem();
        break;
      case 'barcodeStocktake':
        result = saveBarcodeStocktake(body.propertyId || '', body.items || [], body.operator || '');
        break;
      case 'addBarcodeMaster':
        result = addBarcodeMasterItem(body.jisCode || '', body.itemName || '', body.category || '', body.unit || '個');
        break;
      default:
        result = { success: false, error: '不明なアクション: ' + action };
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== ユーティリティ ====================

// 自社タグIDのプレフィックス（RFIDチップ種別に依存）
var TAG_PREFIX = 'E2806A96';
var TAG_LENGTH = 24; // 1タグIDの標準文字数

/**
 * タグID配列の前処理
 * - 結合されたタグIDの分割（DataWedge出力で連結される問題への対策）
 * - TAG_LENGTH超過分の切り詰め（自社タグ+非対象タグ結合ケース）
 * - 非対象タグ（プレフィックス不一致）の除外
 * - 大文字正規化
 * - 重複除去
 */
function _preprocessTagIds(tagIds) {
  var result = [];
  var seen = {};

  tagIds.forEach(function(raw) {
    var tid = String(raw).trim().toUpperCase();
    if (!tid) return;

    // TAG_PREFIXで分割を試みる
    if (TAG_PREFIX && tid.indexOf(TAG_PREFIX) !== -1) {
      var parts = [];
      var remaining = tid;
      while (remaining.length > 0) {
        var nextIdx = remaining.indexOf(TAG_PREFIX, 1);
        if (nextIdx > 0) {
          parts.push(remaining.substring(0, nextIdx));
          remaining = remaining.substring(nextIdx);
        } else {
          parts.push(remaining);
          remaining = '';
        }
      }
      parts.forEach(function(p) {
        if (p.length < 16) return;
        // TAG_LENGTH超過分を切り詰め（非対象タグが末尾に結合したケース）
        if (TAG_LENGTH && p.indexOf(TAG_PREFIX) === 0 && p.length > TAG_LENGTH) {
          p = p.substring(0, TAG_LENGTH);
        }
        // 非対象タグ（プレフィックス不一致）は除外
        if (TAG_PREFIX && p.indexOf(TAG_PREFIX) !== 0) return;
        if (!seen[p]) {
          seen[p] = true;
          result.push(p);
        }
      });
    } else {
      // プレフィックス不一致の単独タグは除外
      if (TAG_PREFIX && tid.indexOf(TAG_PREFIX) !== 0) return;
      if (!seen[tid]) {
        seen[tid] = true;
        result.push(tid);
      }
    }
  });

  return result;
}

function ss_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function sheet_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) s = ss_().insertSheet(name);
  return s;
}

function now_() { return new Date(); }

function uuid_() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 12);
}

// ==================== 初期化・マスタ ====================

function initializeSystem() {
  _initItemMaster();
  _initPropertyMaster();
  _initTagMaster();
  _initInventory();
  _initScanHistory();
  _initOutputSheets();
  initBarcodeSheets_();
  Logger.log('システム初期化完了');
  return { success: true, message: 'システム初期化完了' };
}

function _initItemMaster() {
  var s = sheet_(CONFIG.SHEETS.ITEM_MASTER);
  if (s.getLastRow() > 0) s.clearContents();
  s.appendRow(['品目コード', '品目名', '管理単位', 'RFID対象', '有効']);
  s.appendRow(['BT', 'バスタオル', '枚', true, true]);
  s.appendRow(['FT', 'フェイスタオル', '枚', true, true]);
  s.appendRow(['FM', 'フットマット', '枚', true, true]);
  s.getRange(1, 1, 1, 5).setFontWeight('bold');
}

function _initPropertyMaster() {
  var s = sheet_(CONFIG.SHEETS.PROP_MASTER);
  if (s.getLastRow() > 0) s.clearContents();
  s.appendRow(['物件ID', '物件名', '有効', '表示順']);
  s.appendRow(['P001', '物件A（渋谷）', true, 1]);
  s.appendRow(['P002', '物件B（新宿）', true, 2]);
  s.appendRow(['P003', '物件C（池袋）', true, 3]);
  s.appendRow(['P004', '物件D（品川）', true, 4]);
  s.getRange(1, 1, 1, 4).setFontWeight('bold');
}

function _initTagMaster() {
  var s = sheet_(CONFIG.SHEETS.TAG_MASTER);
  if (s.getLastRow() > 0) s.clearContents();
  s.appendRow(['タグID', '品目コード', '有効', '登録日', '備考']);
  s.getRange(1, 1, 1, 5).setFontWeight('bold');
  var d = now_();
  var rows = [];
  for (var i = 1; i <= 10; i++) {
    rows.push(['BT-' + ('000' + i).slice(-4), 'BT', true, d, 'テスト']);
  }
  for (var i = 1; i <= 10; i++) {
    rows.push(['FT-' + ('000' + i).slice(-4), 'FT', true, d, 'テスト']);
  }
  for (var i = 1; i <= 5; i++) {
    rows.push(['FM-' + ('000' + i).slice(-4), 'FM', true, d, 'テスト']);
  }
  if (rows.length > 0) {
    s.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function _initInventory() {
  var s = sheet_(CONFIG.SHEETS.INVENTORY);
  if (s.getLastRow() > 0) s.clearContents();
  s.appendRow(['タグID', '品目コード', 'ステータス', '所在', '最終搬入先', '最終操作', '最終スキャン日時', '最終更新日時', '洗濯回数', '染み抜き回数']);
  s.getRange(1, 1, 1, 10).setFontWeight('bold');
  var tagSheet = sheet_(CONFIG.SHEETS.TAG_MASTER);
  var lastRow = tagSheet.getLastRow();
  if (lastRow <= 1) return;
  var tags = tagSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var d = now_();
  var rows = tags.map(function(t) {
    return [t[0], t[1], 'クリーン', 'メイン倉庫', '', '初期登録', d, d, 0, 0];
  });
  s.getRange(2, 1, rows.length, 10).setValues(rows);
}

function _initScanHistory() {
  var s = sheet_(CONFIG.SHEETS.SCAN_HISTORY);
  if (s.getLastRow() > 0) s.clearContents();
  s.appendRow(['履歴ID', 'スキャン日時', '操作種別', '搬入先物件', 'タグID', '品目コード', '実行者', '警告', '異常種別', '送信状態']);
  s.getRange(1, 1, 1, 10).setFontWeight('bold');
}

function _initOutputSheets() {
  var s1 = sheet_(CONFIG.SHEETS.OUT_ITEM);
  if (s1.getLastRow() > 0) s1.clearContents();
  s1.appendRow(['ID(タグID)', '品目', 'ステータス', '搬入先', '実行者', 'スキャン日時']);
  s1.getRange(1, 1, 1, 6).setFontWeight('bold');
  var s2 = sheet_(CONFIG.SHEETS.OUT_SCAN);
  if (s2.getLastRow() > 0) s2.clearContents();
  s2.appendRow(['スキャン日時', '操作種別', '搬入先', 'タグID', '品目', '数量', '実行者']);
  s2.getRange(1, 1, 1, 7).setFontWeight('bold');
}

// ==================== マスタ取得 ====================

function getItemMaster() {
  try {
    var s = sheet_(CONFIG.SHEETS.ITEM_MASTER);
    var lr = s.getLastRow();
    if (lr <= 1) return { success: true, items: [] };
    var data = s.getRange(2, 1, lr - 1, 5).getValues();
    var items = data.filter(function(r) { return r[4]; }).map(function(r) {
      return { code: r[0], name: r[1], unit: r[2] };
    });
    return { success: true, items: items };
  } catch (e) { return { success: false, error: e.message }; }
}

function getPropertyMaster() {
  try {
    var s = sheet_(CONFIG.SHEETS.PROP_MASTER);
    var lr = s.getLastRow();
    if (lr <= 1) return { success: true, properties: [] };
    var data = s.getRange(2, 1, lr - 1, 4).getValues();
    var props = data.filter(function(r) { return r[2]; }).map(function(r) {
      return { id: r[0], name: r[1], order: r[3] };
    });
    props.sort(function(a, b) { return a.order - b.order; });
    return { success: true, properties: props };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==================== タグ → 品目 解決 ====================

function _getTagItemMap() {
  var s = sheet_(CONFIG.SHEETS.TAG_MASTER);
  var lr = s.getLastRow();
  if (lr <= 1) return {};
  var data = s.getRange(2, 1, lr - 1, 3).getValues();
  var map = {};
  data.forEach(function(r) {
    if (r[2]) map[r[0]] = r[1];
  });
  return map;
}

function _getItemNameMap() {
  var s = sheet_(CONFIG.SHEETS.ITEM_MASTER);
  var lr = s.getLastRow();
  if (lr <= 1) return {};
  var data = s.getRange(2, 1, lr - 1, 2).getValues();
  var map = {};
  data.forEach(function(r) { map[r[0]] = r[1]; });
  return map;
}

// ==================== コア：スキャン処理 ====================

function executeScan(operationType, tagIds, propertyId, operator) {
  try {
    if (!tagIds || tagIds.length === 0) {
      return { success: false, error: 'タグが読み取られていません' };
    }
    if (operationType === '出庫' && !propertyId) {
      return { success: false, error: '出庫時は搬入先物件を選択してください' };
    }

    // タグID前処理: 結合タグの分割 + 正規化
    tagIds = _preprocessTagIds(tagIds);
    if (tagIds.length === 0) {
      return { success: false, error: '有効なタグがありません' };
    }

    var tagItemMap = _getTagItemMap();
    var itemNameMap = _getItemNameMap();
    var invSheet = sheet_(CONFIG.SHEETS.INVENTORY);
    var histSheet = sheet_(CONFIG.SHEETS.SCAN_HISTORY);
    var d = now_();
    var op = operator || '共通アカウント';

    var invLR = invSheet.getLastRow();
    var invMap = {};
    if (invLR > 1) {
      var invData = invSheet.getRange(2, 1, invLR - 1, 10).getValues();
      invData.forEach(function(r, idx) {
        invMap[r[0]] = { row: idx + 2, data: r };
      });
    }

    var uniqueTags = [];
    var seen = {};
    tagIds.forEach(function(t) {
      var tid = String(t).trim();
      if (tid && !seen[tid]) {
        seen[tid] = true;
        uniqueTags.push(tid);
      }
    });

    var results = { processed: 0, warnings: [], itemCounts: {} };
    var histRows = [];

    uniqueTags.forEach(function(tagId) {
      var itemCode = tagItemMap[tagId];
      var warning = '';
      var errorType = '';

      if (!itemCode) {
        warning = '未登録タグ';
        errorType = '未登録';
        itemCode = '不明';
      }

      var newStatus, newLocation;
      if (operationType === '入荷') {
        newStatus = 'クリーン';
        newLocation = 'メイン倉庫';
      } else if (operationType === '出庫') {
        newStatus = '出庫中';
        newLocation = propertyId;
      } else if (operationType === '入庫') {
        newStatus = 'クリーン';
        newLocation = 'メイン倉庫';
      }

      var destination = (operationType === '出庫') ? propertyId : '';

      if (invMap[tagId]) {
        var r = invMap[tagId].row;
        var prevStatus = invMap[tagId].data[2]; // 変更前のステータス
        invSheet.getRange(r, 3).setValue(newStatus);
        invSheet.getRange(r, 4).setValue(newLocation);
        if (destination) invSheet.getRange(r, 5).setValue(destination);
        invSheet.getRange(r, 6).setValue(operationType);
        invSheet.getRange(r, 7).setValue(d);
        invSheet.getRange(r, 8).setValue(d);
        // 入庫時（出庫中→クリーン）に洗濯回数を+1
        if (operationType === '入庫' && prevStatus === '出庫中') {
          var washCount = invMap[tagId].data[8] || 0;
          invSheet.getRange(r, 9).setValue(washCount + 1);
        }
      } else {
        invSheet.appendRow([tagId, itemCode, newStatus, newLocation, destination, operationType, d, d, 0, 0]);
      }

      var iName = itemNameMap[itemCode] || itemCode;
      results.itemCounts[iName] = (results.itemCounts[iName] || 0) + 1;

      histRows.push([uuid_(), d, operationType, destination, tagId, itemCode, op, warning ? true : false, errorType, '送信済']);

      if (warning) results.warnings.push(tagId + ': ' + warning);
      results.processed++;
    });

    if (histRows.length > 0) {
      var hlr = histSheet.getLastRow();
      histSheet.getRange(hlr + 1, 1, histRows.length, 10).setValues(histRows);
    }

    _updateOutputSheets(operationType, uniqueTags, tagItemMap, itemNameMap, propertyId, op, d);

    return {
      success: true,
      operationType: operationType,
      totalScanned: uniqueTags.length,
      itemCounts: results.itemCounts,
      warnings: results.warnings,
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== 染み抜き処理 ====================

function executeStainRemoval(tagIds, operator) {
  try {
    if (!tagIds || tagIds.length === 0) {
      return { success: false, error: 'タグが読み取られていません' };
    }

    tagIds = _preprocessTagIds(tagIds);
    if (tagIds.length === 0) {
      return { success: false, error: '有効なタグがありません' };
    }

    var tagItemMap = _getTagItemMap();
    var itemNameMap = _getItemNameMap();
    var invSheet = sheet_(CONFIG.SHEETS.INVENTORY);
    var histSheet = sheet_(CONFIG.SHEETS.SCAN_HISTORY);
    var d = now_();
    var op = operator || '共通アカウント';

    var invLR = invSheet.getLastRow();
    var invMap = {};
    if (invLR > 1) {
      var invData = invSheet.getRange(2, 1, invLR - 1, 10).getValues();
      invData.forEach(function(r, idx) {
        invMap[r[0]] = { row: idx + 2, data: r };
      });
    }

    var uniqueTags = [];
    var seen = {};
    tagIds.forEach(function(t) {
      var tid = String(t).trim();
      if (tid && !seen[tid]) {
        seen[tid] = true;
        uniqueTags.push(tid);
      }
    });

    var results = { processed: 0, warnings: [], itemCounts: {}, tagDetails: [] };
    var histRows = [];

    uniqueTags.forEach(function(tagId) {
      var itemCode = tagItemMap[tagId];
      var warning = '';
      var errorType = '';

      if (!itemCode) {
        warning = '未登録タグ';
        errorType = '未登録';
        results.warnings.push(tagId + ': 未登録タグ');
        return; // 未登録タグは染み抜き対象外
      }

      if (invMap[tagId]) {
        var r = invMap[tagId].row;
        var stainCount = invMap[tagId].data[9] || 0;
        var newStainCount = stainCount + 1;
        invSheet.getRange(r, 10).setValue(newStainCount);
        invSheet.getRange(r, 8).setValue(d); // 最終更新日時

        var iName = itemNameMap[itemCode] || itemCode;
        results.itemCounts[iName] = (results.itemCounts[iName] || 0) + 1;
        results.tagDetails.push({
          tagId: tagId,
          itemName: iName,
          stainCount: newStainCount,
          washCount: invMap[tagId].data[8] || 0
        });
      } else {
        warning = '在庫未登録';
        errorType = '在庫未登録';
        results.warnings.push(tagId + ': 在庫に未登録');
      }

      histRows.push([uuid_(), d, '染み抜き', '', tagId, itemCode || '不明', op, warning ? true : false, errorType, '送信済']);
      results.processed++;
    });

    if (histRows.length > 0) {
      var hlr = histSheet.getLastRow();
      histSheet.getRange(hlr + 1, 1, histRows.length, 10).setValues(histRows);
    }

    return {
      success: true,
      operationType: '染み抜き',
      totalProcessed: results.processed,
      itemCounts: results.itemCounts,
      tagDetails: results.tagDetails,
      warnings: results.warnings,
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==================== スプレッドシート出力 ====================

function _updateOutputSheets(opType, tagIds, tagItemMap, itemNameMap, propertyId, operator, dateTime) {
  var s1 = sheet_(CONFIG.SHEETS.OUT_ITEM);
  var s2 = sheet_(CONFIG.SHEETS.OUT_SCAN);

  var itemRows = tagIds.map(function(tid) {
    var ic = tagItemMap[tid] || '不明';
    var iName = itemNameMap[ic] || ic;
    return [tid, iName, opType, propertyId || '', operator, dateTime];
  });
  if (itemRows.length > 0) {
    var lr1 = s1.getLastRow();
    s1.getRange(lr1 + 1, 1, itemRows.length, 6).setValues(itemRows);
  }

  var counts = {};
  tagIds.forEach(function(tid) {
    var ic = tagItemMap[tid] || '不明';
    var iName = itemNameMap[ic] || ic;
    counts[iName] = (counts[iName] || 0) + 1;
  });
  var scanRows = Object.keys(counts).map(function(itemName) {
    return [dateTime, opType, propertyId || '', '', itemName, counts[itemName], operator];
  });
  if (scanRows.length > 0) {
    var lr2 = s2.getLastRow();
    s2.getRange(lr2 + 1, 1, scanRows.length, 7).setValues(scanRows);
  }
}

// ==================== 現在庫一覧 ====================

function getInventorySummary() {
  try {
    var invSheet = sheet_(CONFIG.SHEETS.INVENTORY);
    var lr = invSheet.getLastRow();
    if (lr <= 1) return { success: true, summary: [], total: { clean: 0, all: 0, shipped: 0 } };

    var data = invSheet.getRange(2, 1, lr - 1, 10).getValues();
    var itemNameMap = _getItemNameMap();

    var stats = {};
    var allTags = []; // 棚卸し用：全タグIDリスト
    var tagLocationMap = {}; // 棚卸し用：タグID → 所在マップ
    var tagDetailMap = {}; // タグ詳細情報マップ
    data.forEach(function(r) {
      var tagId = r[0]; // タグID（A列）
      var itemCode = r[1];
      var status = r[2];
      var location = r[3] || ''; // 所在（D列）
      var washCount = r[8] || 0; // 洗濯回数（I列）
      var stainCount = r[9] || 0; // 染み抜き回数（J列）
      var itemName = itemNameMap[itemCode] || itemCode;
      if (!stats[itemName]) stats[itemName] = { name: itemName, code: itemCode, clean: 0, all: 0, shipped: 0 };
      stats[itemName].all++;
      if (status === 'クリーン') stats[itemName].clean++;
      if (status === '出庫中') stats[itemName].shipped++;
      if (tagId) {
        var tid = String(tagId).toUpperCase();
        allTags.push(tid);
        tagLocationMap[tid] = location;
        tagDetailMap[tid] = { washCount: washCount, stainCount: stainCount, status: status, itemName: itemName };
      }
    });

    var summary = Object.keys(stats).map(function(k) { return stats[k]; });
    summary.sort(function(a, b) { return a.code.localeCompare(b.code); });

    var total = { clean: 0, all: 0, shipped: 0 };
    summary.forEach(function(s) {
      total.clean += s.clean;
      total.all += s.all;
      total.shipped += s.shipped;
    });

    return { success: true, summary: summary, total: total, allTags: allTags, tagLocationMap: tagLocationMap, tagDetailMap: tagDetailMap };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==================== スキャン履歴取得 ====================

function getScanHistory(limit) {
  try {
    var s = sheet_(CONFIG.SHEETS.SCAN_HISTORY);
    var lr = s.getLastRow();
    if (lr <= 1) return { success: true, history: [] };

    var n = Math.min(limit || 50, lr - 1);
    var startRow = Math.max(2, lr - n + 1);
    var data = s.getRange(startRow, 1, lr - startRow + 1, 10).getValues();
    var itemNameMap = _getItemNameMap();

    var history = data.map(function(r) {
      return {
        id: r[0],
        datetime: r[1] ? Utilities.formatDate(new Date(r[1]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss') : '',
        operation: r[2],
        destination: r[3],
        tagId: r[4],
        itemCode: r[5],
        itemName: itemNameMap[r[5]] || r[5],
        operator: r[6],
        warning: r[7],
        errorType: r[8],
      };
    });
    history.reverse();
    return { success: true, history: history };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==================== テスト用：タグ一覧取得 ====================

function getTestTagList() {
  try {
    var s = sheet_(CONFIG.SHEETS.TAG_MASTER);
    var lr = s.getLastRow();
    if (lr <= 1) return { success: true, tags: [] };
    var data = s.getRange(2, 1, lr - 1, 3).getValues();
    var itemNameMap = _getItemNameMap();
    var tags = data.filter(function(r) { return r[2]; }).map(function(r) {
      return { id: r[0], itemCode: r[1], itemName: itemNameMap[r[1]] || r[1] };
    });
    return { success: true, tags: tags };
  } catch (e) { return { success: false, error: e.message }; }
}

// ==================== バーコード棚卸し ====================

/**
 * バーコードマスタ取得
 * シート構成: A=JISコード, B=品目名, C=カテゴリ, D=単位
 */
function getBarcodeMaster() {
  try {
    var s = sheet_(CONFIG.SHEETS.BARCODE_MASTER);
    var lr = s.getLastRow();
    if (lr <= 1) return { success: true, items: [] };
    var data = s.getRange(2, 1, lr - 1, 4).getValues();
    var items = data.filter(function(r) { return r[0]; }).map(function(r) {
      return {
        jisCode: String(r[0]).trim(),
        itemName: String(r[1]).trim(),
        category: String(r[2]).trim(),
        unit: String(r[3]).trim() || '個'
      };
    });
    return { success: true, items: items };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * バーコードマスタに新規アイテムを追加
 * 重複チェック（JISコード）あり
 */
function addBarcodeMasterItem(jisCode, itemName, category, unit) {
  try {
    if (!jisCode || !itemName) {
      return { success: false, error: 'JISコードと品目名は必須です' };
    }
    jisCode = String(jisCode).trim();
    itemName = String(itemName).trim();
    category = String(category).trim();
    unit = String(unit).trim() || '個';

    var s = sheet_(CONFIG.SHEETS.BARCODE_MASTER);
    var lr = s.getLastRow();

    // 重複チェック
    if (lr > 1) {
      var existingCodes = s.getRange(2, 1, lr - 1, 1).getValues();
      for (var i = 0; i < existingCodes.length; i++) {
        if (String(existingCodes[i][0]).trim() === jisCode) {
          return { success: false, error: 'JISコード ' + jisCode + ' は既に登録されています' };
        }
      }
    }

    s.appendRow([jisCode, itemName, category, unit]);

    return {
      success: true,
      message: itemName + ' を登録しました',
      item: { jisCode: jisCode, itemName: itemName, category: category, unit: unit }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * バーコード棚卸し結果をシートに保存
 * items: [{ jisCode, itemName, category, quantity }]
 */
function saveBarcodeStocktake(propertyId, items, operator) {
  try {
    var s = sheet_(CONFIG.SHEETS.STOCKTAKE_LOG);
    // ヘッダーがなければ作成
    if (s.getLastRow() === 0) {
      s.appendRow(['ID', '日時', '物件', '実行者', 'JISコード', '品目名', 'カテゴリ', '数量', '単位']);
      s.getRange(1, 1, 1, 9).setFontWeight('bold');
    }
    var now = new Date();
    var ts = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    var batchId = 'ST-' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMddHHmmss');

    items.forEach(function(item) {
      s.appendRow([
        batchId,
        ts,
        propertyId,
        operator,
        item.jisCode,
        item.itemName,
        item.category || '',
        item.quantity,
        item.unit || '個'
      ]);
    });

    return {
      success: true,
      batchId: batchId,
      itemCount: items.length,
      totalQuantity: items.reduce(function(sum, i) { return sum + (i.quantity || 0); }, 0)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * システム初期化時にバーコードマスタと棚卸し履歴シートも作成
 */
function initBarcodeSheets_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // バーコードマスタ
  var bcSheet = ss.getSheetByName(CONFIG.SHEETS.BARCODE_MASTER);
  if (!bcSheet) {
    bcSheet = ss.insertSheet(CONFIG.SHEETS.BARCODE_MASTER);
    bcSheet.appendRow(['JISコード', '品目名', 'カテゴリ', '単位']);
    bcSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    // サンプルデータ
    bcSheet.appendRow(['4901234567890', 'バスタオル（白）', 'タオル', '枚']);
    bcSheet.appendRow(['4901234567891', 'フェイスタオル（白）', 'タオル', '枚']);
    bcSheet.appendRow(['4901234567892', 'シャンプー 30ml', 'アメニティ', '個']);
    bcSheet.appendRow(['4901234567893', '歯ブラシセット', 'アメニティ', '個']);
    bcSheet.appendRow(['4901234567894', 'バスマット', '備品', '枚']);
    bcSheet.setColumnWidth(1, 160);
    bcSheet.setColumnWidth(2, 200);
    bcSheet.setColumnWidth(3, 100);
    bcSheet.setColumnWidth(4, 60);
  }

  // 棚卸し履歴
  var stSheet = ss.getSheetByName(CONFIG.SHEETS.STOCKTAKE_LOG);
  if (!stSheet) {
    stSheet = ss.insertSheet(CONFIG.SHEETS.STOCKTAKE_LOG);
    stSheet.appendRow(['ID', '日時', '物件', '実行者', 'JISコード', '品目名', 'カテゴリ', '数量', '単位']);
    stSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
}

// ==================== テスト実行 ====================

function testPhase2() {
  Logger.log('=== Phase 2 テスト ===');
  initializeSystem();
  var inv = getInventorySummary();
  Logger.log('初期在庫: ' + JSON.stringify(inv));
  var arrivalResult = executeScan('入荷', ['BT-0001', 'BT-0002', 'FT-0001'], '', '管理者');
  Logger.log('入荷結果: ' + JSON.stringify(arrivalResult));
  var shipResult = executeScan('出庫', ['BT-0001', 'FT-0001', 'FM-0001'], 'P001', 'スタッフA');
  Logger.log('出庫結果: ' + JSON.stringify(shipResult));
  var returnResult = executeScan('入庫', ['BT-0001', 'FT-0001'], '', 'スタッフA');
  Logger.log('入庫結果: ' + JSON.stringify(returnResult));
  var inv2 = getInventorySummary();
  Logger.log('最終在庫: ' + JSON.stringify(inv2));
  Logger.log('=== テスト完了 ===');
}
