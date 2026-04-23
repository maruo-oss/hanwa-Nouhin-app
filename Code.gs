// ============================================================
// 納品書管理GASウェブアプリ - バックエンド
// ============================================================

const SHEET_NAME = '納品書データ';
const HEADERS = [
  'row_key', 'file_id', 'file_name', 'status',
  '日時', '納品書No', '取引先', '作業所', '貸出期間', '請求計上日', '伝票摘要',
  '区分', '機種', '号機', '型式', '管理No', '数量', '単位', '単価', '金額', '基本管理料', '備考',
  'processed_at', 'item_order'
];

// --- プロパティ取得ヘルパー ---
function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// シートの実際のヘッダー行から列名→0始まり列indexのマップを作る
function getColumnMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const map = {};
  if (lastCol === 0) return map;
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  row.forEach(function(name, i) {
    const s = String(name || '').trim();
    if (s && map[s] === undefined) map[s] = i;
  });
  return map;
}

const PROMPT_SHEET_NAME = 'プロンプト';
const DEFAULT_PROMPT = `この納品書PDFを解析し、以下のJSON形式で情報を抽出してください。

出力形式:
{
  "header": {
    "日時": "YYYY-MM-DD形式の日付",
    "納品書No": "文書番号（No.欄の値）",
    "取引先": "納品書の発行元会社名（右上に記載の会社名。宛先の「阪和建機」ではなく、納品書を発行した側の会社名）",
    "作業所": "現場名（「現場名:」の値）",
    "貸出期間": "貸出期間の値（例: 2026年2月10日〜）",
    "請求計上日": "請求計上日の値（YYYY-MM-DD形式）",
    "伝票摘要": "伝票摘要の値"
  },
  "items": [
    {
      "区分": "販売またはレンタル（行の左端の区分）",
      "機種": "品名・規格欄の上段に型番がある場合のみ抽出（例: 「SGT JCL-015C」→「JCL-015C」）。左側のアルファベット略称（SGT, SSK等）は除き、右側の型番のみ。型番が無く略称のみの場合（例: SSK）は空文字にする",
      "型式": "品名・規格欄の下段にある日本語の品名（例: 月例点検費、指導員交通費）",
      "管理No": "管理No.の値",
      "数量": "数量（数値のみ）",
      "単位": "単位（例: 回、台、式）",
      "単価": "単価（数値のみ）",
      "金額": "金額（数値のみ）",
      "基本管理料": "基本管理料の値（数値のみ）",
      "備考": "備考欄の値"
    }
  ]
}

注意事項:
- 金額・数量・単価・基本管理料は数値のみ（カンマや円記号は除去）
- 日付はYYYY-MM-DD形式に統一
- 明細テーブルの全行を抽出すること
- 品名・規格欄に機種名（例: SGT JCL-015C）と品名（例: 月例点検費）が分かれている場合、機種名は"機種"に、品名は"型式"に分けて格納
- 該当情報がない場合は空文字""を設定`;

function getPromptFromSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(PROMPT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROMPT_SHEET_NAME);
    sheet.getRange('A1').setValue(DEFAULT_PROMPT);
    sheet.setColumnWidth(1, 800);
  }
  const val = sheet.getRange('A1').getValue();
  if (!val) {
    throw new Error('プロンプトシートのA1セルが空です。プロンプトを入力してください。');
  }
  return String(val);
}

// ============================================================
// ルーティング
// ============================================================

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('dashboard')
      .setTitle('納品書管理 - ダッシュボード')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('納品書管理 - データ管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ============================================================
// PDF処理（Drive フォルダから一括取込）
// ============================================================

// GAS Web アプリの実行時間上限（6分）に対する安全マージン
const SCAN_MAX_RUNTIME_MS = 5 * 60 * 1000;

// 未処理フォルダのURLを返す（Webアプリから「開く」ボタンで使用）
function getInboxFolderUrl() {
  const inboxId = getProp('INBOX_FOLDER_ID');
  if (!inboxId) return { success: false, error: 'INBOX_FOLDER_ID が Script Properties に設定されていません' };
  return { success: true, url: 'https://drive.google.com/drive/folders/' + inboxId };
}

function processInboxFolder() {
  const inboxId = getProp('INBOX_FOLDER_ID');
  const processedId = getProp('PROCESSED_FOLDER_ID');
  if (!inboxId) return { success: false, error: 'Script Properties に INBOX_FOLDER_ID が設定されていません' };
  if (!processedId) return { success: false, error: 'Script Properties に PROCESSED_FOLDER_ID が設定されていません' };

  let inbox, processed;
  try { inbox = DriveApp.getFolderById(inboxId); }
  catch (e) { return { success: false, error: 'INBOX_FOLDER_ID が無効です: ' + e.message }; }
  try { processed = DriveApp.getFolderById(processedId); }
  catch (e) { return { success: false, error: 'PROCESSED_FOLDER_ID が無効です: ' + e.message }; }

  const files = inbox.getFilesByType(MimeType.PDF);
  const result = {
    success: true,
    processed: 0,
    duplicateSkipped: 0,
    pendingConflicts: [],
    errors: [],
    timedOut: false
  };

  const start = Date.now();
  while (files.hasNext()) {
    if (Date.now() - start > SCAN_MAX_RUNTIME_MS) {
      result.timedOut = true;
      console.log('実行時間上限により中断');
      break;
    }
    const file = files.next();
    const fileId = file.getId();
    const fileName = file.getName();
    try {
      console.log('処理開始: ' + fileName);
      const base64 = Utilities.base64Encode(file.getBlob().getBytes());
      const geminiResult = callGeminiApi(base64, 'application/pdf');
      const records = normalizeGeminiResponse(geminiResult, fileId, fileName);
      const deliveryNo = records.length > 0 ? String(records[0]['納品書No'] || '').trim() : '';

      if (deliveryNo) {
        const existing = findByDeliveryNo(deliveryNo);
        if (existing && existing.rows.length > 0) {
          if (isSameDeliveryContent(existing.rows, records)) {
            file.moveTo(processed);
            result.duplicateSkipped++;
            console.log('完全重複 → 処理済へ移動: ' + fileName);
            continue;
          }
          // 差分あり → 取込はしないが警告を出したので PDF は処理済へ移動
          file.moveTo(processed);
          result.pendingConflicts.push({
            pendingFileId: fileId,
            pendingFileName: fileName,
            pendingRecordsJson: JSON.stringify(records),
            existingFileIds: existing.fileIds,
            existingFileName: existing.fileName,
            existingFileCount: existing.rows.length,
            existingNo: deliveryNo
          });
          console.log('差分検出 → 処理済へ移動 (要確認): ' + fileName + ' (No=' + deliveryNo + ')');
          continue;
        }
      }

      appendRecordsToSheet(records);
      file.moveTo(processed);
      result.processed++;
      console.log('新規登録 → 処理済へ移動: ' + fileName);
    } catch (err) {
      console.error('processInboxFolder(' + fileName + ') エラー: ' + err.message + '\n' + err.stack);
      result.errors.push({ fileName: fileName, error: err.message });
    }
  }
  return result;
}

// レコード群をシートに追記（実シートのヘッダー順に整列）
function appendRecordsToSheet(records) {
  const sheet = getSheet();
  const cmap = getColumnMap(sheet);
  const lastCol = sheet.getLastColumn();
  const rows = records.map(function(rec) {
    const row = new Array(lastCol).fill('');
    Object.keys(rec).forEach(function(k) {
      const idx = cmap[k];
      if (idx !== undefined) row[idx] = rec[k];
    });
    return row;
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, lastCol).setValues(rows);
  }
  return rows.length;
}

// 納品書No が一致する既存レコードを全て取得
function findByDeliveryNo(noValue) {
  if (!noValue) return null;
  const target = String(noValue).trim();
  if (!target) return null;
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const cmap = getColumnMap(sheet);
  const noCol = cmap['納品書No'];
  const fileIdCol = cmap['file_id'];
  const fileNameCol = cmap['file_name'];
  if (noCol === undefined) return null;
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];
  const fileIds = [];
  const fileIdSet = {};
  let firstFileName = '';
  for (let i = 0; i < values.length; i++) {
    const existingNo = String(values[i][noCol] || '').trim();
    if (existingNo && existingNo === target) {
      const rec = {};
      Object.keys(cmap).forEach(function(k) { rec[k] = values[i][cmap[k]]; });
      rows.push(rec);
      const fid = fileIdCol !== undefined ? String(values[i][fileIdCol] || '') : '';
      if (fid && !fileIdSet[fid]) {
        fileIdSet[fid] = true;
        fileIds.push(fid);
        if (!firstFileName && fileNameCol !== undefined) {
          firstFileName = String(values[i][fileNameCol] || '');
        }
      }
    }
  }
  if (rows.length === 0) return null;
  return { rows: rows, fileIds: fileIds, fileName: firstFileName };
}

// 既存シート行と新規レコードが全データフィールドで一致するか
function isSameDeliveryContent(existingRows, newRecords) {
  if (!existingRows || !newRecords) return false;
  if (existingRows.length !== newRecords.length) return false;
  const compareFields = [
    '日時', '納品書No', '取引先', '作業所', '貸出期間', '請求計上日', '伝票摘要',
    '区分', '機種', '号機', '型式', '管理No', '数量', '単位', '単価', '金額', '基本管理料', '備考'
  ];
  const byOrder = function(a, b) {
    return Number(a['item_order'] || 0) - Number(b['item_order'] || 0);
  };
  const sortedA = existingRows.slice().sort(byOrder);
  const sortedB = newRecords.slice().sort(byOrder);
  for (let i = 0; i < sortedA.length; i++) {
    for (let j = 0; j < compareFields.length; j++) {
      const key = compareFields[j];
      const a = String(sortedA[i][key] == null ? '' : sortedA[i][key]).trim();
      const b = String(sortedB[i][key] == null ? '' : sortedB[i][key]).trim();
      if (a !== b) return false;
    }
  }
  return true;
}

// 差分ありの pending 群をまとめて上書き（フロントの「選択分を上書き」で呼ばれる）
function bulkOverwriteByNo(conflictsJson) {
  try {
    const conflicts = JSON.parse(conflictsJson) || [];
    const processedId = getProp('PROCESSED_FOLDER_ID');
    let processed = null;
    if (processedId) {
      try { processed = DriveApp.getFolderById(processedId); }
      catch (e) { console.error('PROCESSED_FOLDER_ID 取得失敗: ' + e.message); }
    }

    let overwritten = 0;
    const errors = [];
    conflicts.forEach(function(c) {
      try {
        const records = JSON.parse(c.pendingRecordsJson);
        records.forEach(function(rec) {
          rec['file_id'] = c.pendingFileId;
          rec['file_name'] = c.pendingFileName;
        });

        deleteRowsByFileIds(c.existingFileIds || []);
        (c.existingFileIds || []).forEach(function(fid) {
          try { DriveApp.getFileById(fid).setTrashed(true); }
          catch (e) { console.error('旧PDF削除失敗: ' + fid + ' - ' + e.message); }
        });

        appendRecordsToSheet(records);

        if (processed) {
          try { DriveApp.getFileById(c.pendingFileId).moveTo(processed); }
          catch (e) { console.error('処理済へ移動失敗: ' + e.message); }
        }

        overwritten++;
      } catch (err) {
        console.error('bulkOverwriteByNo('+ c.pendingFileName + ') エラー: ' + err.message);
        errors.push({ fileName: c.pendingFileName, error: err.message });
      }
    });
    return { success: true, overwritten: overwritten, errors: errors };
  } catch (err) {
    console.error('bulkOverwriteByNo エラー: ' + err.message + '\n' + err.stack);
    return { success: false, error: err.message };
  }
}

// 指定 file_id 群に紐づくシート行を全て削除
function deleteRowsByFileIds(fileIds) {
  if (!fileIds || fileIds.length === 0) return 0;
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;
  const cmap = getColumnMap(sheet);
  const fileIdCol = cmap['file_id'];
  if (fileIdCol === undefined) return 0;
  const data = sheet.getRange(2, fileIdCol + 1, lastRow - 1, 1).getValues();
  const idSet = {};
  fileIds.forEach(function(id) { idSet[String(id)] = true; });
  const rowsToDelete = [];
  for (let r = 0; r < data.length; r++) {
    if (idSet[String(data[r][0])]) {
      rowsToDelete.push(r + 2);
    }
  }
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowNum) { sheet.deleteRow(rowNum); });
  return rowsToDelete.length;
}

function callGeminiApi(base64Data, mimeType) {
  const apiKey = getProp('GEMINI_API_KEY');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  const prompt = getPromptFromSheet();

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        {
          inline_data: {
            mime_type: mimeType,
            data: base64Data
          }
        }
      ]
    }],
    generationConfig: {
      response_mime_type: 'application/json',
      temperature: 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error('Gemini API Error: ' + result.error.message);
  }

  const text = result.candidates[0].content.parts[0].text;
  return JSON.parse(text);
}

function normalizeGeminiResponse(data, fileId, fileName) {
  const records = [];
  const h = data.header || {};
  const items = data.items || [];
  const now = formatDate(new Date());

  // ファイル名から取引先・機種を抽出
  const clientName = extractClientFromFileName(fileName);
  const machineName = extractMachineFromFileName(fileName);

  const headerFields = {
    '日時': h['日時'] || '',
    '納品書No': h['納品書No'] || '',
    '取引先': clientName,
    '作業所': h['作業所'] || '',
    '貸出期間': h['貸出期間'] || '',
    '請求計上日': h['請求計上日'] || '',
    '伝票摘要': h['伝票摘要'] || ''
  };

  const baseRecord = function(idx) {
    const rec = {
      'row_key': fileId + '_' + idx,
      'file_id': fileId,
      'file_name': fileName,
      'status': '未処理',
      'processed_at': now,
      'item_order': idx
    };
    Object.keys(headerFields).forEach(function(k) { rec[k] = headerFields[k]; });
    return rec;
  };

  if (items.length === 0) {
    const rec = baseRecord(0);
    ['区分', '機種', '号機', '型式', '管理No', '数量', '単位', '単価', '金額', '基本管理料', '備考']
      .forEach(function(k) { rec[k] = ''; });
    records.push(rec);
  } else {
    items.forEach(function(item, idx) {
      const rec = baseRecord(idx);
      rec['区分'] = item['区分'] || '';
      rec['機種'] = machineName;
      rec['号機'] = ''; // 手動入力用
      rec['型式'] = item['型式'] || '';
      rec['管理No'] = item['管理No'] || '';
      rec['数量'] = safeParseNumber(item['数量']);
      rec['単位'] = item['単位'] || '';
      rec['単価'] = safeParseNumber(item['単価']);
      rec['金額'] = safeParseNumber(item['金額']);
      rec['基本管理料'] = safeParseNumber(item['基本管理料']);
      rec['備考'] = item['備考'] || '';
      records.push(rec);
    });
  }
  return records;
}

// ============================================================
// CRUD操作
// ============================================================

function getDataFromSpreadsheet() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const cmap = getColumnMap(sheet);
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const get = function(row, name) {
    const i = cmap[name];
    return i === undefined ? '' : String(row[i] || '');
  };
  const grouped = {};

  data.forEach(function(row) {
    const fileId = get(row, 'file_id');
    if (!fileId) return;
    if (!grouped[fileId]) {
      grouped[fileId] = {
        file_id: fileId,
        file_name: get(row, 'file_name'),
        status: get(row, 'status'),
        日時: get(row, '日時'),
        納品書No: get(row, '納品書No'),
        取引先: get(row, '取引先'),
        作業所: get(row, '作業所'),
        貸出期間: get(row, '貸出期間'),
        請求計上日: get(row, '請求計上日'),
        伝票摘要: get(row, '伝票摘要'),
        items: []
      };
    }
    grouped[fileId].items.push({
      row_key: get(row, 'row_key'),
      区分: get(row, '区分'),
      機種: get(row, '機種'),
      号機: get(row, '号機'),
      型式: get(row, '型式'),
      管理No: get(row, '管理No'),
      数量: get(row, '数量'),
      単位: get(row, '単位'),
      単価: get(row, '単価'),
      金額: get(row, '金額'),
      基本管理料: get(row, '基本管理料'),
      備考: get(row, '備考'),
      processed_at: get(row, 'processed_at'),
      item_order: get(row, 'item_order')
    });
  });

  // 配列に変換しソート
  return Object.values(grouped).sort(function(a, b) {
    return String(b['日時']).localeCompare(String(a['日時']));
  });
}

function updateOrderData(updates) {
  // updates: [{row_key, field, value}, ...]
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, updated: 0 };

  const cmap = getColumnMap(sheet);
  const rowKeyCol = cmap['row_key'];
  if (rowKeyCol === undefined) return { success: false, updated: 0, error: 'row_key列が見つかりません' };
  const data = sheet.getRange(2, rowKeyCol + 1, lastRow - 1, 1).getValues();

  let updated = 0;
  updates.forEach(function(u) {
    for (let r = 0; r < data.length; r++) {
      if (String(data[r][0]) === String(u.row_key)) {
        const col = cmap[u.field];
        if (col !== undefined) {
          sheet.getRange(r + 2, col + 1).setValue(u.value);
          updated++;
        }
        break;
      }
    }
  });

  return { success: true, updated: updated };
}

function deleteOrderRows(keys) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, deleted: 0 };

  const cmap = getColumnMap(sheet);
  const rowKeyCol = cmap['row_key'];
  if (rowKeyCol === undefined) return { success: false, deleted: 0, error: 'row_key列が見つかりません' };
  const data = sheet.getRange(2, rowKeyCol + 1, lastRow - 1, 1).getValues();
  const rowsToDelete = [];

  keys.forEach(function(key) {
    for (let r = 0; r < data.length; r++) {
      if (String(data[r][0]) === String(key)) {
        rowsToDelete.push(r + 2); // シート行番号
      }
    }
  });

  // 下から削除
  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });

  return { success: true, deleted: rowsToDelete.length };
}

function getFilePreviewUrl(fileId) {
  return 'https://drive.google.com/file/d/' + fileId + '/preview';
}

// ============================================================
// ダッシュボードデータ
// ============================================================

function getDashboardData(startDate, endDate) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { monthly: [], clients: [], machines: [] };

  const cmap = getColumnMap(sheet);
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const start = startDate ? new Date(startDate) : null;
  const end = endDate ? new Date(endDate) : null;

  const colDate = cmap['日時'];
  const colAmount = cmap['金額'];
  const colClient = cmap['取引先'];
  const colMachine = cmap['機種'];

  const monthly = {};
  const clients = {};
  const machines = {};

  data.forEach(function(row) {
    const dateStr = String(colDate !== undefined ? row[colDate] : '');
    const amount = safeParseNumber(colAmount !== undefined ? row[colAmount] : 0);
    const client = String(colClient !== undefined ? row[colClient] : '');
    const machine = String(colMachine !== undefined ? row[colMachine] : '');

    // 日付フィルタ
    if (dateStr) {
      const d = new Date(dateStr);
      if (start && d < start) return;
      if (end && d > end) return;

      // 月別集計
      const monthKey = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
      monthly[monthKey] = (monthly[monthKey] || 0) + amount;
    }

    // 取引先別
    if (client) {
      clients[client] = (clients[client] || 0) + amount;
    }

    // 機種別
    if (machine) {
      machines[machine] = (machines[machine] || 0) + amount;
    }
  });

  // ソートして返却
  const monthlyArr = Object.keys(monthly).sort().map(function(k) {
    return { month: k, amount: monthly[k] };
  });
  const clientsArr = Object.keys(clients).map(function(k) {
    return { name: k, amount: clients[k] };
  }).sort(function(a, b) { return b.amount - a.amount; });
  const machinesArr = Object.keys(machines).map(function(k) {
    return { name: k, amount: machines[k] };
  }).sort(function(a, b) { return b.amount - a.amount; });

  return { monthly: monthlyArr, clients: clientsArr, machines: machinesArr };
}

// ============================================================
// Excel出力
// ============================================================

function exportToExcel() {
  const ss = getSpreadsheet();
  const ssId = ss.getId();
  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });
  const blob = response.getBlob().setName(ss.getName() + '.xlsx');
  const file = DriveApp.createFile(blob);
  const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + file.getId();
  return { success: true, url: downloadUrl, fileName: blob.getName() };
}

// ============================================================
// ユーティリティ
// ============================================================

function extractClientFromFileName(fileName) {
  // 拡張子を除去
  const name = fileName.replace(/\.pdf$/i, '');
  // _ (半角), ＿ (全角), ‗ (U+2017) で分割し先頭を取得
  const parts = name.split(/[_＿‗]/);
  return parts[0] || '';
}

function extractMachineFromFileName(fileName) {
  // 拡張子を除去
  const name = fileName.replace(/\.pdf$/i, '');
  // _ (半角), ＿ (全角), ‗ (U+2017) で分割
  const parts = name.split(/[_＿‗]/);
  // 右から見て最初の_と次の_の間 = 後ろから2番目の要素
  if (parts.length >= 3) {
    return parts[parts.length - 2] || '';
  }
  return '';
}

function safeParseNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  const str = String(val).replace(/[,¥￥円\s]/g, '');
  const num = Number(str);
  return isNaN(num) ? 0 : num;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
