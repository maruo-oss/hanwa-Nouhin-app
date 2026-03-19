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
// PDF処理
// ============================================================

function uploadAndProcessPDF(base64, fileName) {
  try {
    // Drive に保存
    console.log('PDF保存開始: ' + fileName);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64),
      'application/pdf',
      fileName
    );
    const rootFolderId = getProp('ROOT_FOLDER_ID');
    let file;
    if (rootFolderId) {
      const folder = DriveApp.getFolderById(rootFolderId);
      file = folder.createFile(blob);
    } else {
      file = DriveApp.createFile(blob);
    }
    const fileId = file.getId();
    console.log('Drive保存完了: ' + fileId);

    // Gemini で解析
    console.log('Gemini API呼び出し開始');
    const geminiResult = callGeminiApi(base64, 'application/pdf');
    console.log('Gemini API完了: ' + JSON.stringify(geminiResult).substring(0, 500));

    const rows = normalizeGeminiResponse(geminiResult, fileId, fileName);
    console.log('正規化完了: ' + rows.length + '行, 列数: ' + (rows[0] ? rows[0].length : 0));

    // スプレッドシートに保存
    const sheet = getSheet();
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
        .setValues(rows);
    }
    console.log('シート書き込み完了');

    return { success: true, fileId: fileId, rowCount: rows.length };
  } catch (err) {
    console.error('uploadAndProcessPDF エラー: ' + err.message + '\n' + err.stack);
    return { success: false, error: err.message };
  }
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
  const rows = [];
  const h = data.header || {};
  const items = data.items || [];
  const now = formatDate(new Date());

  // 取引先はファイル名の先頭（_ or ‗ の手前）から抽出
  const clientName = extractClientFromFileName(fileName);

  const headerPart = [
    h['日時'] || '', h['納品書No'] || '',
    clientName, h['作業所'] || '',
    h['貸出期間'] || '', h['請求計上日'] || '', h['伝票摘要'] || ''
  ];

  if (items.length === 0) {
    const rowKey = fileId + '_0';
    rows.push([
      rowKey, fileId, fileName, '完了',
      ...headerPart,
      '', '', '', '', '', '', '', '', '', '', '',
      now, 0
    ]);
  } else {
    items.forEach(function(item, idx) {
      const rowKey = fileId + '_' + idx;
      rows.push([
        rowKey, fileId, fileName, '完了',
        ...headerPart,
        item['区分'] || '', item['機種'] || '',
        '', // 号機（手動入力用、Geminiでは空）
        item['型式'] || '', item['管理No'] || '',
        safeParseNumber(item['数量']),
        item['単位'] || '',
        safeParseNumber(item['単価']),
        safeParseNumber(item['金額']),
        safeParseNumber(item['基本管理料']),
        item['備考'] || '',
        now, idx
      ]);
    });
  }
  return rows;
}

// ============================================================
// CRUD操作
// ============================================================

function getDataFromSpreadsheet() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getDisplayValues();
  const grouped = {};

  data.forEach(function(row) {
    const fileId = String(row[1]);
    if (!fileId) return;
    if (!grouped[fileId]) {
      grouped[fileId] = {
        file_id: fileId,
        file_name: String(row[2]),
        status: String(row[3]),
        日時: String(row[4]),
        納品書No: String(row[5]),
        取引先: String(row[6]),
        作業所: String(row[7]),
        貸出期間: String(row[8]),
        請求計上日: String(row[9]),
        伝票摘要: String(row[10]),
        items: []
      };
    }
    grouped[fileId].items.push({
      row_key: String(row[0]),
      区分: String(row[11]),
      機種: String(row[12]),
      号機: String(row[13]),
      型式: String(row[14]),
      管理No: String(row[15]),
      数量: String(row[16]),
      単位: String(row[17]),
      単価: String(row[18]),
      金額: String(row[19]),
      基本管理料: String(row[20]),
      備考: String(row[21]),
      processed_at: String(row[22]),
      item_order: String(row[23])
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

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const fieldIndex = {};
  HEADERS.forEach(function(h, i) { fieldIndex[h] = i; });

  let updated = 0;
  updates.forEach(function(u) {
    for (let r = 0; r < data.length; r++) {
      if (data[r][0] === u.row_key) {
        const col = fieldIndex[u.field];
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

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const rowsToDelete = [];

  keys.forEach(function(key) {
    for (let r = 0; r < data.length; r++) {
      if (data[r][0] === key) {
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

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const start = startDate ? new Date(startDate) : null;
  const end = endDate ? new Date(endDate) : null;

  const monthly = {};
  const clients = {};
  const machines = {};

  data.forEach(function(row) {
    const dateStr = String(row[4]);
    const amount = safeParseNumber(row[19]); // 金額列
    const client = String(row[6]);           // 取引先列
    const machine = String(row[12]);         // 機種列（明細行ごと）

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

function safeParseNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  const str = String(val).replace(/[,¥￥円\s]/g, '');
  const num = Number(str);
  return isNaN(num) ? 0 : num;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
