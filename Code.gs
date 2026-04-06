const CONFIG = {
  SHEET_NAME: '母材キャンセル記録',
  MONTHLY_SUMMARY_SHEET: '集計_月別',
  MATERIAL_SUMMARY_SHEET: '集計_材料別',
  CUSTOMER_SUMMARY_SHEET: '集計_客先別',
  SAVE_DEBUG_IMAGE: false,
  DEBUG_FOLDER_NAME: 'MaterialCancelDebugImages',
  GEMINI_MODEL: 'gemini-2.5-flash'
};

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput({ ok: false, error: 'POSTデータがありません' });
    }

    const req = JSON.parse(e.postData.contents);
    const action = String(req.action || '').trim();

    if (action === 'analyze') {
      const image = req.image;
      if (!image) {
        return jsonOutput({ ok: false, error: 'image がありません' });
      }

      const result = processImage(image);
      return jsonOutput({ ok: true, items: result });
    }

    if (action === 'save') {
      const list = req.list;
      const sender = String(req.sender || '').trim();

      if (!Array.isArray(list)) {
        return jsonOutput({ ok: false, error: 'list は配列で送ってください' });
      }

      if (!sender) {
        return jsonOutput({ ok: false, error: '送信者名がありません' });
      }

      const saved = saveData(list, sender);
      updateSummarySheets_();

      if (saved === 0) {
        return jsonOutput({
          ok: true,
          savedRows: 0,
          message: 'キャンセル枚数が1以上の行がなかったため、保存はありませんでした'
        });
      }

      return jsonOutput({
        ok: true,
        savedRows: saved,
        message: saved + '件保存しました'
      });
    }

    if (action === 'summary') {
      const summary = getSummaryData();
      return jsonOutput({ ok: true, summary: summary });
    }

    return jsonOutput({ ok: false, error: 'action が不正です' });
  } catch (err) {
    return jsonOutput({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function processImage(base64Image) {
  const apiKey = getGeminiApiKey_();
  const mimeType = detectMimeType_(base64Image);
  const base64Data = extractBase64_(base64Image);

  if (CONFIG.SAVE_DEBUG_IMAGE) {
    saveDebugImage_(base64Data, mimeType);
  }

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(CONFIG.GEMINI_MODEL) +
    ':generateContent';

  const promptText = [
    '画像は工場の材料伝票です。',
    '画像の向きは一定ではありません。縦向き、横向き、回転があっても対応してください。',
    '位置で決めず、画像全体を見て内容から判断してください。',
    '',
    '伝票内から各母材行を抽出してください。',
    '各行について以下の項目を返してください。',
    '',
    '- customer: 伝票全体の中から読み取れる客先名。位置は固定しない。母材明細とは別に単独で記載されている会社名・客先名を優先する。不明なら空文字。',
    '- material: 材質',
    '- thickness: 板厚',
    '- size: サイズ（例 914x1829）',
    '- planned_qty: 使用予定枚数',
    '',
    '必ずJSON配列のみで返してください。',
    'Markdownや説明文は不要です。',
    '抽出できない場合は [] を返してください。',
    '',
    '出力例:',
    '[',
    '  {',
    '    "customer": "キンキテック",',
    '    "material": "ボンデ",',
    '    "thickness": "1.6",',
    '    "size": "914x1829",',
    '    "planned_qty": 3',
    '  }',
    ']'
  ].join('\n');

  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          { text: promptText },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      response_mime_type: 'application/json',
      temperature: 0.1
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const bodyText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Gemini APIエラー: HTTP ' + statusCode + ' / ' + bodyText);
  }

  const result = JSON.parse(bodyText);

  if (result.error) {
    throw new Error(result.error.message || 'Gemini APIエラー');
  }

  const text = result?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    throw new Error('Geminiの応答が空です');
  }

  let parsed;
  try {
    parsed = JSON.parse(cleanJsonText_(text));
  } catch (err) {
    throw new Error('解析結果のJSON変換に失敗しました: ' + text);
  }

  if (!Array.isArray(parsed)) {
    throw new Error('解析結果が配列ではありません');
  }

  return parsed.map(function(item) {
    return {
      customer: normalizeString_(item.customer),
      material: normalizeString_(item.material),
      thickness: normalizeString_(item.thickness),
      size: normalizeSize_(item.size),
      planned_qty: normalizeInt_(item.planned_qty),
      cancelled_qty: 0
    };
  }).filter(function(item) {
    return item.customer || item.material || item.thickness || item.size || item.planned_qty;
  });
}

function saveData(list, sender) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '日時',
      '日付',
      '月',
      '送信者',
      '客先',
      '材質',
      '板厚',
      'サイズ',
      '使用枚数',
      'キャンセル枚数'
    ]);
  }

  const now = new Date();
  const dateTimeStr = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd HH:mm:ss'
  );

  const dateStr = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd'
  );

  const monthStr = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyyy/MM'
  );

  const rows = list
    .map(function(item) {
      return {
        customer: normalizeString_(item.customer),
        material: normalizeString_(item.material),
        thickness: normalizeString_(item.thickness),
        size: normalizeSize_(item.size),
        planned_qty: normalizeInt_(item.planned_qty),
        cancelled_qty: normalizeInt_(item.cancelled_qty)
      };
    })
    .filter(function(item) {
      return item.customer || item.material || item.thickness || item.size || item.planned_qty;
    })
    .filter(function(item) {
      return item.cancelled_qty > 0;
    })
    .map(function(item) {
      return [
        dateTimeStr,
        dateStr,
        monthStr,
        sender,
        item.customer,
        item.material,
        item.thickness,
        item.size,
        item.planned_qty,
        item.cancelled_qty
      ];
    });

  if (rows.length === 0) {
    return 0;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

function updateSummarySheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);
  const monthlySheet = getOrCreateSheet_(ss, CONFIG.MONTHLY_SUMMARY_SHEET);
  const materialSheet = getOrCreateSheet_(ss, CONFIG.MATERIAL_SUMMARY_SHEET);
  const customerSheet = getOrCreateSheet_(ss, CONFIG.CUSTOMER_SUMMARY_SHEET);

  const values = rawSheet.getDataRange().getValues();
  if (values.length <= 1) {
    monthlySheet.clear();
    materialSheet.clear();
    customerSheet.clear();

    monthlySheet.appendRow(['月', '登録件数', 'キャンセル合計枚数']);
    materialSheet.appendRow(['材料', '登録件数', 'キャンセル合計枚数']);
    customerSheet.appendRow(['客先', '登録件数', 'キャンセル合計枚数']);
    return;
  }

  const rows = values.slice(1);

  const monthlyMap = {};
  const materialMap = {};
  const customerMap = {};

  rows.forEach(function(row) {
    const month = row[2] || '';
    const customer = row[4] || '';
    const material = row[5] || '';
    const thickness = row[6] || '';
    const size = row[7] || '';
    const cancelled = normalizeInt_(row[9]);

    const materialKey = [material, thickness, size].join(' / ');

    if (!monthlyMap[month]) {
      monthlyMap[month] = { count: 0, cancelled: 0 };
    }
    monthlyMap[month].count += 1;
    monthlyMap[month].cancelled += cancelled;

    if (!materialMap[materialKey]) {
      materialMap[materialKey] = { count: 0, cancelled: 0 };
    }
    materialMap[materialKey].count += 1;
    materialMap[materialKey].cancelled += cancelled;

    if (!customerMap[customer]) {
      customerMap[customer] = { count: 0, cancelled: 0 };
    }
    customerMap[customer].count += 1;
    customerMap[customer].cancelled += cancelled;
  });

  const monthlyRows = Object.keys(monthlyMap)
    .sort()
    .map(function(month) {
      return [month, monthlyMap[month].count, monthlyMap[month].cancelled];
    });

  const materialRows = Object.keys(materialMap)
    .sort()
    .map(function(materialKey) {
      return [materialKey, materialMap[materialKey].count, materialMap[materialKey].cancelled];
    });

  const customerRows = Object.keys(customerMap)
    .sort()
    .map(function(customer) {
      return [customer, customerMap[customer].count, customerMap[customer].cancelled];
    });

  monthlySheet.clear();
  materialSheet.clear();
  customerSheet.clear();

  monthlySheet.appendRow(['月', '登録件数', 'キャンセル合計枚数']);
  materialSheet.appendRow(['材料', '登録件数', 'キャンセル合計枚数']);
  customerSheet.appendRow(['客先', '登録件数', 'キャンセル合計枚数']);

  if (monthlyRows.length > 0) {
    monthlySheet.getRange(2, 1, monthlyRows.length, monthlyRows[0].length).setValues(monthlyRows);
  }

  if (materialRows.length > 0) {
    materialSheet.getRange(2, 1, materialRows.length, materialRows[0].length).setValues(materialRows);
  }

  if (customerRows.length > 0) {
    customerSheet.getRange(2, 1, customerRows.length, customerRows[0].length).setValues(customerRows);
  }
}

function getGeminiApiKey_() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('Script Properties に GEMINI_API_KEY が設定されていません');
  }
  return apiKey;
}

function extractBase64_(dataUrlOrBase64) {
  const str = String(dataUrlOrBase64 || '');
  const commaIndex = str.indexOf(',');
  return commaIndex >= 0 ? str.slice(commaIndex + 1) : str;
}

function detectMimeType_(dataUrlOrBase64) {
  const str = String(dataUrlOrBase64 || '');
  const match = str.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,/);
  return match ? match[1] : 'image/jpeg';
}

function cleanJsonText_(text) {
  return String(text || '')
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();
}

function normalizeString_(value) {
  return value == null ? '' : String(value).trim();
}

function normalizeInt_(value) {
  const n = parseInt(String(value == null ? '' : value).replace(/[^\d-]/g, ''), 10);
  return isNaN(n) ? 0 : n;
}

function normalizeSize_(value) {
  return normalizeString_(value)
    .replace(/[×✕]/g, 'x')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function getOrCreateSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  return sheet || ss.insertSheet(sheetName);
}

function saveDebugImage_(base64Data, mimeType) {
  try {
    const folder = getOrCreateFolder_(CONFIG.DEBUG_FOLDER_NAME);
    const bytes = Utilities.base64Decode(base64Data);
    const ext = mimeTypeToExt_(mimeType);
    const fileName =
      'material-slip-' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-SSS') +
      '.' + ext;

    folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));
  } catch (err) {
    Logger.log('saveDebugImage_ error: ' + err);
  }
}

function getOrCreateFolder_(folderName) {
  const iter = DriveApp.getFoldersByName(folderName);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(folderName);
}

function mimeTypeToExt_(mimeType) {
  switch (mimeType) {
    case 'image/png':
      return 'png';
    case 'image/webp':
      return 'webp';
    case 'image/heic':
      return 'heic';
    case 'image/heif':
      return 'heif';
    default:
      return 'jpg';
  }
}

function getSummaryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const monthlySheet = getOrCreateSheet_(ss, CONFIG.MONTHLY_SUMMARY_SHEET);
  const materialSheet = getOrCreateSheet_(ss, CONFIG.MATERIAL_SUMMARY_SHEET);
  const customerSheet = getOrCreateSheet_(ss, CONFIG.CUSTOMER_SUMMARY_SHEET);
  const rawSheet = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);

  return {
    monthly: sheetToObjects_(monthlySheet),
    material: sheetToObjects_(materialSheet),
    customer: sheetToObjects_(customerSheet),
    raw: sheetToObjects_(rawSheet)
  };
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  const rows = values.slice(1);

  return rows.map(function(row) {
    const obj = {};
    headers.forEach(function(header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}
