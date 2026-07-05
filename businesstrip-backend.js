/**
 * 詠雋稅務記帳士事務所 — 出差旅費報告表後端
 * Google Apps Script Web App
 *
 * Sheet ID: 1sUXWmZfAblSfh8imndFnipBd1j50rL1ELa5tmTHCgAs
 * 印鑑 Drive 資料夾 ID: 1hf-qQju9x4gMs8dF5jl8A2uwhYZdm-Io
 *
 * 接收 POST JSON:
 * {
 *   type: "domestic" | "intl",
 *   common: { ... 共同欄位 },
 *   rows: [ { ... 明細列 }, ... ],
 *   stamps: { key: { filename, base64, mimeType }, ... }
 * }
 *
 * domestic stamps keys: manager(主管印鑑), applicant(申請人印鑑)
 * intl     stamps keys: manager(部門主管), traveler(出差人)
 */

const SHEET_ID = '1sUXWmZfAblSfh8imndFnipBd1j50rL1ELa5tmTHCgAs';
const STAMP_FOLDER_ID = '1hf-qQju9x4gMs8dF5jl8A2uwhYZdm-Io';

const SHEET_DOMESTIC = '國內出差記錄';
const SHEET_INTL = '國外出差記錄';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const type = payload.type;

    if (type !== 'domestic' && type !== 'intl') {
      return jsonResponse({ ok: false, error: 'type 必須是 domestic 或 intl' });
    }

    // 1. 印鑑圖片存 Drive，取得連結
    const stampLinks = saveStamps(payload.stamps || {}, payload.common || {}, type);

    // 2. 寫入對應分頁
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheetName = type === 'domestic' ? SHEET_DOMESTIC : SHEET_INTL;
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return jsonResponse({ ok: false, error: '找不到分頁：' + sheetName });
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    if (type === 'domestic') {
      writeDomestic(sheet, now, payload.common, payload.rows || [], stampLinks);
    } else {
      writeIntl(sheet, now, payload.common, payload.rows || [], stampLinks, payload.extra || {});
    }

    return jsonResponse({ ok: true, stampLinks: stampLinks });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/**
 * 國內出差記錄 欄位順序（1008 新版型定案）：
 * A 送出時間 | B 公司名稱 | C 統一編號 | D 填報人 | E 職務 | F 拜訪對象
 * G 出差期間(起) | H 出差期間(迄) | I 天數 | J 合計金額 | K 主管印鑑連結 | L 申請人印鑑連結
 * M 日期 | N 地點 | O 摘要 | P 交通費 | Q 住宿費 | R 膳什費 | S 停車費 | T 其他 | U 小計
 *
 * 共同欄位只寫第一列，明細每筆一列。
 */
function writeDomestic(sheet, now, common, rows, stampLinks) {
  const output = [];
  const detailRows = rows.length > 0 ? rows : [{}];

  detailRows.forEach(function (r, i) {
    const isFirst = i === 0;
    output.push([
      isFirst ? now : '',
      isFirst ? (common.companyName || '') : '',
      isFirst ? (common.taxId || '') : '',
      isFirst ? (common.filler || '') : '',
      isFirst ? (common.title || '') : '',
      isFirst ? (common.visitTarget || '') : '',
      isFirst ? (common.dateStart || '') : '',
      isFirst ? (common.dateEnd || '') : '',
      isFirst ? (common.days || '') : '',
      isFirst ? (common.total || '') : '',
      isFirst ? (stampLinks.manager || '') : '',
      isFirst ? (stampLinks.applicant || '') : '',
      r.date || '',
      r.city || '',
      r.summary || '',
      num(r.transport),
      num(r.lodging),
      num(r.meals),
      num(r.parking),
      num(r.other),
      num(r.subtotal)
    ]);
  });

  appendRows(sheet, output);
}

/**
 * 國外出差記錄 欄位順序（Ragic 式兩表版）：
 * 共同：A 送出時間 | B 公司名稱 | C 統一編號 | D 出差人 | E 職稱
 *       F 出差期間(起) | G 出差期間(迄) | H 天數 | I 主管印鑑連結 | J 出差人簽名連結
 *       K 新台幣合計 | L 外幣合計(原幣)
 * 明細：M 日期 | N 出發地 | O 到達地 | P 交通費幣別 | Q 交通費金額 | R 住宿費幣別 | S 住宿費金額
 *       T 膳什費幣別 | U 膳什費金額 | V 辦公費幣別 | W 辦公費金額
 * 憑證：X 序號 | Y 報帳憑證種類 | Z 張數 | AA 備註 | AB 機票幣別 | AC 機票金額
 *       AD 預支幣別 | AE 預支金額 | AF 公司卡幣別 | AG 公司卡金額
 *
 * 共同欄位只寫第一列；明細與憑證列並排（列數取兩者較大值）。
 */
function writeIntl(sheet, now, common, rows, stampLinks, extra) {
  const output = [];
  const x = extra || {};
  const vouchers = x.vouchers || [];
  const n = Math.max(rows.length, vouchers.length, 1);

  for (var i = 0; i < n; i++) {
    const isFirst = i === 0;
    const r = rows[i] || {};
    const vc = vouchers[i] || {};
    output.push([
      isFirst ? now : '',
      isFirst ? (common.companyName || '') : '',
      isFirst ? (common.taxId || '') : '',
      isFirst ? (common.filler || '') : '',
      isFirst ? (common.title || '') : '',
      isFirst ? (common.dateStart || '') : '',
      isFirst ? (common.dateEnd || '') : '',
      isFirst ? (common.days || '') : '',
      isFirst ? (stampLinks.manager || '') : '',
      isFirst ? (stampLinks.traveler || '') : '',
      isFirst ? num(x.ntdTotal) : '',
      isFirst ? num(x.fxTotal) : '',
      r.date || '',
      r.from || '',
      r.dest || '',
      r.transportCur || '',
      num(r.transportAmt),
      r.lodgingCur || '',
      num(r.lodgingAmt),
      r.mealsCur || '',
      num(r.mealsAmt),
      r.officeCur || '',
      num(r.officeAmt),
      vc.serial || '',
      vc.kind || '',
      num(vc.count),
      vc.note || '',
      vc.ticketCur || '',
      num(vc.ticketAmt),
      vc.advanceCur || '',
      num(vc.advanceAmt),
      vc.cardCur || '',
      num(vc.cardAmt)
    ]);
  }

  appendRows(sheet, output);
}

/** 印鑑圖片 base64 → Drive，檔名：日期_填報人_用途.副檔名，回傳 {key: url} */
function saveStamps(stamps, common, type) {
  const links = {};
  const keys = Object.keys(stamps);
  if (keys.length === 0) return links;

  const folder = DriveApp.getFolderById(STAMP_FOLDER_ID);
  const dateTag = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
  const company = (common.companyName || '').replace(/[\\/:*?"<>|]/g, '');
  const filler = ((company ? company + '_' : '') + (common.filler || '未填名')).replace(/[\\/:*?"<>|]/g, '');

  const labelMap = {
    manager: '主管印鑑',
    applicant: '申請人印鑑',
    traveler: '出差人簽名'
  };

  keys.forEach(function (key) {
    const s = stamps[key];
    if (!s || !s.base64) return;
    const mimeType = s.mimeType || 'image/png';
    const ext = mimeType.indexOf('jpeg') > -1 ? 'jpg' : 'png';
    const label = labelMap[key] || key;
    const filename = dateTag + '_' + filler + '_' + label + '.' + ext;

    const blob = Utilities.newBlob(
      Utilities.base64Decode(s.base64),
      mimeType,
      filename
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    links[key] = file.getUrl();
  });

  return links;
}

function appendRows(sheet, output) {
  if (output.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, output.length, output[0].length).setValues(output);
}

function num(v) {
  if (v === undefined || v === null || v === '') return '';
  const n = Number(v);
  return isNaN(n) ? v : n;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 測試用：GET 回傳狀態 */
function doGet() {
  return jsonResponse({ ok: true, service: 'businesstrip-backend', time: new Date().toISOString() });
}
