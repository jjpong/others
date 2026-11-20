/**
 * 臺北市115年度國民中學彈性學習領航課程徵件比賽｜資料整理協作網站（後端）
 * ---------------------------------------------------------------
 * 使用方式（詳見「部署說明」）：
 * 1. 新增一份 Google 試算表 → 擴充功能 → Apps Script
 * 2. 貼上本檔（Code.gs）與 Index.html
 * 3. 修改下方 EDIT_PASSWORD，執行一次 setup()
 * 4. 部署 → 新增部署作業 → 網頁應用程式
 *    執行身分：我；誰可以存取：所有人
 */

// ★ 請改成你和夥伴共用的編輯密碼（只有知道密碼的人能修改內容）
const EDIT_PASSWORD = 'changeme';

const SHEET_FIELDS = 'fields';
const SHEET_FILES = 'files';
const SHEET_LOG = 'log';
const MAX_CELL = 49000;

/* ========== 初始化 ========== */

function setup() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從試算表的「擴充功能 → Apps Script」開啟並執行 setup()');
  props.setProperty('SS_ID', ss.getId());

  ensureSheet_(ss, SHEET_FIELDS, ['key', 'value', 'version', 'updatedAt', 'updatedBy']);
  ensureSheet_(ss, SHEET_FILES, ['id', 'name', 'mime', 'size', 'kind', 'uploadedAt', 'uploadedBy']);
  ensureSheet_(ss, SHEET_LOG, ['time', 'user', 'action', 'detail']);

  let folderId = props.getProperty('FOLDER_ID');
  let folder = null;
  if (folderId) { try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; } }
  if (!folder) {
    const ssFile = DriveApp.getFileById(ss.getId());
    const parents = ssFile.getParents();
    const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    folder = parent.createFolder('領航課程徵件_上傳檔案');
    props.setProperty('FOLDER_ID', folder.getId());
  }
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 3 && def.getLastRow() === 0) ss.deleteSheet(def);

  if (EDIT_PASSWORD === 'changeme') {
    Logger.log('⚠️ 你還沒有修改 EDIT_PASSWORD，建議改掉後再部署。');
  }
  Logger.log('✅ 設定完成。上傳檔案資料夾：' + folder.getUrl());
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('尚未執行 setup()');
  return SpreadsheetApp.openById(id);
}
function folder_() {
  const id = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!id) throw new Error('尚未執行 setup()');
  return DriveApp.getFolderById(id);
}
function checkPw_(pw) {
  if (String(pw || '') !== EDIT_PASSWORD) throw new Error('編輯密碼錯誤');
}
function log_(user, action, detail) {
  try {
    const sh = ss_().getSheetByName(SHEET_LOG);
    sh.appendRow([new Date(), user || '', action, String(detail || '').slice(0, 200)]);
  } catch (e) { /* ignore */ }
}

/* ========== 網頁 ========== */

function doGet() {
  // 把目前資料直接放進網頁，省下一次往返
  const t = HtmlService.createTemplateFromFile('Index');
  let boot = '';
  try { boot = JSON.stringify(api_getAll()).replace(/</g, '\\u003c'); } catch (e) { boot = ''; }
  t.boot = boot;
  return t.evaluate()
    .setTitle('彈性學習領航課程徵件｜資料整理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ========== 給 GitHub Pages 等外部網頁呼叫的接口 ========== */

const API_FNS = {
  api_getAll: api_getAll, api_changes: api_changes, api_checkPassword: api_checkPassword,
  api_save: api_save, api_upload: api_upload, api_getFile: api_getFile, api_getFiles: api_getFiles,
  api_deleteFile: api_deleteFile, api_saveExport: api_saveExport
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API_FNS[req.fn];
    if (!fn) throw new Error('不支援的操作');
    out = { result: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ========== API（前端以 google.script.run 或 doPost 呼叫） ========== */

/** 讀取全部欄位與檔案清單 */
function api_getAll() {
  const ss = ss_();
  const fields = readFields_(ss, 0);
  const files = readFiles_(ss);
  return { fields: fields, files: files, now: Date.now() };
}

/** 讀取 since（毫秒）之後有變動的欄位 */
function api_changes(since) {
  const ss = ss_();
  return { fields: readFields_(ss, Number(since) || 0), now: Date.now() };
}

function readFields_(ss, since) {
  const sh = ss.getSheetByName(SHEET_FIELDS);
  const n = sh.getLastRow() - 1;
  const out = {};
  if (n <= 0) return out;
  const rows = sh.getRange(2, 1, n, 5).getValues();
  rows.forEach(function (r) {
    if (!r[0]) return;
    const at = Number(r[3]) || 0;
    if (at <= since) return;
    out[r[0]] = { v: r[1], ver: Number(r[2]) || 0, at: at, by: r[4] };
  });
  return out;
}

function readFiles_(ss) {
  const sh = ss.getSheetByName(SHEET_FILES);
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  return sh.getRange(2, 1, n, 7).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { id: r[0], name: r[1], mime: r[2], size: r[3], kind: r[4], at: Number(r[5]) || 0, by: r[6] };
    });
}

function api_checkPassword(pw) {
  return String(pw || '') === EDIT_PASSWORD;
}

/**
 * 儲存一個欄位。value 為 JSON 字串。
 * baseVer 與伺服器版本不同時回傳 conflict（除非 force = true）
 */
function api_save(pw, key, value, baseVer, by, force) {
  checkPw_(pw);
  key = String(key);
  value = String(value);
  if (value.length > MAX_CELL) throw new Error('這一欄內容太長（超過 ' + MAX_CELL + ' 字元），請拆成較短的段落。');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = ss_().getSheetByName(SHEET_FIELDS);
    const n = sh.getLastRow() - 1;
    let rowIdx = -1, cur = null;
    if (n > 0) {
      const keys = sh.getRange(2, 1, n, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (keys[i][0] === key) { rowIdx = i + 2; break; }
      }
    }
    if (rowIdx > 0) {
      const r = sh.getRange(rowIdx, 1, 1, 5).getValues()[0];
      cur = { v: r[1], ver: Number(r[2]) || 0, at: Number(r[3]) || 0, by: r[4] };
      if (!force && Number(baseVer || 0) !== cur.ver) {
        return { conflict: true, current: cur };
      }
    }
    const now = Date.now();
    const ver = (cur ? cur.ver : 0) + 1;
    const row = [key, value, ver, now, by || ''];
    if (rowIdx < 0) rowIdx = sh.getLastRow() + 1;
    const rng = sh.getRange(rowIdx, 1, 1, 5);
    rng.setNumberFormat('@');          // 以純文字儲存，避免試算表自動轉型
    rng.setValues([row.map(String)]);
    return { ok: true, ver: ver, at: now, by: by || '' };
  } finally {
    lock.releaseLock();
  }
}

/** 上傳檔案（base64），回傳檔案 id */
function api_upload(pw, meta) {
  checkPw_(pw);
  const bytes = Utilities.base64Decode(meta.b64);
  const blob = Utilities.newBlob(bytes, meta.mime || 'application/octet-stream', meta.name || 'file');
  const file = folder_().createFile(blob);
  const now = Date.now();
  ss_().getSheetByName(SHEET_FILES).appendRow([
    file.getId(), meta.name || '', meta.mime || '', bytes.length, meta.kind || '', now, meta.by || ''
  ]);
  log_(meta.by, 'upload', meta.name);
  return { id: file.getId(), name: meta.name, mime: meta.mime, size: bytes.length, kind: meta.kind || '', at: now, by: meta.by || '' };
}

/** 讀取檔案內容（base64）。只提供本系統上傳過的檔案 */
function api_getFile(id) {
  const ss = ss_();
  const ok = readFiles_(ss).some(function (f) { return f.id === id; });
  if (!ok) throw new Error('找不到檔案');
  const file = DriveApp.getFileById(id);
  const blob = file.getBlob();
  return { id: id, mime: blob.getContentType(), name: file.getName(), b64: Utilities.base64Encode(blob.getBytes()) };
}

/**
 * 一次讀取多個檔案（base64）。小檔（縮圖）會放進快取 6 小時。
 * 只提供本系統上傳過的檔案。
 */
function api_getFiles(ids) {
  ids = (ids || []).slice(0, 20);
  const cache = CacheService.getScriptCache();
  const out = {};
  const cached = cache.getAll(ids.map(function (id) { return 'f_' + id; }));
  const need = [];
  ids.forEach(function (id) {
    const c = cached['f_' + id];
    if (c) { try { out[id] = JSON.parse(c); return; } catch (e) { /* ignore */ } }
    need.push(id);
  });
  if (need.length) {
    const allowed = {};
    readFiles_(ss_()).forEach(function (f) { allowed[f.id] = true; });
    need.forEach(function (id) {
      if (!allowed[id]) return;
      try {
        const file = DriveApp.getFileById(id);
        const blob = file.getBlob();
        const r = { mime: blob.getContentType(), name: file.getName(), b64: Utilities.base64Encode(blob.getBytes()) };
        out[id] = r;
        const str = JSON.stringify(r);
        if (str.length < 95000) { try { cache.put('f_' + id, str, 21600); } catch (e) { /* ignore */ } }
      } catch (e) { /* ignore */ }
    });
  }
  return out;
}

/** 刪除檔案（移到垃圾桶，30 天內可從雲端硬碟救回） */
function api_deleteFile(pw, id, by) {
  checkPw_(pw);
  const sh = ss_().getSheetByName(SHEET_FILES);
  const n = sh.getLastRow() - 1;
  if (n <= 0) return { ok: true };
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === id) {
      sh.deleteRow(i + 2);
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* ignore */ }
      log_(by, 'delete', id);
      break;
    }
  }
  return { ok: true };
}

/** 把輸出的 Word 檔存到雲端資料夾（下載被瀏覽器擋住時的備案） */
function api_saveExport(pw, name, b64, mime) {
  checkPw_(pw);
  let sub;
  const it = folder_().getFoldersByName('輸出檔');
  sub = it.hasNext() ? it.next() : folder_().createFolder('輸出檔');
  const file = sub.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, name));
  return { url: file.getUrl() };
}
